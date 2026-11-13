"""
Step: Build IDxAGE tab from the client's monthly raw data file.

Source: raw survey export (e.g. June-Rawdata.xlsx), sheet "Sheet1"
  - "Internal Respondent Id"  -> IDxAGE column A
  - "あなたの年齢は？"         -> IDxAGE column B AND column C (identical values)
  - "あなたは何歳ですか？"     -> the 16-17 respondent's OWN age; where it holds a number it overrides
                                 the line above (which holds the parent's age for those respondents)

Design notes:
  - Headers are located BY NAME, not by hardcoded column letter, since the raw
    file's column layout is not guaranteed to stay fixed month to month.
  - Age is coerced through a numeric-extraction step (regex-pull the leading
    number) as a defensive measure, even though this month's data is already clean -
    the client has flagged that survey exports sometimes attach label text to
    numeric responses (e.g. "10 - it is highly...").
  - IDxAGE is fully REPLACED each month (rows below the header are cleared before
    writing), not appended - it is a single-month scratch table by design.
"""
import re
import openpyxl
from openpyxl.utils import get_column_letter
from safe_xlsx_writer import patch_workbook

ID_HEADER = "Internal Respondent Id"
AGE_HEADER = "あなたの年齢は？"


def get_sheet_max_row(path, sheet_name):
    """Fast read-only lookup of a sheet's current last used row."""
    wb = openpyxl.load_workbook(path, read_only=True)
    return wb[sheet_name].max_row


def find_header_columns(ws, headers_wanted, search_rows=3):
    """Scan the first `search_rows` rows for exact header text matches.
    Returns {header_text: column_index}. Raises if any header isn't found."""
    found = {}
    for r in range(1, search_rows + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v in headers_wanted and v not in found:
                found[v] = c
        if len(found) == len(headers_wanted):
            break
    missing = set(headers_wanted) - set(found)
    if missing:
        raise ValueError(
            f"Could not locate required header(s) {missing} in the first "
            f"{search_rows} rows of '{ws.title}'. The raw file's layout may "
            f"have changed - check manually before proceeding."
        )
    return found


def clean_age(raw_value):
    """Extract a clean integer age from a raw cell value that may be a plain
    number, a numeric string, or (defensively) a labeled string like
    '10 - it is highly...'. Returns int or None if nothing numeric is found."""
    if raw_value is None:
        return None
    if isinstance(raw_value, (int, float)):
        return int(raw_value)
    match = re.search(r"-?\d+(?:\.\d+)?", str(raw_value))
    if match is None:
        return None
    return int(float(match.group()))


MINOR_AGE_HEADER = "あなたは何歳ですか？"   # JP: the 16-17 respondent's own age (asked only when a parent allowed the child to take part)
MINOR_AGE_RANGE = (10, 17)                   # a value outside this range in the minor column is not treated as a child's age

_WHOLE_NUMBER = re.compile(r"^\s*(\d{1,3})(?:\.0+)?\s*$")


def _norm(v):
    return re.sub(r"\s+", "", v) if isinstance(v, str) else v


def _minor_age(raw):
    """Return the child's age if the cell holds a plain whole number, else None.
    Deliberately stricter than clean_age: in the KO file this column also holds text
    ('...') for respondents who are not 16-17, and those must NOT be read as ages."""
    if isinstance(raw, bool) or raw is None:
        return None
    if isinstance(raw, (int, float)):
        return int(raw) if float(raw).is_integer() else None
    m = _WHOLE_NUMBER.match(str(raw))
    return int(m.group(1)) if m else None


def extract_id_age_pairs(raw_path, sheet_name="Sheet1", age_header=AGE_HEADER, minor_age_header=MINOR_AGE_HEADER):
    """Read (Internal Respondent Id, Age) pairs from the raw monthly file.

    Age rule (confirmed by the client, Oct 2026): the adult age question is the default. For
    respondents aged 16-17 the adult question holds the PARENT's age and the child's own age
    sits in a second age question (minor_age_header). Where that second column holds a real
    number it overrides the first; where it is blank or holds text, the first column is used.

    age_header / minor_age_header default to JP's question text - pass the market's own
    (market_config.py) for KO. Headers are matched by text, ignoring spaces; if both age
    questions share the same text, the first occurrence is the adult age and the second the
    child's."""
    wb = openpyxl.load_workbook(raw_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found in {raw_path}. "
            f"Available sheets: {wb.sheetnames}"
        )
    ws = wb[sheet_name]

    # locate headers (first 3 rows), keeping every column that carries a wanted header
    want = {_norm(ID_HEADER): "id", _norm(age_header): "age"}
    if minor_age_header:
        want.setdefault(_norm(minor_age_header), "minor")
    hits = {}          # normalised header -> [(row, col), ...]
    for r in range(1, 4):
        for c in range(1, ws.max_column + 1):
            v = _norm(ws.cell(row=r, column=c).value)
            if v in want:
                hits.setdefault(v, []).append((r, c))
    for h in (_norm(ID_HEADER), _norm(age_header)):
        if h not in hits:
            raise ValueError(
                f"Could not locate required header {h!r} in the first 3 rows of '{ws.title}'. "
                f"The raw file's layout may have changed - check manually before proceeding."
            )
    header_row, id_col = hits[_norm(ID_HEADER)][0]
    age_col = hits[_norm(age_header)][0][1]
    minor_col = None
    if minor_age_header:
        mh = _norm(minor_age_header)
        cands = [c for _, c in hits.get(mh, []) if c != age_col]
        minor_col = cands[0] if cands else None
        if minor_col is None:
            print(f"  WARNING: the 16-17 respondents' own-age question ({minor_age_header!r}) was NOT found in the raw file. "
                  f"Any 16-17 respondents will be given their PARENT's age. Check the raw file's headers.")

    pairs = []
    dirty_count = 0
    from_minor = 0
    ignored_minor = []
    for r in range(header_row + 1, ws.max_row + 1):
        rid = ws.cell(row=r, column=id_col).value
        if rid is None:
            continue
        raw_age = ws.cell(row=r, column=age_col).value
        age = clean_age(raw_age)
        if minor_col is not None:
            child = _minor_age(ws.cell(row=r, column=minor_col).value)
            if child is not None:
                if MINOR_AGE_RANGE[0] <= child <= MINOR_AGE_RANGE[1]:
                    age = child
                    from_minor += 1
                else:
                    ignored_minor.append((rid, child))
        if age is None:
            print(f"  WARNING: row {r} (ID {rid}) has no parseable age: {raw_age!r}")
            continue
        if not isinstance(raw_age, (int, float)):
            dirty_count += 1
        pairs.append((rid, age))

    print(f"Extracted {len(pairs)} (ID, Age) pairs from '{raw_path}' sheet '{sheet_name}' "
          f"(columns {get_column_letter(id_col)}, {get_column_letter(age_col)}"
          f"{', ' + get_column_letter(minor_col) if minor_col else ''}).")
    if minor_col is not None:
        n_minor = sum(1 for _, a in pairs if a <= 17)
        print(f"  {from_minor} respondents are 16-17: their own age (column {get_column_letter(minor_col)}) replaces the parent's age in column {get_column_letter(age_col)}.")
        if n_minor != from_minor:
            print(f"  NOTE: {n_minor} respondents end up with an age of 17 or under, {from_minor} of them from the own-age column.")
        if ignored_minor:
            print(f"  WARNING: {len(ignored_minor)} value(s) in the own-age column are numbers outside {MINOR_AGE_RANGE[0]}-{MINOR_AGE_RANGE[1]} and were ignored (adult age kept): {ignored_minor[:5]}")
    if dirty_count:
        print(f"  {dirty_count} age values needed numeric cleaning (were text, not plain numbers).")
    return pairs


def make_idxage_edit(pairs, current_max_row, duplicate_column=True):
    """Returns an edit function for use with safe_xlsx_writer.patch_workbook:
    clears the age sheet's existing data rows (2..current_max_row) and
    writes the new pairs into column A (ID) and column B (age). Clearing
    first matters because a shrinking respondent count (fewer rows this
    month than last) would otherwise leave stale rows behind past the new
    data's end.

    duplicate_column: JP's IDxAGE sheet duplicates the age into both
    columns B and C (confirmed: downstream formulas only ever reference
    column B, column C's duplication appears purely historical/cosmetic,
    but is preserved for JP to avoid any risk of an undiscovered
    dependency). KO's Age sheet has only one age column - pass False."""
    def edit(editor):
        editor.clear_rows(2, max(current_max_row, 2))
        for i, (rid, age) in enumerate(pairs, start=2):
            editor.set_cell(f"A{i}", rid)
            editor.set_cell(f"B{i}", age)
            if duplicate_column:
                editor.set_cell(f"C{i}", age)
        print(f"Age sheet: wrote {len(pairs)} rows")
    return edit


def write_idxage(target_wb_path, pairs, sheet_name="IDxAGE", save_path=None, duplicate_column=True):
    """Standalone entry point: patch and save, preserving every other part
    of the workbook byte-for-byte (Power Query connections, pivot caches,
    dynamic-array metadata, etc.) - see safe_xlsx_writer.py for why this
    matters. The monthly pipeline combines this with the other steps into
    one patch_workbook() call instead of calling this separately."""
    current_max_row = get_sheet_max_row(target_wb_path, sheet_name)
    out_path = save_path or target_wb_path
    patch_workbook(target_wb_path, out_path, {
        sheet_name: make_idxage_edit(pairs, current_max_row, duplicate_column)
    })
    print(f"Saved to {out_path}")
    return out_path


def validate_against_respondent_data(target_wb_path, pairs, current_month_filter=None):
    """Sanity check: confirm every current-month respondent ID in Respondent Data
    has a matching age entry, and flag anything in `pairs` that doesn't belong to
    the current month's respondent set.

    Uses read_only mode with sequential iter_rows access, not normal-mode
    load (which costs ~2 minutes on this workbook) or random ws.cell()
    access (which read_only mode handles very slowly)."""
    wb = openpyxl.load_workbook(target_wb_path, read_only=True)
    ws = wb["Respondent Data"]

    header_row = None
    for r in range(1, 12):
        row_vals = next(ws.iter_rows(min_row=r, max_row=r, values_only=True))
        if len(row_vals) > 1 and row_vals[1] == "Resp id":
            header_row = r
            header_vals = row_vals
            break
    if header_row is None:
        raise ValueError("Could not locate 'Resp id' header in Respondent Data")

    month_col = None
    for c, v in enumerate(header_vals, start=1):
        if v == "MM/YY":
            month_col = c
            break

    current_ids = set()
    for row_vals in ws.iter_rows(min_row=header_row + 1, min_col=2, max_col=month_col, values_only=True):
        rid = row_vals[0]
        dt = row_vals[month_col - 2]
        if rid is None:
            continue
        if current_month_filter is None or (
            dt is not None and (dt.year, dt.month) == current_month_filter
        ):
            current_ids.add(rid)

    pair_ids = {rid for rid, _ in pairs}
    missing = current_ids - pair_ids
    extra = pair_ids - current_ids

    print(f"Respondent Data current-month respondents: {len(current_ids)}")
    print(f"Raw file respondents supplied: {len(pair_ids)}")
    if missing:
        print(f"  WARNING: {len(missing)} current-month respondents have NO age in the raw file: "
              f"{sorted(missing)[:10]}{'...' if len(missing) > 10 else ''}")
    if extra:
        print(f"  NOTE: {len(extra)} raw-file respondents are not in the current month's "
              f"Respondent Data (may be normal if raw file is a superset): "
              f"{sorted(extra)[:10]}{'...' if len(extra) > 10 else ''}")
    if not missing and not extra:
        print("  OK: raw file respondents match Respondent Data's current-month set exactly.")
    return missing, extra


if __name__ == "__main__":
    RAW_FILE = "jun-rawdata.xlsx"
    TARGET_FILE = "JP-June.xlsx"
    OUTPUT_FILE = "JP-June-idxage-rebuilt.xlsx"

    pairs = extract_id_age_pairs(RAW_FILE, sheet_name="Sheet1")
    validate_against_respondent_data(TARGET_FILE, pairs, current_month_filter=(2026, 6))
    write_idxage(TARGET_FILE, pairs, save_path=OUTPUT_FILE)