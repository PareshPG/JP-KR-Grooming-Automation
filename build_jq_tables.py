"""
Step: Build the J:Q "table two" in Worth it & UBA_App and Worth it & UBA_BR from
each sheet's own A:H "table one" - replacing the manual copy/paste + Remove
Duplicates step described by the client.

Rule (validated against the real June file - see below): remove duplicates
using every column as the match key EXCEPT the ones named below, keep the
first occurrence, preserve row order. The S:X formulas already reference a
fixed $J$2:$...$10000 bound, so we only need to write the deduped rows - no
formula edits required.

Validation performed before this script was written:
  - Worth it & UBA_App: this exact rule, applied in pandas to the real A:H
    data, reproduced the file's existing J:Q block EXACTLY (4,184 rows,
    identical respondent/KPI/brand combinations, identical row order).
  - Worth it & UBA_BR: the client's instruction reused the word "Styler",
    which doesn't exist as a column in this sheet (it has Dispo/System
    instead). Four candidate readings were tested (exclude Dispo only,
    exclude System only, exclude both, exclude neither) - all four produced
    an EXACT match to the sheet's existing J:Q block on this month's data,
    because Dispo/System happen to be constant per (Resp id, KPI) here. The
    rule below (exclude System, Country, Brand - i.e. dedup key is Month +
    Resp id + B&R + KPI + Dispo) is now CONFIRMED directly from a screenshot
    of the client's own Excel "Remove Duplicates" dialog: Month, Resp id,
    B&R, KPI, and Dispo are checked; System, Country, and Brand are not.
"""
import openpyxl
from openpyxl.utils import get_column_letter
from safe_xlsx_writer import patch_workbook

# --- per-sheet configuration -------------------------------------------------
SHEET_CONFIG = {
    "Worth it & UBA_App": {
        "exclude_from_dedup_key": ["Styler", "Country", "Brand"],
        "confirmed": True,   # validated exact match against existing file
    },
    "Worth it & UBA_BR": {
        "exclude_from_dedup_key": ["System", "Country", "Brand"],
        "confirmed": True,   # confirmed via screenshot of client's actual
                              # Remove Duplicates dialog: Month/Resp id/B&R/
                              # KPI/Dispo checked, System/Country/Brand unchecked
    },
}
TABLE_ONE_START_COL = 1   # A
TABLE_TWO_START_COL = 10  # J
DATA_START_ROW = 2
HEADER_ROW = 1


def read_headers_ro(ws, start_col, header_row=HEADER_ROW):
    """Sequential single-row read, not random ws.cell() access per column -
    read_only mode is only fast when accessed sequentially."""
    row_vals = next(ws.iter_rows(min_row=header_row, max_row=header_row,
                                    min_col=start_col, values_only=True))
    headers = []
    for v in row_vals:
        if v is None:
            break
        headers.append(v)
    return headers


def read_table_ro(wb_path, sheet_name, start_col):
    """Fast read of a table (headers + data rows starting at start_col) using
    read_only mode - openpyxl's normal mode takes ~2 minutes to load this
    workbook (tens of thousands of formulas across 23 sheets); read_only
    mode loads it in well under a second since it streams rather than
    building the full formula object model."""
    wb = openpyxl.load_workbook(wb_path, read_only=True, data_only=False)
    ws = wb[sheet_name]
    headers = read_headers_ro(ws, start_col)
    n = len(headers)
    rows = []
    for row_vals in ws.iter_rows(min_row=DATA_START_ROW, min_col=start_col,
                                    max_col=start_col + n - 1, values_only=True):
        if row_vals[0] is None:
            break
        rows.append(dict(zip(headers, row_vals)))
    wb.close()
    return headers, rows


def dedup_rows(rows, headers, exclude_cols):
    """Drop-duplicates on all headers except exclude_cols, keep first
    occurrence, preserve order."""
    key_cols = [h for h in headers if h not in exclude_cols]
    seen = set()
    out = []
    for row in rows:
        key = tuple(row[h] for h in key_cols)
        if key in seen:
            continue
        seen.add(key)
        out.append(row)
    return out


def write_table_two(editor, headers, rows, current_max_row, start_col=TABLE_TWO_START_COL):
    """Clear existing table-two data rows, then write the deduped rows, using
    a SheetEditor (safe_xlsx_writer) rather than an openpyxl worksheet."""
    end_col = start_col + len(headers) - 1
    if current_max_row >= DATA_START_ROW:
        # Only clear within the J:Q column span - clear_rows removes whole
        # <row> elements, so restrict scope by clearing just the target
        # cells if the row might have OTHER columns' data too (it doesn't
        # here - J:Q is the last content on these rows - but this keeps the
        # function correct even if that ever changes).
        from openpyxl.utils import get_column_letter as gcl
        for r in range(DATA_START_ROW, current_max_row + 1):
            for c in range(start_col, end_col + 1):
                editor.clear_cell(f"{gcl(c)}{r}")

    for i, row in enumerate(rows):
        r = DATA_START_ROW + i
        for j, h in enumerate(headers):
            from openpyxl.utils import get_column_letter as gcl
            editor.set_cell(f"{gcl(start_col + j)}{r}", row[h])


def get_table_two_max_row(wb_path, sheet_name, start_col=TABLE_TWO_START_COL):
    """Fast read-only lookup of table two's current last used row."""
    wb = openpyxl.load_workbook(wb_path, read_only=True)
    ws = wb[sheet_name]
    last_row = DATA_START_ROW - 1
    for row_vals in ws.iter_rows(min_row=DATA_START_ROW, min_col=start_col,
                                    max_col=start_col, values_only=True):
        if row_vals[0] is None:
            break
        last_row += 1
    return last_row


def make_table_two_edit(headers, deduped_rows, current_max_row):
    """Returns an edit function for use with safe_xlsx_writer.patch_workbook."""
    def edit(editor):
        write_table_two(editor, headers, deduped_rows, current_max_row)
    return edit


def build_table_two(target_wb_path, save_path=None, sheets=None):
    """Standalone entry point: compute dedup for each sheet (read-only,
    fast), then patch and save - preserving every other part of the
    workbook byte-for-byte. The monthly pipeline combines this with the
    other steps instead of calling this separately."""
    sheets = sheets or list(SHEET_CONFIG.keys())

    edit_funcs = {}
    for sheet_name in sheets:
        cfg = SHEET_CONFIG[sheet_name]
        headers, rows = read_table_ro(target_wb_path, sheet_name, TABLE_ONE_START_COL)
        deduped = dedup_rows(rows, headers, cfg["exclude_from_dedup_key"])
        current_max_row = get_table_two_max_row(target_wb_path, sheet_name)
        edit_funcs[sheet_name] = make_table_two_edit(headers, deduped, current_max_row)

        confidence = "VALIDATED" if cfg["confirmed"] else "UNCONFIRMED - please verify"
        print(f"{sheet_name}: table one = {len(rows)} rows -> table two (deduped) = "
              f"{len(deduped)} rows | dedup key excludes {cfg['exclude_from_dedup_key']} "
              f"[{confidence}]")

    out_path = save_path or target_wb_path
    patch_workbook(target_wb_path, out_path, edit_funcs)
    print(f"Saved to {out_path}")


def validate_against_existing(original_wb_path, rebuilt_wb_path, sheets=None):
    """Compare the rebuilt J:Q block against whatever was already in the
    original file, cell-for-cell on (Resp id, KPI, Brand). Uses read_only
    mode - this is a read-only comparison, no need for the slow full load."""
    sheets = sheets or list(SHEET_CONFIG.keys())

    for sheet_name in sheets:
        headers, orig_rows = read_table_ro(original_wb_path, sheet_name, TABLE_TWO_START_COL)
        _, new_rows = read_table_ro(rebuilt_wb_path, sheet_name, TABLE_TWO_START_COL)

        key = lambda row: (row.get("Resp id"), row.get("KPI"), row.get("Brand"))
        orig_keys = [key(r) for r in orig_rows]
        new_keys = [key(r) for r in new_rows]

        print(f"\n{sheet_name}: original J:Q = {len(orig_rows)} rows, "
              f"rebuilt J:Q = {len(new_rows)} rows")
        print(f"  Same set of rows: {set(orig_keys) == set(new_keys)}")
        print(f"  Same row order:   {orig_keys == new_keys}")


if __name__ == "__main__":
    build_table_two(
        target_wb_path="JP-June-ratings-cleaned.xlsx",   # chain onto the previous step
        save_path="JP-June-jq-rebuilt.xlsx",
    )
    validate_against_existing(
        original_wb_path="JP-June.xlsx",       # original, hand-maintained J:Q
        rebuilt_wb_path="JP-June-jq-rebuilt.xlsx",
    )
