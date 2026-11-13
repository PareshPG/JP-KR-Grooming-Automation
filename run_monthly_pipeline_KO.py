"""
Monthly pipeline orchestrator - KOREA (flat-file + fixed-template architecture).

REPLACES the previous version of this script. The client no longer sends
one fully-combined file each month - instead: a "flat file" (Respondent
Data, Worth it & UBA_App, Worth it & UBA_BR only) and a raw data file
arrive each month, while DRY/WET SCORECARD, AE(DRY), AE(WET), and IDxAGE
live in a STABLE report template reused every month. This script rebuilds
the data tabs inside a fresh copy of that template each run.

Every piece of logic here was individually validated against real data
and a known-correct ground truth before being assembled - see conversation
history for each validation (Respondent Data reconstruction: 0 mismatches
across 2,853 rows x 677 columns; Worth it & UBA dedup: exact match; the
bulk-write mechanism itself: exact match at full scale).

Usage:
    python pull_files_KO.py              <- get this month's files first
    python run_monthly_pipeline_KO.py
    python recalculate_via_macro_KO.py
    python equity_color_coding_KO.py
"""
import sys
import shutil
import tempfile
import re
import collections
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

from build_idxage import extract_id_age_pairs, make_idxage_edit, get_sheet_max_row
from clean_ratings import find_rating_columns, clean_rating_value
from build_jq_tables import dedup_rows, SHEET_CONFIG
from set_report_month import resolve_report_month, make_all_month_edits, find_month_label_cells, month_label
from safe_xlsx_writer import (
    patch_workbook, patch_workbook_bulk_replace, combine_edits, set_full_calc_on_load,
    SheetEditor, resolve_sheet_path,
)
from pipeline_config import find_single_file_in_folder
from market_config import KO
from flat_raw_alignment import check_flat_against_raw
from remap_template_layout import (
    adapt_template, format_layout_report, load_aliases, read_template_header, new_layout_header, _sheet_paths, _rd_table_parts,
)
import zipfile

# Folder paths now come from market_config.KO - no longer duplicated here.
# This duplication was the actual root cause of a real bug: when the KO
# folder structure moved under KO\, this file's own separate copies of
# these paths were updated, but market_config.py's copies were not - two
# sources of truth silently drifted apart, and recalculate_via_macro_KO.py
# was still reading the stale one.

RESPONDENT_DATA_HEADER_ROW = 10
RESPONDENT_DATA_START_ROW = 11
AGE_ABSOLUTE_INSERT_INDEX = 5  # 0-indexed - Age absolute lands at column F
# NOTE: there is deliberately NO fixed column count here any more. The width of
# Respondent Data is whatever the flat file has (its last populated header
# column + the inserted Age absolute), and the report template is re-pointed
# at that layout on every run by remap_template_layout.adapt_template().
COLUMN_ALIASES_FILE = "Column_Aliases.csv"   # optional, in the Formatting folder: old header,new header
LAYOUT_REPORT_FILE = "Layout_Change_Report.txt"  # written next to the template when the layout changed

# Expected flat-file headers, at the exact positions the reconstruction
# logic assumes them to be at. Checked against the real flat file before
# every run - see validate_flat_file_structure() below. If the client's
# export format ever changes, this is what stops silent misalignment.
EXPECTED_RESPONDENT_DATA_HEADERS_START = [
    "Index", "Resp id", "Country", "MM/YY", "Gender",
]  # columns A-E, immediately before where Age absolute gets inserted
EXPECTED_WORTH_IT_APP_HEADERS = ["Month", "Resp id", "App", "KPI", "Electric Shaver", "Styler", "Country", "Brand"]
EXPECTED_WORTH_IT_BR_HEADERS = ["Month", "Resp id", "B&R", "KPI", "Dispo", "System", "Country", "Brand"]

RESPONDENT_DATA_FORMULA_BOUNDS_TO_CHECK = [
    ("AE(DRY)", "A3"),
    ("AE(WET)", "A3"),
    ("DRY SCORECARD", "C3"),
    ("WET SCORECARD", "D3"),  # confirmed: WET's Sample Size sits in column D,
    # DRY's in column C - the same pre-existing template quirk found and
    # fixed earlier in check_sample_sizes().
    ("Worth it & UBA_App", "T5"),  # these sheets' own T:X summary formulas embed
    ("Worth it & UBA_BR", "T5"),   # direct Respondent Data references too - found
    # missing from this check during a later audit while building the KO
    # equivalent; not previously monitored here despite the discovery scan
    # having confirmed these references exist (110/145 occurrences).
]


def validate_flat_file_structure(flat_path, sheet_name, expected_headers, header_row=1):
    """Confirms the flat file's actual header row matches what the
    reconstruction logic assumes, at the exact positions it assumes them
    to be at. Raises a clear error naming exactly what differs, rather
    than silently truncating or misreading columns if the client's export
    structure ever changes - confirmed this was a real, unguarded gap
    before this check existed."""
    wb = openpyxl.load_workbook(flat_path, data_only=True, read_only=True)
    ws = wb[sheet_name]
    actual = next(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))
    actual = actual[:len(expected_headers)]

    mismatches = [(i, expected_headers[i], actual[i] if i < len(actual) else None)
                  for i in range(len(expected_headers)) if (actual[i] if i < len(actual) else None) != expected_headers[i]]

    if mismatches:
        details = "; ".join(f"col {i+1}: expected {exp!r}, found {act!r}" for i, exp, act in mismatches)
        raise RuntimeError(
            f"STOP: '{sheet_name}' in the flat file doesn't match the expected "
            f"structure this pipeline was built against: {details}. Proceeding could "
            f"silently misalign columns or drop new ones. Check with the client "
            f"whether their export format changed, then update the pipeline's "
            f"expected structure to match before re-running."
        )
    print(f"  Structure check OK: '{sheet_name}' header matches expected layout.")


def check_formula_headroom(template_path, actual_last_data_row, warn_threshold=400):
    """Scan the template's own formulas that reference Respondent Data's
    row range, extract their hardcoded bound directly from the formula
    text (not a hardcoded copy of my own - that duplication is exactly
    what caused a real bug earlier), and compare against how many rows
    were actually just reconstructed. FILTER()/COUNTIFS() formulas fail
    SILENTLY when data exceeds their bound - no error, just missing rows
    - so this needs to be loud, not just informational, when the risk is
    real.

    Returns a list of (sheet, cell, bound, headroom, status) and raises
    RuntimeError if any bound is already exceeded."""
    wb = openpyxl.load_workbook(template_path, data_only=False, read_only=True)
    results = []
    exceeded = []

    for sheet_name, cell_ref in RESPONDENT_DATA_FORMULA_BOUNDS_TO_CHECK:
        ws = wb[sheet_name]
        v = ws[cell_ref].value
        formula_text = v.text if hasattr(v, "text") else v
        if not formula_text:
            continue
        bounds = re.findall(r"'Respondent Data'!\$[A-Z]+\$\d+:\$[A-Z]+\$(\d+)", formula_text)
        if not bounds:
            continue
        bound = min(int(b) for b in bounds)
        headroom = bound - actual_last_data_row
        if headroom < 0:
            status = "EXCEEDED"
            exceeded.append((sheet_name, cell_ref, bound, headroom))
        elif headroom < warn_threshold:
            status = "LOW HEADROOM"
        else:
            status = "OK"
        results.append((sheet_name, cell_ref, bound, headroom, status))

    print(f"\nFormula range headroom check (data currently ends at row {actual_last_data_row}):")
    for sheet_name, cell_ref, bound, headroom, status in results:
        print(f"  {sheet_name}!{cell_ref}: bound=row {bound}, headroom={headroom} rows [{status}]")

    if exceeded:
        details = "; ".join(f"{s}!{c} (bound row {b}, over by {-h} rows)" for s, c, b, h in exceeded)
        raise RuntimeError(
            f"STOP: Respondent Data now exceeds the hardcoded range of one or more "
            f"template formulas: {details}. These formulas will silently EXCLUDE the "
            f"newest respondents from their calculations - the report would be wrong "
            f"with no visible error. The template's formula ranges need to be manually "
            f"extended in Excel before this report can be trusted. Do not deliver this "
            f"report until that's done."
        )
    return results


def validate_template_layout(adapted_template_path, flat_path, expected_table_last_row=None):
    """Safety net AFTER adaptation: the adapted template's Respondent Data header (row 10) must equal the
    header this run is about to write (flat header with Age absolute at F). The scorecard formulas read
    columns by letter, so any difference here would put the wrong questions under the scorecards with no
    error. adapt_template() is built so this can never fail - if it does, that is a bug, so stop."""
    def _n(x):
        return None if x in (None, "") else re.sub(r"\s+", " ", str(x)).strip()
    expected = [_n(h) for h in new_layout_header(flat_path)]
    zf = zipfile.ZipFile(adapted_template_path)
    paths = _sheet_paths(zf)
    got = [_n(h) for h in read_template_header(zf, paths)]
    # The Respondent Data Excel table (if the template has one) must describe exactly the header row and data
    # range this run writes - otherwise Excel opens the file with "We found a problem with some content ... Table".
    from lxml import etree
    q = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    raw_header = [str(h) for h in new_layout_header(flat_path)]
    table_problems = []
    for part in _rd_table_parts(zf, paths["Respondent Data"]):
        t = etree.fromstring(zf.read(part))
        names = [c.get("name") for c in t.find(q + "tableColumns")]
        want_ref = f"A10:{get_column_letter(len(raw_header))}{expected_table_last_row}" if expected_table_last_row else None
        if names != raw_header:
            table_problems.append(f"{part}: column names differ from the header row")
        if want_ref and t.get("ref") != want_ref:
            table_problems.append(f"{part}: range {t.get('ref')} should be {want_ref}")
        if t.find(q + "autoFilter") is not None and len(t.find(q + "autoFilter")):
            table_problems.append(f"{part}: still carries a column filter")
    zf.close()
    if table_problems:
        raise RuntimeError("STOP: internal check failed - the Respondent Data table does not match this run's layout: "
                           + "; ".join(table_problems) + ". Do not deliver this report.")
    n = len(expected)
    mism = [i for i in range(max(n, len(got))) if (expected[i] if i < n else None) != (got[i] if i < len(got) else None)]
    if mism:
        raise RuntimeError(
            f"STOP: internal check failed - after adaptation the template header still differs from the flat file "
            f"in {len(mism)} column(s) (first: column {get_column_letter(mism[0] + 1)}). Do not deliver this report; "
            f"send the files to whoever maintains the pipeline.")
    print(f"  Adapted template header matches the flat file in all {n} columns.")


def check_ae_capacity(template_path, rd_rows, year, month):
    """The AE(DRY)/AE(WET) tabs hold ONE formula row per respondent of the
    report month (rows 3 down; column A is a FILTER() spill), and the
    scorecards read those rows through fixed ranges. A month with more
    respondents than rows silently drops the extra respondents from the
    Equity Score - no error. September 2026 was the first month to exceed
    the template's original 300 rows (501 respondents). Stops if the month
    does not fit; the fix is remap_template_layout.py's AE_LAST_ROW."""
    n = sum(1 for r in rd_rows
            if hasattr(r[3], "year") and (r[3].year, r[3].month) == (year, month))
    wb = openpyxl.load_workbook(template_path, read_only=True)
    capacity = min(wb[s].max_row for s in ("AE(DRY)", "AE(WET)")) - 2
    print(f"  AE capacity check: {n} respondents in {year}-{month:02d}, AE tabs have formula rows for {capacity}.")
    if n > capacity:
        raise RuntimeError(
            f"STOP: {year}-{month:02d} has {n} respondents but the AE(DRY)/AE(WET) tabs only have "
            f"formula rows for {capacity}. The extra {n - capacity} respondents would be silently left "
            f"out of the Equity Score. Extend the AE rows and the scorecard ranges that read them "
            f"(AE_LAST_ROW in remap_template_layout.py) before delivering this month."
        )


def get_style_map(template_path, sheet_name, sample_row, num_cols):
    """One style index per column, taken from an existing data row -
    confirmed the real template uses one consistent style per column
    across all rows, not per-row styling."""
    zf = zipfile.ZipFile(template_path)
    path = resolve_sheet_path(zf, sheet_name)
    editor = SheetEditor(zf.read(path))
    zf.close()
    return [editor.get_cell_style_index(f"{get_column_letter(c)}{sample_row}") for c in range(1, num_cols + 1)]


def reconstruct_respondent_data(flat_path, idxage_lookup, num_cols, known_brands=None, raw_path=None, raw_brand_names=None, repair_flat=False):
    """Returns (rows, num_cols) - the fully reconstructed Respondent Data
    content: Age absolute inserted at column F (computed from idxage_lookup,
    blank for respondents not in the current month's IDxAGE) and rating
    columns cleaned using the same logic as clean_ratings.py. num_cols is
    the flat file's own populated width + 1 (Age absolute), worked out from
    its header row - nothing is cut off and nothing is assumed."""
    wb = openpyxl.load_workbook(flat_path, data_only=True, read_only=True)
    ws = wb["Respondent Data"]
    ws.reset_dimensions()  # do not trust the file's own <dimension> tag for the width

    header = next(ws.iter_rows(min_row=RESPONDENT_DATA_HEADER_ROW, max_row=RESPONDENT_DATA_HEADER_ROW, values_only=True))

    rating_cols, _ = find_rating_columns(ws, header_row=RESPONDENT_DATA_HEADER_ROW,
                                         **({"known_brands": known_brands} if known_brands else {}))
    rating_col_indices_0based = {c - 1 for c, _ in rating_cols}

    flat_rows = [r for r in ws.iter_rows(min_row=RESPONDENT_DATA_START_ROW, values_only=True) if r[0] is not None]

    # the flat file is made by hand from the raw export: prove each brand column holds the answers of its own brand
    # (a block pasted one column off calculates without any error and reports the wrong brands). STOPs unless repair_flat.
    if raw_path and raw_brand_names:
        flat_rows = check_flat_against_raw(header, flat_rows, raw_path, raw_brand_names, repair=repair_flat)

    reconstructed = []
    for row in flat_rows:
        resp_id = row[1]
        age_absolute = idxage_lookup.get(resp_id, None)
        cleaned_tail = []
        for i, v in enumerate(row[AGE_ABSOLUTE_INSERT_INDEX:], start=AGE_ABSOLUTE_INSERT_INDEX):
            if i in rating_col_indices_0based and isinstance(v, str):
                v, _ = clean_rating_value(v)
            cleaned_tail.append(v)
        new_row = (list(row[:AGE_ABSOLUTE_INSERT_INDEX]) + [age_absolute] + cleaned_tail)
        reconstructed.append(new_row[:num_cols])  # drop the flat file's own trailing blank slots

    return reconstructed, num_cols


def reconstruct_worth_it_sheet(flat_path, sheet_name):
    """Returns (rows, num_cols, n_table_one, n_table_two) - table one (raw)
    and table two (deduped) combined into one full-width row list, since
    both tables share the same row range in the actual sheet. The two
    tables have DIFFERENT row counts (table two is always shorter after
    deduping), so the shorter one is padded with blanks rather than
    assumed to align 1:1."""
    cfg = SHEET_CONFIG[sheet_name]
    wb = openpyxl.load_workbook(flat_path, data_only=True, read_only=True)
    ws = wb[sheet_name]

    header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    table_one_headers = header[:8]

    table_one = []
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=8, values_only=True):
        if row[0] is None:
            break
        table_one.append(list(row))

    dicts = [dict(zip(table_one_headers, r)) for r in table_one]
    deduped_dicts = dedup_rows(dicts, table_one_headers, cfg["exclude_from_dedup_key"])
    table_two = [[d[h] for h in table_one_headers] for d in deduped_dicts]

    max_len = max(len(table_one), len(table_two))
    combined_rows = []
    for i in range(max_len):
        t1 = table_one[i] if i < len(table_one) else [None] * 8
        t2 = table_two[i] if i < len(table_two) else [None] * 8
        combined_rows.append(t1 + [None] + t2)  # column I is a blank gap between tables

    return combined_rows, 17, len(table_one), len(table_two)


def clean_stale_working_files(report_folder):
    """Delete any leftover working files sitting in the Report folder
    before creating fresh ones this run. Matches both the single-month
    pattern (<stem>_Working.xlsx) and the quarterly pattern
    (<stem>_Working_YYYY-MM.xlsx, and the intermediate
    <stem>_Base_Working.xlsx which should never survive a completed run
    but is matched here too as a safety net in case a prior run was
    interrupted before reaching its own cleanup).

    Added after a real, confirmed incident: a stale working file left
    over from an earlier run coexisted with a freshly-created one, and
    the recalculation macro's own file search picked one without any
    clear indication of which - only caught because the numbers happened
    to be checked carefully rather than assumed correct. This removes
    the possibility entirely rather than relying on remembering to clean
    up manually."""
    folder = Path(report_folder)
    pattern = re.compile(r"_(Base_)?Working(_\d{4}-\d{2})?\.xlsx$")
    stale = [f for f in folder.iterdir() if f.is_file() and pattern.search(f.name)]
    if stale:
        for f in stale:
            f.unlink()
        names = ", ".join(f.name for f in stale)
        print(f"  Cleared {len(stale)} stale working file(s) before starting: {names}")
    else:
        print("  No stale working files found - clean start.")


def prompt_for_report_months():
    """Interactive fallback for quarterly (or any multi-month) delivery.
    The client has moved from monthly to quarterly delivery - one flat
    file now arrives covering 3 months at once, and one output report
    file is needed per month, all built from the same reconstruction
    pass (see Step 6). Accepts one or more comma-separated YYYY-MM values
    on a single line, so a single-month ad-hoc run (just one value) and
    a full quarterly run (three values) both work through the same
    prompt without a separate mode to remember."""
    while True:
        raw = input("Enter report month(s) as comma-separated YYYY-MM values "
                    "(e.g. 2026-07,2026-08,2026-09): ").strip()
        parts = [p.strip() for p in raw.split(",") if p.strip()]
        if not parts:
            print("  Enter at least one month.")
            continue
        try:
            months = []
            for part in parts:
                year_s, month_s = part.split("-")
                year, month = int(year_s), int(month_s)
                if not (1 <= month <= 12):
                    raise ValueError
                months.append((year, month))
            return months
        except ValueError:
            print(f"  Couldn't parse {raw!r} as comma-separated YYYY-MM values, try again.")


def run_pipeline(flat_file=None, raw_data_file=None, report_template=None, report_months=None, repair_flat=False):
    """report_months: list of (year, month) tuples - one stamped, ready-
    to-recalculate output file is produced per entry, ALL from a single
    data-reconstruction pass (Steps 0-5 run exactly once regardless of
    how many months are requested - only Step 6's stamping repeats).
    Pass a single-element list for an ad-hoc one-month run. Leave as
    None to be prompted interactively."""
    print("=" * 70)
    print("STEP 0: Locate files")
    print("=" * 70)
    if flat_file is None:
        flat_file = find_single_file_in_folder(KO.flat_file_folder)
    if raw_data_file is None:
        raw_data_file = find_single_file_in_folder(KO.raw_data_folder)
    if report_template is None:
        report_template = find_single_file_in_folder(
            KO.report_folder, name_filter=lambda name: not name.endswith("_Working.xlsx")
        )
    print(f"Flat file:       {flat_file}")
    print(f"Raw data file:   {raw_data_file}")
    print(f"Report template: {report_template}")
    clean_stale_working_files(Path(report_template).parent)  # = KO.report_folder in production

    print("\n" + "=" * 70)
    print("STEP 0b: Validate flat file structure")
    print("=" * 70)
    validate_flat_file_structure(str(flat_file), "Respondent Data",
                                   EXPECTED_RESPONDENT_DATA_HEADERS_START, header_row=RESPONDENT_DATA_HEADER_ROW)
    validate_flat_file_structure(str(flat_file), "Worth it & UBA_App", EXPECTED_WORTH_IT_APP_HEADERS)
    validate_flat_file_structure(str(flat_file), "Worth it & UBA_BR", EXPECTED_WORTH_IT_BR_HEADERS)

    print("\n" + "=" * 70)
    print("STEP 0c: Adapt the report template to this flat file's layout")
    print("=" * 70)
    # The scorecard/AE/Worth it formulas read Respondent Data by column LETTER, so a changed questionnaire
    # (columns added / removed / moved) would silently point them at the wrong questions. Instead of
    # depending on a fixed layout, every run re-points an in-memory COPY of the template at the flat
    # file's real layout (by header text), sizes the row limits from the data, and uses that copy.
    # The template file in the Report folder is never modified. A question the formulas read that the
    # flat file no longer has STOPS the run with the column named.
    template_path = Path(report_template)
    work_dir = Path(tempfile.mkdtemp(prefix="jp_template_"))
    source_template = work_dir / template_path.name
    aliases = load_aliases(Path(KO.formatting_folder) / COLUMN_ALIASES_FILE)
    if aliases:
        print(f"  Column aliases loaded: {len(aliases)} ({COLUMN_ALIASES_FILE})")
    layout = adapt_template(str(template_path), str(flat_file), str(source_template), aliases)
    layout_file = template_path.with_name(LAYOUT_REPORT_FILE)
    if layout["moved"] or layout["added"] or layout["removed"]:
        layout_file.write_text(format_layout_report(layout), encoding="utf-8")
        print(f"  Layout differs from the template: {layout['moved']} moved, {len(layout['added'])} new, "
              f"{len(layout['removed'])} removed - details in {layout_file.name}")
        for col, header, used in layout["removed"]:
            if not used:
                print(f"    note: column {col} ({header[:60]!r}) is gone from the flat file - no formula read it")
    else:
        print("  Layout identical to the template - nothing to re-point.")
    validate_template_layout(str(source_template), str(flat_file), layout["table_last_row"])
    num_cols = len(new_layout_header(str(flat_file)))
    print(f"  Respondent Data width for this run: {num_cols} columns")

    base_working_path = template_path.with_name(f"{template_path.stem}_Base_Working{template_path.suffix}")
    shutil.copy2(source_template, base_working_path)
    print(f"Base working copy: {base_working_path}")
    print("  (intermediate file - not itself a deliverable; per-month copies are produced in Step 6)")

    print("\n" + "=" * 70)
    print(f"STEP 1: Rebuild {KO.age_sheet_name} (respondent ages) from raw data")
    print("=" * 70)
    pairs = extract_id_age_pairs(str(raw_data_file), sheet_name="Sheet1",
                                 age_header=KO.age_question_header,
                                 minor_age_header=KO.minor_age_question_header)
    idxage_lookup = {rid: age for rid, age in pairs}
    idxage_max_row = get_sheet_max_row(str(base_working_path), KO.age_sheet_name)
    idxage_edit = make_idxage_edit(pairs, idxage_max_row, duplicate_column=KO.age_sheet_has_duplicate_column)

    print("\n" + "=" * 70)
    print("STEP 2: Reconstruct Respondent Data")
    print("=" * 70)
    rd_rows, rd_cols = reconstruct_respondent_data(str(flat_file), idxage_lookup, num_cols, KO.known_brands,
                                                   str(raw_data_file), KO.raw_brand_names, repair_flat)
    if RESPONDENT_DATA_HEADER_ROW + len(rd_rows) != layout["table_last_row"]:
        raise RuntimeError(f"STOP: internal check failed - the Respondent Data table was sized for {layout['table_last_row'] - RESPONDENT_DATA_HEADER_ROW} rows "
                           f"but this run writes {len(rd_rows)}. Do not deliver this report.")
    print(f"Reconstructed {len(rd_rows)} rows x {rd_cols} columns")
    # consistency checks on Age absolute (column F), which every age-cut row of the scorecards reads:
    #  (a) the raw file covers one month: every respondent of the month that has the most ages must have an age -
    #      a respondent-id type/format mismatch between raw and flat would otherwise leave ages blank without an error;
    #  (b) the flat file's own age band (column G, "16-17", "18-24", ...) must agree with the age this run computed -
    #      a "16-17" respondent carrying an adult age (the parent's) is the problem this check exists to catch.
    _aged = collections.Counter(); _all = collections.Counter(); _missing = collections.defaultdict(list)
    for _r in rd_rows:
        _all[_r[3]] += 1
        if _r[5] in (None, ""):
            _missing[_r[3]].append(_r[1])
        else:
            _aged[_r[3]] += 1
    if not _aged:
        raise RuntimeError("STOP: no respondent in the flat file was found in the raw data file (by Internal Respondent Id), so "
                           "no age could be set. Check that the raw file is this month's, and that its ids are numbers like the flat file's.")
    _mm = max(_aged, key=_aged.get)
    if _missing[_mm]:
        raise RuntimeError(f"STOP: {len(_missing[_mm])} of the {_all[_mm]} respondents of {_mm} have no age in the raw file "
                           f"(e.g. ids {_missing[_mm][:3]}). The age-cut rows would be wrong. Check the raw data file.")
    _band_ok = _band_bad = _band_blank = 0
    _bad_ex = []
    for _r in rd_rows:
        _band, _age = _r[6], _r[5]
        if _band in (None, "") or _age in (None, ""):
            _band_blank += 1
            continue
        _m = re.match(r"^\s*(\d+)\s*[-\u2013]\s*(\d+)\s*$", str(_band)) or re.match(r"^\s*(\d+)\s*\+\s*$", str(_band))
        if not _m:
            _band_blank += 1
            continue
        _lo = int(_m.group(1)); _hi = int(_m.group(2)) if _m.lastindex == 2 else 999
        if _lo <= _age <= _hi:
            _band_ok += 1
        else:
            _band_bad += 1
            if len(_bad_ex) < 5:
                _bad_ex.append((_r[1], _band, _age))
    print(f"Age check: {_band_ok} respondents' Age absolute agrees with the flat file's age band, "
          f"{_band_bad} disagree, {_band_blank} not comparable (prior months have no raw age).")
    if _band_bad:
        print(f"  WARNING: {_band_bad} respondents have an Age absolute outside their age band, e.g. {_bad_ex}. "
              f"Check the raw file's age columns before delivering - the age-cut rows (16-34 / 35-49 / 50+) read Age absolute.")
    rd_style_map = get_style_map(str(source_template), "Respondent Data", RESPONDENT_DATA_START_ROW, rd_cols)

    actual_last_data_row = RESPONDENT_DATA_START_ROW + len(rd_rows) - 1
    check_formula_headroom(str(source_template), actual_last_data_row)

    print("\n" + "=" * 70)
    print("STEP 3: Reconstruct Worth it & UBA_App / UBA_BR")
    print("=" * 70)
    worth_it_results = {}
    for sheet_name in ["Worth it & UBA_App", "Worth it & UBA_BR"]:
        rows, num_cols, n_one, n_two = reconstruct_worth_it_sheet(str(flat_file), sheet_name)
        style_map = get_style_map(str(source_template), sheet_name, 2, num_cols)
        worth_it_results[sheet_name] = (rows, style_map)
        print(f"{sheet_name}: table one={n_one} rows -> table two (deduped)={n_two} rows "
              f"(summary formulas read down to row {layout['worth_bound']}; headroom {layout['worth_bound'] - (n_one + 1)} rows)")
        if n_one + 1 > layout['worth_bound']:
            raise RuntimeError(
                f"STOP: {sheet_name} table one now has {n_one} rows but its summary formulas stop at row "
                f"{layout['worth_bound']}. Rows beyond that would be silently excluded - this should have been sized "
                f"automatically, so stop and check remap_template_layout.py before delivering.")

    print("\n" + "=" * 70)
    print(f"STEP 4: Write {KO.age_sheet_name} (DOM edit - small sheet)")
    print("=" * 70)
    patch_workbook(str(base_working_path), str(base_working_path), {KO.age_sheet_name: idxage_edit})
    print(f"{KO.age_sheet_name} written.")

    print("\n" + "=" * 70)
    print("STEP 5: Bulk-write Respondent Data and Worth it & UBA tabs")
    print("=" * 70)
    bulk_replacements = {
        "Respondent Data": (RESPONDENT_DATA_START_ROW, rd_rows, rd_style_map),
    }
    for sheet_name, (rows, style_map) in worth_it_results.items():
        # preserve_columns_beyond=17: these sheets have a small block of
        # live summary formulas in columns T onward (COUNTIFS, SUMPRODUCT,
        # even a UNIQUE/FILTER combination) that must never be touched -
        # confirmed via a real recalculation run that silently destroying
        # them here caused a cascade of #DIV/0! errors throughout
        # DRY/WET SCORECARD, since those formulas depend on this block.
        bulk_replacements[sheet_name] = (2, rows, style_map, 17)
    patch_workbook_bulk_replace(str(base_working_path), str(base_working_path), bulk_replacements)
    print("Bulk write complete.")
    print(f"\nThe reconstruction above (Steps 1-5) ran exactly once and covers ALL months "
          f"present in the flat file, regardless of how many report months are requested "
          f"below - this is the piece that made quarterly delivery efficient: only the "
          f"stamping in Step 6 repeats per month, not the data reconstruction itself.")

    print("\n" + "=" * 70)
    print("STEP 6: Set report month(s) and produce per-month output files")
    print("=" * 70)
    if report_months is None:
        report_months = prompt_for_report_months()

    output_paths = []
    for year, month in report_months:
        month_working_path = template_path.with_name(
            f"{template_path.stem}_Working_{year:04d}-{month:02d}{template_path.suffix}"
        )
        shutil.copy2(base_working_path, month_working_path)
        print(f"\n  -- {year:04d}-{month:02d} --")
        check_ae_capacity(str(source_template), rd_rows, year, month)
        y, m, exact_date, month_name = resolve_report_month(str(month_working_path), year=year, month=month)
        label_cells = find_month_label_cells(str(month_working_path))
        month_edits = make_all_month_edits(exact_date, month_name, label_cells, month_label(y, m))
        patch_workbook(str(month_working_path), str(month_working_path), month_edits)
        set_full_calc_on_load(str(month_working_path))   # Excel must recalculate everything when it opens this file
        print(f"  Report month cells set -> {month_working_path.name}")
        output_paths.append(month_working_path)

    base_working_path.unlink()
    shutil.rmtree(work_dir, ignore_errors=True)
    print(f"\n  (Removed intermediate base file: {base_working_path.name} - it was never a "
          f"deliverable on its own, only a shared source for the per-month copies above)")

    print(f"\nDone. Produced {len(output_paths)} report file(s), formulas not recalculated yet:")
    for p in output_paths:
        print(f"  {p}")
    print(f"\nNext: each file above still needs recalculating. recalculate_via_macro_KO.py "
          f"is being updated to handle multiple files in one pass - check current status "
          f"before assuming a single call handles all of them; it may currently need "
          f"pointing at one file at a time.\n")
    return output_paths


if __name__ == "__main__":
    run_pipeline(repair_flat="--repair-flat-from-raw" in sys.argv)