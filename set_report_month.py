"""
Step: Set the "current report month" across the six hardcoded reference cells
found in DRY SCORECARD!B1, WET SCORECARD!B1, AE(DRY)!A1, AE(WET)!A1,
Worth it & UBA_App!T1, Worth it & UBA_BR!T1.

Design choice - why this asks for YEAR + MONTH, not an exact date:
  DRY SCORECARD!B1 / WET SCORECARD!B1 / AE(DRY)!A1 / AE(WET)!A1 are matched
  with EXACT date equality against Respondent Data's MM/YY column in several
  formulas (e.g. COUNTIFS('Respondent Data'!$D$..., $B$1, ...)) - not a
  month range. Checked this file directly: every respondent row for a given
  month shares one identical stamped date (April/May/June all landed on the
  26th here) - not the 1st, not the last day of the month, and not
  guaranteed to be the same day next month either. Asking a person to type
  that exact day from memory is exactly the kind of thing that fails
  silently - a wrong day means zero matching rows, and the scorecard would
  quietly show blank/zero instead of erroring.

  So: the user provides year + month, and this script finds the ACTUAL
  stamped date by reading it directly out of Respondent Data - guaranteed
  to match whatever day it turns out to be. If Respondent Data has no rows
  for that month yet, or (unexpectedly) more than one distinct date for the
  same month, this stops and says so rather than guessing.

Why there's no default month anywhere in this pipeline: a leftover default
from last month is exactly the failure mode this whole step exists to
prevent - it would run without complaint and quietly regenerate last
month's report. The caller must either pass year/month explicitly or let
this module prompt for it interactively.
"""
import datetime
import openpyxl
from safe_xlsx_writer import patch_workbook

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

CELL_TARGETS = {
    "DRY SCORECARD": ("B1", "date"),
    "WET SCORECARD": ("B1", "date"),
    "AE(DRY)": ("A1", "date"),
    "AE(WET)": ("A1", "date"),
    "Worth it & UBA_App": ("T1", "month_name"),
    "Worth it & UBA_BR": ("T1", "month_name"),
}


def derive_exact_date(ws_respondent_data, year, month, header_row=10):
    """Read Respondent Data's month column (found by header name, not fixed
    letter) and return the single exact date used for this year/month.
    Raises if zero or more than one distinct date is found.

    Uses sequential iter_rows, not random ws.cell() access - read_only mode
    is only fast when accessed sequentially (measured: ~89 random cell()
    calls took 15.7s; the equivalent sequential scan of the whole sheet
    takes under 10s)."""
    header_vals = next(ws_respondent_data.iter_rows(
        min_row=header_row, max_row=header_row, values_only=True))
    month_col = None
    for c, v in enumerate(header_vals, start=1):
        if v == "MM/YY":
            month_col = c
            break
    if month_col is None:
        raise ValueError("Could not find the 'MM/YY' header in Respondent Data")

    found_dates = set()
    for row_vals in ws_respondent_data.iter_rows(
        min_row=header_row + 1, min_col=month_col, max_col=month_col, values_only=True
    ):
        v = row_vals[0]
        if isinstance(v, datetime.datetime) and v.year == year and v.month == month:
            found_dates.add(v)

    if not found_dates:
        raise ValueError(
            f"No rows found in Respondent Data for {year}-{month:02d}. "
            f"Has this month's data actually been added yet?"
        )
    if len(found_dates) > 1:
        raise ValueError(
            f"Found {len(found_dates)} different dates for {year}-{month:02d} "
            f"in Respondent Data: {sorted(found_dates)}. Expected exactly one "
            f"stamped date per month - this needs a human look before "
            f"proceeding, since it's not clear which date the formulas should use."
        )
    return found_dates.pop()


def prompt_for_month():
    """Interactive fallback when year/month weren't supplied programmatically."""
    while True:
        raw = input("Enter report month as YYYY-MM (e.g. 2026-07): ").strip()
        try:
            year_s, month_s = raw.split("-")
            year, month = int(year_s), int(month_s)
            if not (1 <= month <= 12):
                raise ValueError
            return year, month
        except ValueError:
            print("  Not a valid YYYY-MM value, try again.")


def resolve_report_month(target_wb_path, year=None, month=None):
    """Resolve the report month (from arguments, or an interactive prompt if
    either is missing) and derive its exact stamped date from Respondent
    Data via a fast read-only pass. Returns (year, month, exact_date,
    month_name) - does not open the workbook for writing at all."""
    if year is None or month is None:
        year, month = prompt_for_month()

    wb = openpyxl.load_workbook(target_wb_path, read_only=True)
    exact_date = derive_exact_date(wb["Respondent Data"], year, month)
    month_name = MONTH_NAMES[month - 1]

    print(f"Report month: {year}-{month:02d} -> exact stamped date found in "
          f"Respondent Data: {exact_date.date()}")
    return year, month, exact_date, month_name


def make_month_cell_edit(cell_ref, value):
    """Returns an edit function for a single cell - used to build one
    closure per target sheet, which the orchestrator combines with any
    other edit that sheet also needs (Worth it & UBA_App/UBA_BR each get
    both a month-cell edit AND a J:Q rebuild edit)."""
    def edit(editor):
        editor.set_cell(cell_ref, value)
        print(f"  set {cell_ref} = {value!r}")
    return edit


LABEL_SHEETS = ("DRY SCORECARD", "WET SCORECARD")
LABEL_RE = None  # compiled lazily: typed month labels such as  Apr'26 / Jun'26


def month_label(year, month):
    """The scorecards' typed month label format, e.g. Sep'26."""
    return f"{MONTH_NAMES[month - 1][:3]}'{str(year)[2:]}"


def find_month_label_cells(workbook_path, sheet_names=LABEL_SHEETS):
    """{sheet: [cell refs]} of cells holding a TYPED month label like Apr'26 (the WET SCORECARD has 21 of them,
    over its 'Equity Attributes' headers). They are plain text, so nothing recalculates them: left alone they
    still say the template's old month. Found by pattern, not by position, so a changed layout cannot hide them."""
    import re
    import zipfile
    from lxml import etree
    from safe_xlsx_writer import resolve_sheet_path, MAIN_NS
    pat = re.compile(r"^[A-Z][a-z]{2}'\d{2}$")
    zf = zipfile.ZipFile(workbook_path)
    shared = []
    if "xl/sharedStrings.xml" in zf.namelist():
        for si in etree.fromstring(zf.read("xl/sharedStrings.xml")).iterfind(f"{{{MAIN_NS}}}si"):
            shared.append("".join(si.itertext()).strip())
    found = {}
    for name in sheet_names:
        root = etree.fromstring(zf.read(resolve_sheet_path(zf, name)))
        for c in root.iterfind(f".//{{{MAIN_NS}}}c"):
            t = c.get("t")
            text = None
            if t == "s":
                v = c.find(f"{{{MAIN_NS}}}v")
                if v is not None and v.text is not None and int(v.text) < len(shared):
                    text = shared[int(v.text)]
            elif t == "inlineStr":
                text = "".join(c.itertext()).strip()
            if text and pat.match(text):
                found.setdefault(name, []).append(c.get("r"))
    zf.close()
    return found


def make_all_month_edits(exact_date, month_name, label_cells=None, label=None):
    """Returns {sheet_name: edit_function} for all six reference cells, plus (if label_cells is given) the
    typed month labels on the scorecards."""
    edits = {}
    for sheet_name, (cell_ref, kind) in CELL_TARGETS.items():
        value = exact_date if kind == "date" else month_name
        edits[sheet_name] = make_month_cell_edit(cell_ref, value)
    for sheet_name, refs in (label_cells or {}).items():
        base = edits.get(sheet_name)

        def relabel(editor, refs=refs, base=base):
            if base:
                base(editor)
            for ref in refs:
                editor.set_cell(ref, label)
            print(f"  relabelled {len(refs)} typed month label(s) on this sheet -> {label!r}")
        edits[sheet_name] = relabel
    return edits


if __name__ == "__main__":
    # Standalone test: operates on a copy so it's safe to run repeatedly.
    import shutil

    shutil.copy2("JP-June.xlsx", "JP-June_monthtest.xlsx")
    year, month, exact_date, month_name = resolve_report_month(
        "JP-June_monthtest.xlsx", 2026, 6
    )
    edits = make_all_month_edits(exact_date, month_name)
    patch_workbook("JP-June_monthtest.xlsx", "JP-June_monthtest.xlsx", edits)
    print("\nSaved JP-June_monthtest.xlsx")