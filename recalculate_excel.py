"""
Step: Recalculate the working file using real Microsoft Excel (via xlwings/COM),
not LibreOffice.

Why not LibreOffice: Worth it & UBA_App / UBA_BR use ROWS(UNIQUE(FILTER(...)))
formulas for the age-cut/segment counts. Tested directly - LibreOffice returns
a hard #VALUE! error on this exact formula shape (not a silent truncation, an
outright failure). Real Excel is required.

Requirements on the machine this runs on:
  - Microsoft Excel installed and licensed (Windows, or Mac with Excel).
  - `pip install xlwings`
  - Excel must be able to launch via COM automation. On a scheduled task /
    server context (not an interactive desktop session), this sometimes needs
    "Allow the service to interact with desktop" or an equivalent - a known
    friction point for unattended automation with Excel COM. If this ever
    needs to run on a headless Linux box with no Excel available at all, that
    is a real blocker, not something this script can work around - the two
    UNIQUE/FILTER formulas would need to be rewritten to avoid dynamic
    arrays, or reimplemented directly in Python.

What this does:
  1. Open the working file in Excel (VISIBLE - see note below - no alerts).
  2. Force a full recalculation (not just a normal one - CalculateFullRebuild
     ensures every dependency is touched, not just cells Excel thinks are dirty).
  3. Save and close.
  4. Re-open (openpyxl, read-only, cached values) and scan every cell in the
     8 in-scope tabs for formula errors, same check used throughout this
     project's manual review. Reports findings; does not decide pass/fail on
     your behalf - some errors may be expected (e.g. a genuinely tiny sample
     segment dividing by zero), so this surfaces them for a human to judge.

Why visible=True, not the more typical invisible headless automation: tested
directly on the real machine - with visible=False, Save() either triggered
an interactive Save As dialog with a garbled temp filename, or the COM
connection died outright with an RPC error. A step-by-step visible debug run
confirmed the file itself was fine (ReadOnly=False, ProtectedViewWindows.Count=0,
no Protected View), and with visible=True the exact same Save() call
completed with zero issues. The cause is invisible-mode's internal dialog/
window handling on this machine, not the file or the save method - Excel
will flash on screen briefly during each monthly run as a result, which is
an acceptable tradeoff for the save step actually working reliably. Note
for future unattended/scheduled runs: visible=True needs an active,
interactive desktop session to render a window - if this ever moves to a
Task Scheduler job running with no one logged in, this specific point would
need re-testing in that context.
"""
import openpyxl

IN_SCOPE_SHEETS = [
    "DRY SCORECARD", "AE(DRY)", "Worth it & UBA_App", "Respondent Data",
    "WET SCORECARD", "AE(WET)", "Worth it & UBA_BR", "IDxAGE",
]


def _check_not_already_open(working_path, xw):
    """Check every running Excel instance for a workbook already open at
    this path, and fail with a clear, actionable message before attempting
    anything - rather than a cryptic COM 'sharing violation' error surfacing
    deep inside Save(). This is a realistic monthly scenario (someone has
    the report open for review when the automation runs), not just a
    debugging artifact - worth a clear error regardless of cause."""
    import os
    target = os.path.normcase(os.path.abspath(working_path))
    for app in xw.apps:
        for book in app.books:
            try:
                if os.path.normcase(os.path.abspath(book.fullname)) == target:
                    raise RuntimeError(
                        f"'{working_path}' is already open in another Excel "
                        f"window (PID {app.pid}). Close it completely first - "
                        f"check Task Manager for any EXCEL.EXE process, not "
                        f"just the visible window, since a debug/automation "
                        f"session can leave one running in the background."
                    )
            except AttributeError:
                continue  # book.fullname can be unavailable for some book states


def recalculate_with_excel(working_path):
    """Open, force-recalculate, save, close - using real Excel via xlwings."""
    try:
        import xlwings as xw
    except ImportError:
        raise ImportError(
            "xlwings is not installed. On the machine that has Excel: "
            "pip install xlwings"
        )

    _check_not_already_open(working_path, xw)

    print(f"Launching Excel to recalculate {working_path} ...")
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        wb = app.books.open(working_path)
        # Full rebuild, not a normal calculate - forces every formula to
        # actually re-evaluate rather than relying on Excel's dependency
        # tracking, which can miss cells after a script has rewritten inputs.
        # This is an Application-level method, not a Workbook-level one.
        app.api.CalculateFullRebuild()
        # Raw COM Save(), not xlwings' wb.save() wrapper. Confirmed directly:
        # wb.save(working_path) with an explicit path STILL triggered an
        # interactive Save As file-picker dialog (with a garbled default
        # filename) - xlwings' save() appears to route through SaveAs
        # internally regardless, and SaveAs is the COM method prone to this
        # exact behavior. wb.api.Save() calls the underlying VBA-equivalent
        # Workbook.Save with zero arguments - the direct equivalent of
        # pressing Ctrl+S on a file already open from disk, with no
        # path/format decision for Excel to be ambiguous about at all.
        wb.api.Save()
        wb.close()
        print("Recalculation complete, file saved.")
    finally:
        app.quit()


EXCEL_ERROR_VALUES = {
    "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!",
    "#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!", "#UNKNOWN!", "#BLOCKED!",
}


def scan_for_errors(working_path, sheets=None):
    """Read back cached values and report every error cell in the in-scope
    tabs. Same check used throughout manual review of this workbook.

    Matches against the actual known Excel error strings, not just "starts
    with #" - found while testing against KO's real data: Respondent Data
    there has legitimate survey response text starting with '#' (e.g.
    '#1 manual razor', a ranking-style answer), which a blanket "#" prefix
    check would misidentify as a formula error. JP's data apparently never
    exercised this edge case, so it went unnoticed until now."""
    sheets = sheets or IN_SCOPE_SHEETS
    wb = openpyxl.load_workbook(working_path, data_only=True, read_only=True)

    errors_found = {}
    for sheet_name in sheets:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v in EXCEL_ERROR_VALUES:
                    errors_found.setdefault(sheet_name, []).append((cell.coordinate, v))

    total = sum(len(v) for v in errors_found.values())
    print(f"\nError scan across {len(sheets)} in-scope tabs: {total} error cell(s) found.")
    for sheet_name, errs in errors_found.items():
        print(f"  {sheet_name}: {len(errs)} error(s)")
        by_type = {}
        for coord, val in errs:
            by_type.setdefault(val, []).append(coord)
        for err_type, coords in by_type.items():
            sample = coords[:5]
            print(f"    {err_type}: {len(coords)} cell(s), e.g. {sample}")

    if total == 0:
        print("  No formula errors in the in-scope tabs.")
    return errors_found


def check_sample_sizes(working_path, threshold=30):
    """Flag any scorecard segment whose sample size is below the threshold -
    the 'Dispo' segment (n=6) in the June JP file is a real example of a
    segment small enough to produce unreliable ratios and, in that specific
    case, actual #DIV/0! errors downstream.

    Reads whichever of column C or D actually holds the numeric N value,
    rather than assuming a fixed column - verified directly against both
    JP and KO files that DRY SCORECARD's segment-level Sample Size sits in
    column C while WET SCORECARD's sits in column D, a pre-existing
    inconsistency in the client's own template (not introduced by anything
    in this pipeline). A hardcoded column D was silently reading None for
    every DRY SCORECARD segment - this was never caught because no DRY
    segment happened to be small enough to matter, but it means DRY was
    never actually being checked here at all.

    Uses read_only mode - a normal-mode load costs ~2 minutes on this
    23-sheet, tens-of-thousands-of-formulas workbook, even though the two
    sheets actually read here (DRY/WET SCORECARD) are small."""
    wb = openpyxl.load_workbook(working_path, data_only=True, read_only=True)
    flagged = []
    for sheet_name in ["DRY SCORECARD", "WET SCORECARD"]:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(min_col=2, max_col=4, values_only=True))
        for r in range(len(rows) - 1):
            if rows[r][0] == "Scorecard":
                segment_name = rows[r][1]
                # column C is rows[r+1][1], column D is rows[r+1][2] within
                # this B:D slice - take whichever one is actually numeric
                n = next((v for v in (rows[r + 1][1], rows[r + 1][2])
                           if isinstance(v, (int, float))), None)
                if n is not None and n < threshold:
                    flagged.append((sheet_name, segment_name, n))

    if flagged:
        print(f"\nSegments below n={threshold} (treat as directional only):")
        for sheet_name, segment, n in flagged:
            print(f"  {sheet_name} / {segment}: n={n}")
    else:
        print(f"\nNo segments below n={threshold}.")
    return flagged


if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else "JP-June_Working.xlsx"
    recalculate_with_excel(path)
    scan_for_errors(path)
    check_sample_sizes(path)
