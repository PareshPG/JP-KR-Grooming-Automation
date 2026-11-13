"""
DEBUG ONLY - not for production use. Same steps as recalculate_with_excel(),
but with the Excel window VISIBLE and a pause after each stage, so you can
watch exactly what happens and when, instead of us inferring it after the
fact from a screenshot or a stack trace.

Run this directly:
    python debug_recalc_visible_JP.py
(edit WORKING_PATH below if needed first)

At each pause, just note what you see (or take a screenshot) and press
Enter to continue. If a dialog appears, DON'T dismiss it yet - screenshot
it first so we can see its exact title and content, then dismiss it and
continue.
"""
import xlwings as xw

WORKING_PATH = r"C:\JP KR Grooming\Python\JP-June_Working.xlsx"


def main():
    print(f"Target file: {WORKING_PATH}")
    input("Press Enter to launch Excel (VISIBLE this time)...")

    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    print(f"Excel launched. PID: {app.pid}")
    input("Excel should be visible now with no document open. Press Enter to open the file...")

    wb = app.books.open(WORKING_PATH)
    print(f"Opened. wb.name = {wb.name!r}, wb.fullname = {wb.fullname!r}")
    print(f"ReadOnly: {wb.api.ReadOnly}")
    try:
        print(f"ProtectedViewWindows.Count: {app.api.ProtectedViewWindows.Count}")
    except Exception as e:
        print(f"Could not check ProtectedViewWindows: {e}")
    input("Look at the Excel window now - any yellow banner (Protected View / "
          "Enable Editing / Enable Content)? Note it, then press Enter to continue...")

    print("Running CalculateFullRebuild...")
    app.api.CalculateFullRebuild()
    print("Calculation done.")
    input("Press Enter to attempt Save...")

    print(f"Before save - wb.api.Saved: {wb.api.Saved}")
    try:
        wb.api.Save()
        print("Save() returned without error.")
    except Exception as e:
        print(f"Save() raised: {e}")

    input("Whatever happened during Save, note it now (screenshot if a dialog "
          "appeared), then press Enter to finish and leave Excel open for "
          "inspection (this script will NOT close Excel automatically)...")

    print("Done. Excel is left open/visible - close it manually when ready.")


if __name__ == "__main__":
    main()
