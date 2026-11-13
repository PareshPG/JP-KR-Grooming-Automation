"""
Validates the 'Rating' columns inside AE(DRY) and AE(WET) for messy text
values that need the same numeric cleanup already applied to Respondent
Data's own rating columns.

WHY THIS EXISTS: most cells in these Rating columns are formulas pulling
live from Respondent Data (e.g. AE(DRY)!G3 =
IFERROR(INDEX('Respondent Data'!$OX$11:$OX$3000,MATCH(...)),"")) - once
Respondent Data's own ratings are cleaned, these formula cells
automatically show the clean value too, no separate action needed.

BUT: a real, confirmed pattern exists in the actual template file - a
meaningful number of these cells (150 across both sheets, in the version
checked) have had their formula overwritten with a static, hardcoded
value at some point (someone manually pasted a value directly into the
cell). A static cell is disconnected from Respondent Data entirely - it
will never inherit a fix made there. If any such static value happens to
be messy text (e.g. "1- Out" instead of 1), cleaning Respondent Data alone
would silently miss it.

Checked directly against the June template: zero messy text values found
currently - all static-value cells found there are already clean numbers.
This module exists so that check is repeatable every month going forward,
not a one-time manual look.
"""
import re
import openpyxl
from openpyxl.utils import column_index_from_string
from clean_ratings import clean_rating_value

# Column positions confirmed directly from the real template's header
# structure (row 1/2): each brand gets a 5-column block
# (Aware, Purchase, Consider, Rating, Flag) starting at D, I, N - so the
# Rating sub-column for each brand is at G, L, Q respectively. Verified
# identical structure in both AE(DRY) and AE(WET), just different brand
# names in row 1 (Braun/Philips/Panasonic vs Gillette/Schick/Kai).
RATING_COLUMN_LETTERS = ["G", "L", "Q"]


def find_ae_rating_issues(working_path, sheets=("AE(DRY)", "AE(WET)")):
    """Scan every Rating sub-column in the given sheets for cells that are
    NOT formulas and NOT already-clean numbers. Returns a dict:
    {sheet_name: [(cell_ref, brand_name, raw_value, cleaned_value_or_None), ...]}
    An empty list for a sheet means every static value found there was
    already clean - nothing to fix. A cell appearing in the results with
    cleaned_value=None means it couldn't be parsed at all (flag for manual
    review, same as clean_ratings.py's "unparseable" case)."""
    wb = openpyxl.load_workbook(working_path, data_only=False, read_only=True)
    issues = {}

    for sheet_name in sheets:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        row1 = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))

        sheet_issues = []
        for row in ws.iter_rows(min_row=3, values_only=False):
            if row[0].value is None:  # column A empty -> end of data
                break
            for col_letter in RATING_COLUMN_LETTERS:
                col_idx = column_index_from_string(col_letter)
                cell = row[col_idx - 1]
                v = cell.value
                is_formula = hasattr(v, "text") or (isinstance(v, str) and v.startswith("="))
                if is_formula or v is None:
                    continue
                if isinstance(v, (int, float)):
                    continue  # already clean, not an issue
                # A static string that isn't a formula - this is the case
                # that matters: check if it needs cleaning
                cleaned, changed = clean_rating_value(v)
                brand = row1[col_idx - 1 - 3]  # brand name is 3 cols before Rating in its block
                if changed == "unparseable":
                    sheet_issues.append((cell.coordinate, brand, v, None))
                elif changed:
                    sheet_issues.append((cell.coordinate, brand, v, cleaned))
        issues[sheet_name] = sheet_issues

    total = sum(len(v) for v in issues.values())
    print(f"\nAE(DRY)/AE(WET) Rating column check: {total} static messy-text value(s) found.")
    for sheet_name, sheet_issues in issues.items():
        if sheet_issues:
            print(f"  {sheet_name}: {len(sheet_issues)} cell(s) need cleaning")
            for cell_ref, brand, raw, cleaned in sheet_issues[:10]:
                status = f"-> {cleaned}" if cleaned is not None else "-> UNPARSEABLE, review manually"
                print(f"    {cell_ref} ({brand}): {raw!r} {status}")
        else:
            print(f"  {sheet_name}: clean - no static messy-text values found")

    return issues


if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else "report_template.xlsx"
    find_ae_rating_issues(path)
