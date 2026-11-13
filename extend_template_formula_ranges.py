"""
Extends hardcoded Respondent Data row bounds in the JP report template's
formulas - built specifically to address the risk found in
check_formula_headroom(): AE(DRY) and DRY SCORECARD's formulas were bound
to row 3000/4000, while Respondent Data already reaches row 2863 and grows
by roughly 300 rows every month. Left alone, this would have silently
excluded the newest respondents from those tabs' calculations within a
month or two, with no error - just wrong numbers.

WHAT THIS DOES: finds every formula in a given sheet that references
'Respondent Data'!$COL$11:$COL$<bound> and replaces ONLY the bound number,
leaving every other part of the formula - column letters, function
structure, unrelated formulas entirely - untouched. Confirmed via exact
occurrence-count matching (1,920 in DRY SCORECARD, 10,888 in AE(DRY) -
matching a full scan of the file exactly) and byte-level comparison of
every other sheet before use.

WHY EXCEL'S ACTUAL MAXIMUM (1,048,576) AND NOT ANOTHER ROUND NUMBER: the
client explicitly asked for this to have no practical limit, rather than
extend the deadline and eventually hit it again. 1,048,576 (2^20) is the
real ceiling for .xlsx/.xlsm row count - there's no larger number to
choose, so this genuinely never needs revisiting. The one honest trade-off:
formulas over a much larger declared range may recalculate marginally
slower than a tightly-bounded one, though WET SCORECARD/AE(WET) already
run fine at a bound of 10,000 with no reported issue, suggesting this
stays negligible in practice - not something verified with real Excel
recalculation timing here, since this environment has none available.

USAGE:
    python extend_template_formula_ranges.py <template_path> <output_path>

Run this ONCE against the current template, review the output file
carefully (see the validation checklist printed at the end), then replace
the live template with it once you're satisfied. The original file is
never modified in place - always written to a separate output path.
"""
import sys
import shutil
import zipfile
import re

# sheet internal path -> (current bound to replace, sheet display name)
# Only sheets confirmed to need extension - WET SCORECARD, AE(WET), and
# Worth it & UBA_App/BR were confirmed already at 6000-10000 and were left
# untouched in the earlier 6000-bound version of this script. Since the
# client asked for no limit at all, those are now included too, extended
# from whatever they're currently at up to the same true maximum - keeping
# every sheet consistent rather than leaving some at a smaller bound.
SHEETS_TO_EXTEND = {
    "xl/worksheets/sheet2.xml": (4000, "DRY SCORECARD"),
    "xl/worksheets/sheet3.xml": (3000, "AE(DRY)"),
    "xl/worksheets/sheet6.xml": (10000, "WET SCORECARD"),
    "xl/worksheets/sheet7.xml": (10000, "AE(WET)"),
    "xl/worksheets/sheet4.xml": (6000, "Worth it & UBA_App"),
    "xl/worksheets/sheet8.xml": (10000, "Worth it & UBA_BR"),
}
NEW_BOUND = 1048576  # Excel's actual maximum row count (2^20) for .xlsx/.xlsm


def extend_formula_ranges(template_path, output_path, new_bound=NEW_BOUND):
    zf_in = zipfile.ZipFile(template_path, "r")
    items = [(item, zf_in.read(item.filename)) for item in zf_in.infolist()]
    zf_in.close()

    replacement_counts = {}
    new_items = []
    for item, data in items:
        if item.filename in SHEETS_TO_EXTEND:
            old_bound, sheet_name = SHEETS_TO_EXTEND[item.filename]
            pattern = re.compile(rf"('Respondent Data'!\$[A-Z]+\$11:\$[A-Z]+\$){old_bound}\b")
            text = data.decode("utf-8")
            new_text, count = pattern.subn(rf"\g<1>{new_bound}", text)
            replacement_counts[sheet_name] = (count, old_bound)
            data = new_text.encode("utf-8")
        new_items.append((item, data))

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as out:
        for item, data in new_items:
            out.writestr(item, data)

    print(f"Extended formula ranges in {output_path}:")
    for sheet_name, (count, old_bound) in replacement_counts.items():
        print(f"  {sheet_name}: {count} formula(s) changed from bound row {old_bound} to row {new_bound}")

    print(
        f"\nBefore trusting this file:\n"
        f"  1. Open it in Excel - confirm no repair dialog appears\n"
        f"  2. Spot-check a few formulas in DRY SCORECARD and AE(DRY) - the "
        f"row bound should now read {new_bound}, everything else in the "
        f"formula should look exactly as before\n"
        f"  3. Best test: run the full pipeline against a KNOWN month "
        f"(e.g. June) using this as the template, and confirm every number "
        f"still matches what's already been validated - proving the range "
        f"extension changes headroom, not results\n"
        f"  4. Only then replace the live template with this file - back up "
        f"the original first, given how much depends on it"
    )
    return replacement_counts


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python extend_template_formula_ranges.py <template_path> <output_path>")
        sys.exit(1)
    extend_formula_ranges(sys.argv[1], sys.argv[2])
