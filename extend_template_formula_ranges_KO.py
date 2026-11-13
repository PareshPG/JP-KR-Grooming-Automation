"""
Extends hardcoded Respondent Data row bounds in the KO report template's
formulas to Excel's true maximum - same fix already applied to JP's
template at the client's explicit request ("without any limit"), applied
here because AE(WET) was found to be ALREADY EXCEEDED (bound row 2263,
actual data already reaches row 2421) - not a future risk, an active one.

WHAT THIS DOES: identical mechanism to extend_template_formula_ranges.py
(JP) - finds every formula referencing 'Respondent Data'!$COL$11:$COL$<bound>
in the six affected sheets and replaces only the bound number, leaving
everything else untouched.

USAGE:
    python extend_template_formula_ranges_KO.py <template_path> <output_path>
"""
import sys
import zipfile
import re

# sheet internal path -> (current bound to replace, sheet display name)
# All six confirmed directly from the real KO template before use.
SHEETS_TO_EXTEND = {
    "xl/worksheets/sheet2.xml": (10000, "WET SCORECARD"),
    "xl/worksheets/sheet3.xml": (10000, "DRY SCORECARD"),
    "xl/worksheets/sheet4.xml": (10000, "AE(DRY)"),
    "xl/worksheets/sheet5.xml": (2263, "AE(WET)"),
    "xl/worksheets/sheet8.xml": (10000, "Worth it & UBA_BR"),
    "xl/worksheets/sheet9.xml": (10000, "Worth it & UBA_App"),
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
        f"  2. Best test: run the full pipeline against a KNOWN month (June) using this "
        f"as the template, and confirm every number still matches what's already been "
        f"validated\n"
        f"  3. Only then replace the live template with this file - back up the original first"
    )
    return replacement_counts


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python extend_template_formula_ranges_KO.py <template_path> <output_path>")
        sys.exit(1)
    extend_formula_ranges(sys.argv[1], sys.argv[2])
