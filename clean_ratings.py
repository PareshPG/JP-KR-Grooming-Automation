"""
Step: Clean the six brand-Rating columns in Respondent Data BEFORE recalculation.

AE(DRY)!Rating and AE(WET)!Rating are live INDEX/MATCH formulas that pull
straight from these columns - whatever sits here at recalculation time is
what the report's Rating figures are built from. Two problems have been
confirmed in this data:
  1. Genuinely messy cells: a number with descriptive text attached, e.g.
     "10 - It is perfect in every way" (seen at both scale endpoints, 1 and 10).
  2. Clean numbers stored as text, e.g. "7" instead of 7 - harmless to read,
     but breaks numeric comparisons/averages downstream if left as text.

Design notes:
  - The six columns are located BY HEADER TEXT, not by hardcoded column
    letter (OX/OY/OZ/PA/PB/PC in THIS file, but not guaranteed to stay there).
    Matched on the stable substring common to all six question headers, plus
    a known brand-name prefix, so the report is self-documenting about which
    column mapped to which brand.
  - The whole column is cleaned (not just the current month), since it costs
    nothing extra and keeps the cumulative Respondent Data table honest for
    any other consumer - even though only the current month's rows actually
    feed this month's DRY/WET SCORECARD.
  - Already-clean numeric cells are left as real numbers; only cells that need
    a fix are rewritten, to avoid unnecessary style/type churn.
"""
import re
import openpyxl
from openpyxl.utils import get_column_letter
from safe_xlsx_writer import patch_workbook_fast_cells

STABLE_HEADER_SUBSTRING = "how would you rate each of the brands"
KNOWN_BRANDS = ["Braun", "Philips", "Panasonic", "Gillette", "Schick", "Kai"]
HEADER_ROW = 10          # Respondent Data's known header row
DATA_START_ROW = 11


def find_rating_columns(ws, header_row=HEADER_ROW, known_brands=KNOWN_BRANDS):
    """Locate every Rating column by header text, tagged with its brand.
    Returns (matched, unmatched) where `matched` is [(col_index, brand_name)]
    for the known brands only, and `unmatched` is [(col_index, header_text)]
    for anything that matches the rating question but isn't a known brand -
    these are reported but never written to, since silently expanding scope
    on a fuzzy text match is exactly the kind of thing that should be a human
    decision, not an automatic one. known_brands defaults to JP's brand list -
    pass KO's (see market_config.py) when processing a KO file.

    Uses sequential iter_rows (not random ws.cell() access) - read_only mode
    only accesses cells efficiently in sequential order. Measured: ~89 random
    ws.cell() calls took 15.7s; the equivalent via iter_rows was near-instant."""
    matched, unmatched = [], []
    header_values = next(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))
    for c, v in enumerate(header_values, start=1):
        if not isinstance(v, str):
            continue
        if STABLE_HEADER_SUBSTRING.lower() not in v.lower():
            continue
        brand = next((b for b in known_brands if v.strip().startswith(b)), None)
        if brand is None:
            unmatched.append((c, v))
        else:
            matched.append((c, brand))

    if not matched:
        raise ValueError(
            f"No columns matched the rating-question header "
            f"('{STABLE_HEADER_SUBSTRING}'). Header layout may have changed - "
            f"check row {header_row} of '{ws.title}' manually."
        )
    if len(matched) != len(known_brands):
        print(f"  NOTE: expected {len(known_brands)} known-brand rating columns, "
              f"found {len(matched)}: {[b for _, b in matched]}")
    for c, header in unmatched:
        print(f"  WARNING: column {get_column_letter(c)} matches the rating question "
              f"but isn't one of the known brands - NOT cleaned, review manually: "
              f"{header[:70]!r}")
    return matched, unmatched


def clean_rating_value(raw_value):
    """Extract a clean number from a raw rating cell.
    Returns (cleaned_value_or_None, was_changed: bool)."""
    if raw_value is None or raw_value == "":
        return raw_value, False
    if isinstance(raw_value, (int, float)):
        return raw_value, False  # already clean
    match = re.search(r"-?\d+(?:\.\d+)?", str(raw_value))
    if match is None:
        return raw_value, "unparseable"  # leave untouched, flag it
    cleaned = float(match.group())
    if cleaned.is_integer():
        cleaned = int(cleaned)
    return cleaned, True


def compute_rating_fixes(ws, known_brands=KNOWN_BRANDS):
    """Read-only pass over Respondent Data: find every rating cell that
    needs fixing. Returns (fixes, stats) where fixes is a list of
    (cell_ref, cleaned_value) pairs ready to write.

    Uses ONE sequential iter_rows pass restricted to the target columns'
    span, not a random ws.cell(row=, column=) call per cell - the earlier
    version did ~17,000 random accesses and was measured to project to
    roughly 50 minutes; this restructured version does the equivalent work
    in under 10 seconds."""
    rating_cols, unmatched_cols = find_rating_columns(ws, known_brands=known_brands)
    col_to_brand = dict(rating_cols)
    min_col, max_col = min(col_to_brand), max(col_to_brand)

    stats = {brand: {"total": 0, "already_clean": 0, "fixed": 0, "unparseable": []}
              for _, brand in rating_cols}
    distinct_messy_patterns = set()
    fixes = []

    for r, row_vals in enumerate(
        ws.iter_rows(min_row=DATA_START_ROW, min_col=min_col, max_col=max_col, values_only=True),
        start=DATA_START_ROW,
    ):
        for offset, raw in enumerate(row_vals):
            col = min_col + offset
            if col not in col_to_brand:
                continue
            if raw is None or raw == "":
                continue
            brand = col_to_brand[col]
            stats[brand]["total"] += 1
            cleaned, changed = clean_rating_value(raw)
            if changed == "unparseable":
                stats[brand]["unparseable"].append((r, raw))
                continue
            if changed:
                if isinstance(raw, str) and not raw.strip().replace(".", "", 1).replace("-", "", 1).isdigit():
                    distinct_messy_patterns.add(raw)
                fixes.append((f"{get_column_letter(col)}{r}", cleaned))
                stats[brand]["fixed"] += 1
            else:
                stats[brand]["already_clean"] += 1

    print(f"Rating cleanup: columns located {[(get_column_letter(c), b) for c, b in rating_cols]}")
    total_fixed = 0
    for brand, s in stats.items():
        print(f"  {brand}: {s['total']} cells | {s['already_clean']} already clean | "
              f"{s['fixed']} fixed | {len(s['unparseable'])} unparseable")
        total_fixed += s["fixed"]
        for r, raw in s["unparseable"][:5]:
            print(f"    UNPARSEABLE row {r}: {raw!r}")
    print(f"  Total cells fixed: {total_fixed}")
    if distinct_messy_patterns:
        print(f"  Distinct messy patterns found and cleaned: {sorted(distinct_messy_patterns)}")

    return fixes, stats


def clean_ratings(target_wb_path, sheet_name="Respondent Data", save_path=None, known_brands=KNOWN_BRANDS):
    """Standalone entry point: compute fixes (read-only, fast), then patch
    and save using the fast existing-cells-only path - Respondent Data has
    ~1.9 million cell elements, too large for a full DOM parse to complete
    in reasonable time, but Rating cleanup only ever rewrites cells that
    already have a value, so the fast regex-substitution path applies.
    Preserves every other part of the workbook byte-for-byte. The monthly
    pipeline combines this with the other steps instead of calling this
    separately."""
    wb = openpyxl.load_workbook(target_wb_path, read_only=True)
    fixes, stats = compute_rating_fixes(wb[sheet_name], known_brands=known_brands)

    out_path = save_path or target_wb_path
    patch_workbook_fast_cells(target_wb_path, out_path, {
        sheet_name: dict(fixes)
    })
    print(f"Saved to {out_path}")
    return stats


if __name__ == "__main__":
    clean_ratings(
        target_wb_path="JP-June-idxage-rebuilt.xlsx",   # chain onto the previous step's output
        save_path="JP-June-ratings-cleaned.xlsx",
    )
