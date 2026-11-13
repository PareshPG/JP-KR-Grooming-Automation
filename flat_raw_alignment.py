"""Cross-check of the client's flat file against the raw data file, and (opt-in) repair of a displaced column block.

WHY: the flat file is made by the client, by hand, from the raw export. On 27 Oct 2026 the KR September flat file had its
"ever seen or heard of" brand block (AA:AH) pasted one column to the right of its headers for the 500 September rows:
'Gillette' empty, 'Schick' holding Gillette's answers, 'Dorco' holding Schick's, ... Every earlier month was right, and the
report calculated without a single error - it just reported awareness of the wrong brands.

HOW: for the rows of the raw file's month, each flat column is matched to the raw column holding exactly the same answers
(per respondent). A flat column whose header names one brand while the raw column it copies names another is a mismatch
(raw brand names come from market_config.<market>.raw_brand_names). The block of neighbouring columns that moves together
is the displaced block. The default is to STOP with an explanation; with repair=True the block is rewritten from the raw
file (matched on Internal Respondent Id) in memory - the client's file is never modified.

Uses openpyxl only (no pandas / numpy).
"""
import re
import collections
import openpyxl
from openpyxl.utils import get_column_letter as L

ID_HEADER = "Internal Respondent Id"


def _norm(v):
    if v is None or (isinstance(v, float) and v != v):
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool) and float(v).is_integer():
        return str(int(v))
    s = str(v).strip()
    return s or None


def read_raw(raw_path, sheet_name="Sheet1"):
    """(header list, {id: row list}) of the raw file."""
    wb = openpyxl.load_workbook(raw_path, read_only=True, data_only=True)
    ws = wb[sheet_name]
    ws.reset_dimensions()
    rows = ws.iter_rows(values_only=True)
    header = list(next(rows))
    i_id = header.index(ID_HEADER)
    by_id = {}
    for r in rows:
        if r[i_id] is None:
            continue
        r = list(r) + [None] * (len(header) - len(r))      # read-only mode drops trailing empty cells
        try:
            by_id[int(float(r[i_id]))] = r
        except (TypeError, ValueError):
            by_id[r[i_id]] = r
    return header, by_id


def _brand_of(text, brand_map):
    t = str(text or "").strip()
    for prefix, english in brand_map.items():          # a raw header starts with the local-language name,
        if t.startswith(prefix):                       # a flat header with the English one
            return english
    for english in set(brand_map.values()):
        if t.startswith(english):
            return english
    return None


def find_displacement(flat_header, flat_rows, raw_header, raw_by_id, brand_map):
    """flat_rows: list of tuples (all rows of the flat file). Returns a list of findings, each
    {flat_first, flat_last, shift, raw_first, mismatches:[(flat col letter, flat brand, raw brand)]} (empty list = aligned)."""
    ids = [r[1] for r in flat_rows if r[1] in raw_by_id]
    if len(ids) < 20:
        return []
    id_set = set(ids)
    rows = [r for r in flat_rows if r[1] in id_set]
    # raw column key = its answers in the flat rows' respondent order
    rawkey = collections.defaultdict(list)
    for j in range(len(raw_header)):
        if raw_header[j] == ID_HEADER:
            continue
        rawkey[tuple(_norm(raw_by_id[r[1]][j]) for r in rows)].append(j)
    match = {}            # flat col (0-based) -> raw col (0-based)
    for j, h in enumerate(flat_header):
        if h in (None, ""):
            continue
        s = tuple(_norm(r[j] if j < len(r) else None) for r in rows)
        if len(set(s)) < 2:
            continue                                  # blank or constant: cannot be matched to one raw column
        cands = rawkey.get(s)
        if cands:
            match[j] = cands[0]
    if not match:
        return []
    # runs of neighbouring flat columns that sit at the same distance from their raw columns
    runs = []
    for j in sorted(match):
        d = j - match[j]
        if runs and runs[-1]["d"] == d and runs[-1]["b"] == j - 1:
            runs[-1]["b"] = j
            runs[-1]["n"] += 1
        else:
            runs.append({"d": d, "a": j, "b": j, "n": 1})
    main = max(runs, key=lambda r: r["n"])["d"]
    out = []
    for r in runs:
        mism = []
        for j in range(r["a"], r["b"] + 1):
            if j in match:
                fb, rb = _brand_of(flat_header[j], brand_map), _brand_of(raw_header[match[j]], brand_map)
                if fb and rb and fb != rb:
                    mism.append((L(j + 1), fb, rb))
        if mism:
            shift = r["d"] - main
            out.append({"flat_first": r["a"], "flat_last": r["b"], "shift": shift, "raw_first": match[r["a"]], "mismatches": mism,
                        "main_offset": main, "n_matched_cols": len(match)})
    return out


def describe(findings, flat_header):
    lines = []
    for f in findings:
        a, b, s = f["flat_first"], f["flat_last"], f["shift"]
        lines.append(f"columns {L(a + 1)}:{L(b + 1)} hold the answers of raw columns {f['raw_first'] + 1}..{f['raw_first'] + (b - a) + 1}, "
                     f"{abs(s)} column(s) {'to the right' if s > 0 else 'to the left'} of where their headers are. "
                     + "Examples: " + "; ".join(f"{c} is headed '{fb}' but holds {rb}'s answers" for c, fb, rb in f["mismatches"][:3]))
    return lines


def repair_rows(flat_rows, findings, raw_by_id, raw_header_len):
    """Rewrite each displaced block from the raw file (in memory). Returns (new rows list, number of rows touched)."""
    out = []
    touched = 0
    for r in flat_rows:
        if r[1] not in raw_by_id:
            out.append(r)
            continue
        r = list(r)
        raw = raw_by_id[r[1]]
        for f in findings:
            shift = f["shift"]
            first_true = f["flat_first"] - shift
            for k, c in enumerate(range(first_true, f["flat_last"] + 1)):
                rj = f["raw_first"] + k
                if c < len(r) and rj < len(raw):
                    r[c] = raw[rj]
        touched += 1
        out.append(tuple(r))
    return out, touched


def check_flat_against_raw(flat_header, flat_rows, raw_path, brand_map, repair=False, log=print):
    """Returns the flat rows, repaired if repair=True and a displacement was found. Raises RuntimeError on a displacement when
    repair is False."""
    if not brand_map:
        return flat_rows
    raw_header, raw_by_id = read_raw(raw_path)
    findings = find_displacement(flat_header, flat_rows, raw_header, raw_by_id, brand_map)
    if not findings:
        log("  Flat vs raw check: every brand column of this month holds the answers of the raw column with the same brand - OK.")
        return flat_rows
    detail = describe(findings, flat_header)
    if not repair:
        raise RuntimeError(
            "STOP: the flat file does not match the raw data file for this month:\n  - " + "\n  - ".join(detail) +
            "\nThe report would calculate without any error but show the wrong brands. Ask the client for a corrected flat file. "
            "To go ahead provisionally, re-run with  --repair-flat-from-raw  (the block is then rewritten from the raw file, in memory; "
            "the flat file itself is never changed) and treat the report as provisional until the client confirms.")
    new_rows, n = repair_rows(flat_rows, findings, raw_by_id, len(raw_header))
    log("  WARNING - flat file defect repaired from the raw file (--repair-flat-from-raw): " + " | ".join(detail))
    log(f"  {n} respondent rows rewritten. PROVISIONAL: confirm with the client and replace the flat file when they re-send it.")
    # prove the repair
    again = find_displacement(flat_header, new_rows, raw_header, raw_by_id, brand_map)
    if again:
        raise RuntimeError("STOP: the repair did not remove the displacement: " + " | ".join(describe(again, flat_header)))
    return new_rows