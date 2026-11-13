"""
Adapts the report template to whatever Respondent Data layout the flat file
arrives in - columns added, removed or moved - so the monthly pipeline never
depends on a fixed column count or fixed column letters.

WHY THIS EXISTS
    The scorecard / AE / Worth it formulas read Respondent Data by COLUMN
    LETTER (e.g. 'Respondent Data'!$BY$11:$BY$1048576 = "Braun - ever seen").
    The pipeline writes only data rows (row 11 down) and never touches the
    template's header row 10 or its formulas. So when the client's
    questionnaire changes, every formula keeps reading the OLD letter - a
    different question - and the report calculates without a single error.
    run_monthly_pipeline_JP.py now calls adapt_template() on every run (it is
    idempotent: a template that already matches comes out unchanged), and
    uses the adapted copy. The template file in the Report folder is never
    modified.

WHAT adapt_template() DOES (XML level - every other part of the .xlsx is
copied through byte-for-byte, same reason as safe_xlsx_writer.py)
    1. Column remap: every 'Respondent Data'!$COL$.. reference in the six
       report sheets moves to the column whose HEADER TEXT matches in the new
       layout (identical repeated headers are matched in order). A renamed
       header can be bridged with Column_Aliases.csv (old header,new header).
       STOPS, naming the columns, if a header the formulas read is missing,
       ambiguous or used in a form it cannot rewrite - that needs a human.
    2. Respondent Data header row 10 is rewritten to the new layout (with
       'Age absolute' at F); the sheet width follows the flat file, however
       many columns it has.
    3. Row limits are sized from the data, never reduced: Respondent Data end
       row -> 1048576; Worth it tabs' table bounds -> 1.5x the data (min
       10,000); AE tabs -> largest month + 100 respondents (min 700); the
       scorecard ranges on the AE tabs follow.
    4. A month typed into a Worth it formula (="April") -> =$T$1, so it
       follows the report month stamped by set_report_month.py.
    5. AE(DRY)/AE(WET): one formula row per respondent of the report month
       (column A is a FILTER() spill). The formula block is copied down to the
       new last row (column span detected from row 3), and hand-typed values
       found inside the block (a 10 pasted into a Rating cell) are replaced by
       the row's formula - row r is a different respondent every month.
    6. Returns a layout report (moved / added / removed columns) for the log.

Sheets deliberately left alone: COPY, Formula Column Map (client's new
ad-tracking work, out of scope for now) and the hidden Power-Query
categorisation tabs (Brand to Attributes / Brand and Attributes - they use
positional VLOOKUP indexes and are not read by any scorecard).

CLI:  python remap_template_layout.py <template.xlsx> <flat.xlsx> <out.xlsx> [Column_Aliases.csv]
"""
import re
import csv
import math
import zipfile
import collections
import difflib
import html
import copy
import posixpath
from pathlib import Path
from lxml import etree
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
from openpyxl.formula.translate import Translator
from safe_xlsx_writer import _drop_calc_chain

REPORT_SHEETS = ["DRY SCORECARD", "AE(DRY)", "WET SCORECARD", "AE(WET)",
                 "Worth it & UBA_App", "Worth it & UBA_BR"]
WORTH_IT_SHEETS = ["Worth it & UBA_App", "Worth it & UBA_BR"]
AE_SHEETS = {"AE(DRY)": "DRY SCORECARD", "AE(WET)": "WET SCORECARD"}  # AE tab -> scorecard that reads it
RD_BOUND_NEW = 1048576          # Excel's last row
RD_HEADER_ROW = 10
RD_DATA_START_ROW = 11
AGE_INSERT_INDEX = 5            # 0-based: 'Age absolute' lands at column F
WORTH_IT_MIN_BOUND = 10000
AE_MIN_CAPACITY = 700           # respondents the AE tabs can hold
AE_HEADROOM = 100               # spare respondent rows above the largest month in the flat file
LAYOUT_DEFAULT_STYLE = 191      # header style used in the template's row 10
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
          "September", "October", "November", "December"]


# --------------------------------------------------------------------------- helpers
def _sheet_paths(zf):
    wb = zf.read("xl/workbook.xml").decode("utf8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf8")
    out = {}
    for m in re.finditer(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb):
        t = (re.search(r'Id="%s"[^>]*Target="([^"]+)"' % m.group(2), rels)
             or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % m.group(2), rels))
        tgt = t.group(1)
        out[html.unescape(m.group(1))] = tgt if tgt.startswith("xl/") else "xl/" + tgt
    return out


def _esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _norm(x):
    if x in (None, ""):
        return None
    s = re.sub(r"\s+", " ", str(x)).strip()
    # the pipeline's own column: the KR template spells it 'age absolute', the JP template 'Age absolute'
    return "Age absolute" if s.lower() == "age absolute" else s


def _shared_strings(zf):
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    t = zf.read("xl/sharedStrings.xml").decode("utf8")
    return [html.unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.DOTALL)))
            for si in re.findall(r"<si>(.*?)</si>|<si/>", t, re.DOTALL)]


def read_template_header(zf, paths, header_row=RD_HEADER_ROW):
    """Respondent Data header row of the TEMPLATE, parsed from the XML (does not
    depend on the sheet's <dimension> tag, which can be stale)."""
    text = zf.read(paths["Respondent Data"]).decode("utf8")
    m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % header_row, text, re.DOTALL)
    if not m:
        raise RuntimeError(f"Respondent Data row {header_row} not found in the template")
    sst = None
    out = {}
    for cm in re.finditer(r'<c r="([A-Z]+)%d"([^>]*?)(?:/>|>(.*?)</c>)' % header_row, m.group(1), re.DOTALL):
        col, attrs, body = CI(cm.group(1)), cm.group(2), cm.group(3) or ""
        if 't="s"' in attrs:
            v = re.search(r"<v>(\d+)</v>", body)
            if v:
                sst = sst if sst is not None else _shared_strings(zf)
                out[col] = sst[int(v.group(1))]
        elif "<is>" in body:
            out[col] = html.unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.DOTALL)))
        else:
            v = re.search(r"<v>(.*?)</v>", body)
            if v:
                out[col] = html.unescape(v.group(1))
    width = max(out) if out else 0
    return [out.get(i) for i in range(1, width + 1)]


def read_flat_header(flat_path, sheet="Respondent Data", header_row=RD_HEADER_ROW):
    import openpyxl
    wb = openpyxl.load_workbook(flat_path, read_only=True, data_only=True)
    ws = wb[sheet]
    ws.reset_dimensions()   # do not trust the file's own <dimension> tag for the width
    return list(next(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True)))


def new_layout_header(flat_path):
    """The header the pipeline will produce: flat header with 'Age absolute'
    inserted at F, up to the flat file's last populated column."""
    fh = read_flat_header(flat_path)
    last = max((i for i, h in enumerate(fh) if _norm(h) is not None), default=-1) + 1
    if any(_norm(h) == "Age absolute" for h in fh):
        raise RuntimeError("The flat file already has an 'Age absolute' column - the pipeline inserts its own at "
                           "column F. Check with the client before running.")
    return fh[:AGE_INSERT_INDEX] + ["Age absolute"] + fh[AGE_INSERT_INDEX:last]


def flat_sizes(flat_path):
    """(largest number of respondents in any one month, {Worth it sheet: rows in table one}, Respondent Data row count)."""
    import openpyxl
    import datetime
    wb = openpyxl.load_workbook(flat_path, read_only=True, data_only=True)
    ws = wb["Respondent Data"]
    ws.reset_dimensions()
    per_month = collections.Counter()
    n_rows = 0
    for (rid, _, _, mm) in ws.iter_rows(min_row=RD_DATA_START_ROW, max_col=4, values_only=True):
        if rid is None:
            continue
        n_rows += 1
        per_month[(mm.year, mm.month) if isinstance(mm, datetime.datetime) else str(mm)] += 1
    worth = {}
    for sh in WORTH_IT_SHEETS:
        w = wb[sh]
        w.reset_dimensions()
        worth[sh] = sum(1 for (a,) in w.iter_rows(min_row=2, max_col=1, values_only=True) if a is not None)
    return (max(per_month.values()) if per_month else 0), worth, n_rows


def load_aliases(csv_path):
    """Optional Column_Aliases.csv: two columns, old header text -> new header text, for a question
    the client re-worded. One pair per line; a first line starting with 'old' is treated as a title."""
    aliases = {}
    p = Path(csv_path) if csv_path else None
    if p and p.exists():
        with open(p, newline="", encoding="utf-8-sig") as f:
            for row in csv.reader(f):
                if len(row) >= 2 and _norm(row[0]) and _norm(row[1]) and not _norm(row[0]).lower().startswith("old"):
                    aliases[_norm(row[0])] = _norm(row[1])
    return aliases


# --------------------------------------------------------------------------- column mapping
def referenced_columns(zf, paths):
    cols = set()
    for name in REPORT_SHEETS:
        t = zf.read(paths[name]).decode("utf8")
        cols |= set(re.findall(r"(?:'|&apos;)Respondent Data(?:'|&apos;)!\$([A-Z]{1,3})\$\d+", t))
        # multi-column ranges (the KR template's VLOOKUP(.., 'Respondent Data'!$B$11:$F$2263, 5, FALSE)): the last column too
        cols |= set(re.findall(r"(?:'|&apos;)Respondent Data(?:'|&apos;)!\$[A-Z]{1,3}\$\d+:\$([A-Z]{1,3})\$\d+", t))
    return sorted(cols, key=lambda c: (len(c), c))


def map_layouts(old_header, new_header, referenced, aliases=None):
    """Match old columns to new columns by header text (repeated identical headers are matched in order).
    Returns (full_map {old col index: new col index}, problems [str], report dict)."""
    aliases = aliases or {}
    old = [_norm(h) for h in old_header]
    new = [_norm(h) for h in new_header]
    new_pos = collections.defaultdict(list)
    for i, h in enumerate(new):
        if h:
            new_pos[h].append(i + 1)
    old_pos = collections.defaultdict(list)
    for i, h in enumerate(old):
        if h:
            old_pos[h].append(i + 1)
    full = {}
    used_new = set()
    for h, olds in old_pos.items():
        target = h if h in new_pos else aliases.get(h)
        news = new_pos.get(target, [])
        if len(olds) == len(news):
            for o, n in zip(olds, news):
                full[o] = n
                used_new.add(n)
    ref_idx = {CI(c): c for c in referenced}
    problems = []
    for idx, col in ref_idx.items():
        h = old[idx - 1] if idx - 1 < len(old) else None
        if not h:
            problems.append(f"column {col}: the template has no header there, so its question is unknown")
        elif idx not in full:
            target = aliases.get(h, h)
            n_new = len(new_pos.get(target, []))
            if n_new == 0:
                near = difflib.get_close_matches(h, [x for x in new if x], n=2, cutoff=0.6)
                hint = ("; closest in the new file: " + " | ".join(repr(x[:60]) for x in near)) if near else ""
                problems.append(f"column {col}: {h[:80]!r} is not in the new flat file{hint}")
            else:
                problems.append(f"column {col}: {h[:80]!r} appears {len(old_pos[h])}x in the template but {n_new}x in the new file")
    added = [(L(i + 1), h) for i, h in enumerate(new) if h and (i + 1) not in used_new]
    removed = [(L(o), old[o - 1], o in ref_idx) for o in range(1, len(old) + 1) if old[o - 1] and o not in full]
    report = {"moved": sum(1 for o, n in full.items() if o != n),
              "unchanged": sum(1 for o, n in full.items() if o == n),
              "added": added, "removed": removed,
              "old_width": len([h for h in old if h]), "new_width": len([h for h in new if h])}
    return full, problems, report


# --------------------------------------------------------------------------- Excel table on Respondent Data
TABLE_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _rd_table_parts(zf, rd_path):
    """Part names of the Excel tables (ListObjects) attached to the Respondent Data sheet."""
    rels = posixpath.join(posixpath.dirname(rd_path), "_rels", posixpath.basename(rd_path) + ".rels")
    if rels not in zf.namelist():
        return []
    out = []
    for m in re.finditer(r"<Relationship [^>]*>", zf.read(rels).decode("utf8")):
        if m.group(0).count("/relationships/table\"") and "TargetMode=\"External\"" not in m.group(0):
            t = re.search(r'Target="([^"]+)"', m.group(0)).group(1)
            out.append(t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(rd_path), t)))
    return out


def _rebuild_table(xml_bytes, full_map, new_header, last_row):
    """Re-define an Excel table over the new Respondent Data layout: new range, one tableColumn per header
    (name = the header text, which Excel requires to match the header cell exactly), carry each moved column's
    formatting, drop the stale April filter and the calculated-column formulas (the pipeline writes values).
    Without this Excel opens the file with 'We found a problem with some content ... Table'."""
    from lxml import etree
    q = "{%s}" % TABLE_NS
    root = etree.fromstring(xml_bytes)
    old_cols = list(root.find(q + "tableColumns"))
    inv = {n: o for o, n in full_map.items()}
    names = [str(h) for h in new_header]
    if len({x.lower() for x in names}) != len(names) or any(x == "" for x in names):
        raise RuntimeError("STOP: the flat file's Respondent Data header has duplicate or empty question names, "
                           "which an Excel table cannot hold - check the header row (row 10) of the flat file.")
    if max(len(x) for x in names) > 255:
        raise RuntimeError("STOP: a Respondent Data header is longer than 255 characters (Excel's limit for a table column name).")
    cols = root.find(q + "tableColumns")
    for c in old_cols:
        cols.remove(c)
    for j, h in enumerate(names, start=1):
        o = inv.get(j)
        if o and o <= len(old_cols):
            c = copy.deepcopy(old_cols[o - 1])
            for child in list(c):
                if child.tag == q + "calculatedColumnFormula":
                    c.remove(child)
            c.attrib.pop("totalsRowFunction", None)
        else:
            c = etree.Element(q + "tableColumn")
        c.set("id", str(j))
        c.set("name", h)
        cols.append(c)
    cols.set("count", str(len(names)))
    ref = f"A{RD_HEADER_ROW}:{L(len(names))}{max(last_row, RD_DATA_START_ROW)}"
    root.set("ref", ref)
    af = root.find(q + "autoFilter")
    if af is not None:
        af.set("ref", ref)
        for fc in list(af):
            af.remove(fc)          # the template was left filtered to one old month
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), ref


# --------------------------------------------------------------------------- AE rows
def _extend_ae_rows(xml_bytes, last_row, template_row=3):
    """Copy the per-respondent formula block of the AE tab down to `last_row`.
    The block = every column from B to the last contiguous formula in row 3. Existing complete rows
    are kept (hand-typed values in them are replaced by the row's formula); rows after the last complete
    row are rebuilt. Column A is left alone: it is the FILTER() spill range."""
    tree = etree.fromstring(xml_bytes)
    ns = {"m": MAIN_NS}
    sd = tree.find("m:sheetData", ns)
    rows = {int(r.get("r")): r for r in sd}

    def cells_of(rn):
        return {CI(re.match(r"[A-Z]+", c.get("r")).group(0)): c for c in rows[rn]} if rn in rows else {}

    row3 = cells_of(template_row)
    c1, c2 = 2, 2
    while c2 + 1 in row3 and row3[c2 + 1].find("m:f", ns) is not None:
        c2 += 1
    if row3.get(2) is None or row3[2].find("m:f", ns) is None:
        raise ValueError(f"AE tab row {template_row} column B has no formula - the AE block cannot be located")
    cols = range(c1, c2 + 1)

    def formula_of(cell):
        f = cell.find("m:f", ns)
        return f.text if f is not None and (f.text or "").strip() else None

    base = {}   # column -> (formula text, origin cell)
    for col in cols:
        for rn in sorted(r for r in rows if r >= template_row):
            c = cells_of(rn).get(col)
            if c is not None and formula_of(c):
                base[col] = ("=" + formula_of(c), f"{L(col)}{rn}")
                break
        else:
            raise ValueError(f"AE tab: no formula found anywhere in column {L(col)} to copy down")

    def tr(col, rn):
        text, origin = base[col]
        return Translator(text, origin=origin).translate_formula(f"{L(col)}{rn}")[1:]

    # last row of the contiguous run of complete rows starting at template_row
    end = template_row
    while end + 1 in rows and all(col in cells_of(end + 1) for col in cols):
        end += 1
    restored = []
    for rn in range(template_row + 1, end + 1):
        have = cells_of(rn)
        for col in cols:
            c = have[col]
            if c.find("m:f", ns) is None:
                old_val = c.findtext("m:v", default="", namespaces=ns) or "".join(c.itertext())
                for ch in list(c):
                    c.remove(ch)
                if "t" in c.attrib:
                    del c.attrib["t"]
                etree.SubElement(c, f"{{{MAIN_NS}}}f").text = tr(col, rn)
                restored.append((f"{L(col)}{rn}", old_val))
    style = {col: c.get("s") for col, c in cells_of(end).items() if col in cols}
    row_attrs = {k: v for k, v in rows[end].attrib.items() if k not in ("r", "spans")}
    # rows after the complete run are rebuilt in full (they may be partial leftovers)
    for rn in [r for r in rows if r > end]:
        sd.remove(rows[rn])
    for rn in range(end + 1, last_row + 1):
        row = etree.SubElement(sd, f"{{{MAIN_NS}}}row", r=str(rn), **row_attrs)
        for col in cols:
            cell = etree.SubElement(row, f"{{{MAIN_NS}}}c", r=f"{L(col)}{rn}")
            if style.get(col):
                cell.set("s", style[col])
            etree.SubElement(cell, f"{{{MAIN_NS}}}f").text = tr(col, rn)
    dim = tree.find("m:dimension", ns)
    if dim is not None:
        m = re.match(r"([A-Z]+\d+):([A-Z]+)(\d+)", dim.get("ref", ""))
        if m:
            dim.set("ref", f"{m.group(1)}:{m.group(2)}{max(int(m.group(3)), last_row, end)}")
    return (etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True),
            restored, max(end, last_row))


# --------------------------------------------------------------------------- the rewrite
def _max_bound(text, pattern):
    vals = [int(x) for x in re.findall(pattern, text)]
    return max(vals) if vals else 0


def remap_template(template_path, out_path, full_map, mapping_letters, new_header, worth_it_bound, ae_last_row,
                   verbose=True, rd_last_row=None):
    """mapping_letters: {old col letter: new col letter} for every column the report formulas read.
    full_map: {old col index: new col index} for all header columns (used to carry column styles over)."""
    zin = zipfile.ZipFile(template_path)
    paths = _sheet_paths(zin)
    items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    path_to_name = {paths[n]: n for n in REPORT_SHEETS}
    rd_path = paths["Respondent Data"]
    table_parts = set(_rd_table_parts(zin, rd_path))
    zin.close()

    ref_re = re.compile(r"((?:'|&apos;)Respondent Data(?:'|&apos;)!)\$([A-Z]{1,3})\$(\d+):\$([A-Z]{1,3})\$(\d+)")
    any_rd_re = re.compile(r"(?:'|&apos;)Respondent Data(?:'|&apos;)!")
    local_re = re.compile(r"(\$[A-Z]{1,2}\$2:\$[A-Z]{1,2}\$)(\d+)")
    month_re = re.compile(r'(\$[A-Z]{1,2}\$2:\$[A-Z]{1,2}\$\d+)=(?:"|&quot;)(%s)(?:"|&quot;)' % "|".join(MONTHS))
    stats = collections.OrderedDict()
    old_row11, old_row10 = {}, {}
    for item, data in items:
        if item.filename == rd_path:
            text = data.decode("utf8")
            for rown, dest in ((RD_HEADER_ROW, old_row10), (RD_DATA_START_ROW, old_row11)):
                m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % rown, text, re.DOTALL)
                if m:
                    for cm in re.finditer(r'<c r="([A-Z]+)%d"([^>]*?)(?:/>|>)' % rown, m.group(1)):
                        sm = re.search(r's="(\d+)"', cm.group(2))
                        if sm:
                            dest[CI(cm.group(1))] = int(sm.group(1))
            break
    inv = {n: o for o, n in full_map.items()}   # new col index -> old col index
    # AE rows/ranges never shrink: use the larger of the requested size and what the scorecards already read
    ae_target = ae_last_row
    for ae, sc in AE_SHEETS.items():
        t = next((d for i, d in items if i.filename == paths[sc]), b"").decode("utf8")
        ae_target = max(ae_target, _max_bound(t, r"AE\(%s\)(?:'|&apos;)!\$[A-Z]{1,2}\$3:\$[A-Z]{1,2}\$(\d+)" % ae[3:6]))
    out_items = []
    for item, data in items:
        name = path_to_name.get(item.filename)
        if name:
            text = data.decode("utf8")
            n_refs = len(any_rd_re.findall(text))
            bad = []

            def sub(m):
                c1, c2 = m.group(2), m.group(4)
                if c1 != c2:
                    # A block of columns read as one table (VLOOKUP(x, 'Respondent Data'!$B$11:$F$2263, 5, FALSE)): it can
                    # only be kept if every column of the block keeps its place relative to the first one in the new
                    # layout - then the lookup's column index stays right. Otherwise a human must look.
                    i1, i2 = CI(c1), CI(c2)
                    if i1 < i2 and all(full_map.get(i) is not None and full_map[i] == full_map.get(i1, -99) + (i - i1)
                                       for i in range(i1, i2 + 1)):
                        return f"{m.group(1)}${L(full_map[i1])}${m.group(3)}:${L(full_map[i2])}${RD_BOUND_NEW}"
                    bad.append(m.group(0)[:90])
                    return m.group(0)
                if c1 not in mapping_letters:
                    bad.append(m.group(0)[:90])
                    return m.group(0)
                return f"{m.group(1)}${mapping_letters[c1]}${m.group(3)}:${mapping_letters[c1]}${RD_BOUND_NEW}"

            text, n_sub = ref_re.subn(sub, text)
            if bad or n_refs != n_sub:
                raise RuntimeError(
                    f"STOP: '{name}' reads Respondent Data in a form this tool cannot safely rewrite "
                    f"({n_refs - n_sub + len(bad)} reference(s), e.g. {bad[:2] or 'single-cell / relative reference'}). "
                    f"Only whole-column absolute ranges like 'Respondent Data'!$BY$11:$BY$1048576 are supported - "
                    f"check the template before running.")
            st = {"respondent_data_refs": n_sub}
            if name in WORTH_IT_SHEETS:
                tgt = max(worth_it_bound, _max_bound(text, r"\$[A-Z]{1,2}\$2:\$[A-Z]{1,2}\$(\d{4,})"))
                text, nb = local_re.subn(lambda m: m.group(1) + str(tgt) if 1000 <= int(m.group(2)) < tgt else m.group(0), text)
                text, nm = month_re.subn(r"\g<1>=$T$1", text)
                # The age-band formulas look every table row up in Respondent Data with MATCH. A table row that is
                # empty (the bound is larger than the table) gives MATCH -> #N/A, and one #N/A makes the whole
                # SUMPRODUCT #N/A. IFERROR(..,0) turns an empty/unmatched row into age 0, which no age band
                # includes, so the result is the same as for an exactly sized range - at any table length.
                idx_re = re.compile(r"(?<!IFERROR\()INDEX\(((?:'|&apos;)Respondent Data(?:'|&apos;)!\$[A-Z]{1,3}\$11:\$[A-Z]{1,3}\$\d+),"
                                    r"(MATCH\(\$[A-Z]{1,2}\$2:\$[A-Z]{1,2}\$\d+,(?:'|&apos;)Respondent Data(?:'|&apos;)!\$[A-Z]{1,3}\$11:\$[A-Z]{1,3}\$\d+,0\))\)")
                text, n_idx = idx_re.subn(r"IFERROR(INDEX(\1,\2),0)", text)
                st.update(table_bound=tgt, typed_month_replaced=nm, blank_row_safe_age_lookups=n_idx)
            for ae, sc in AE_SHEETS.items():
                if name == sc:
                    rng = re.compile(r"((?:'|&apos;)%s(?:'|&apos;)!\$[A-Z]{1,2}\$3:\$[A-Z]{1,2}\$)(\d+)" % re.escape(ae))
                    text = rng.sub(lambda m: m.group(1) + str(ae_target) if int(m.group(2)) < ae_target else m.group(0), text)
                    st["ae_range_end"] = ae_target
                    # The AE tabs now hold more rows than there are respondents, and the unused rows show "" in the
                    # Age column. VALUE("") is #VALUE!, which would turn every age-band formula built on
                    # VALUE('AE(..)'!$Y$3:$Y$n) into #VALUE!. Wrapped in IFERROR(..,0) the blank rows become 0,
                    # and they are excluded anyway (their classification flag, column C, is 0).
                    val_re = re.compile(r"(?<!IFERROR\()VALUE\(((?:'|&apos;)%s(?:'|&apos;)!\$[A-Z]{1,2}\$3:\$[A-Z]{1,2}\$\d+)\)" % re.escape(ae))
                    text, n_val = val_re.subn(r"IFERROR(VALUE(\1),0)", text)
                    if n_val:
                        st["value_wrapped_for_blank_rows"] = n_val
            data = text.encode("utf8")
            if name in AE_SHEETS:
                data, restored, last = _extend_ae_rows(data, ae_target)
                st.update(ae_formula_rows_to=last, typed_values_replaced_by_formula=len(restored),
                          typed_values_were=sorted({v for _, v in restored}))
            stats[name] = st
        elif item.filename == rd_path:
            text = data.decode("utf8")
            n = len(new_header)

            def hdr_style(j):
                o = inv.get(j)
                return old_row10.get(o) or old_row10.get(j) or LAYOUT_DEFAULT_STYLE

            def row_style(j):
                o = inv.get(j)
                return old_row11.get(o, old_row11.get(j))
            cells = [f'<c r="{L(j)}{RD_HEADER_ROW}" s="{hdr_style(j)}" t="inlineStr"><is><t xml:space="preserve">{_esc(str(h))}</t></is></c>'
                     for j, h in enumerate(new_header, start=1) if h not in (None, "")]
            m = re.search(r'<row r="%d"([^>]*)>(.*?)</row>' % RD_HEADER_ROW, text, re.DOTALL)
            attrs = re.sub(r'spans="[^"]*"', f'spans="1:{n}"', m.group(1))
            text = text[:m.start()] + f'<row r="{RD_HEADER_ROW}"{attrs}>{"".join(cells)}</row>' + text[m.end():]
            m = re.search(r'<row r="%d"([^>]*)>(.*?)</row>' % RD_DATA_START_ROW, text, re.DOTALL)
            attrs = re.sub(r'spans="[^"]*"', f'spans="1:{n}"', m.group(1))
            blanks = "".join(f'<c r="{L(j)}{RD_DATA_START_ROW}"' + (f' s="{row_style(j)}"' if row_style(j) else "") + "/>"
                             for j in range(1, n + 1))
            text = text[:m.start()] + f'<row r="{RD_DATA_START_ROW}"{attrs}>{blanks}</row>' + text[m.end():]
            dm = re.search(r'<dimension ref="([A-Z]+\d+):([A-Z]+)(\d+)"/>', text)
            if dm:
                text = text[:dm.start()] + f'<dimension ref="{dm.group(1)}:{L(n)}{dm.group(3)}"/>' + text[dm.end():]
            stats.setdefault("Respondent Data", {})["columns"] = n
            data = text.encode("utf8")
        elif item.filename in table_parts:
            data, tref = _rebuild_table(data, full_map, new_header, rd_last_row if rd_last_row else RD_DATA_START_ROW)
            stats.setdefault("Respondent Data", {})["table"] = tref
        out_items.append((item, data))
    out_items = _drop_calc_chain(out_items)   # formulas were moved/added; Excel rebuilds the chain
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for item, data in out_items:
            z.writestr(item, data)
    if verbose:
        for k, v in stats.items():
            print(f"    {k}: {v}")
    return stats


def adapt_template(template_path, flat_path, out_path, aliases=None, verbose=True):
    """Write an adapted copy of the template to out_path and return the layout report.
    Raises RuntimeError (naming the columns) if the report formulas read something the new flat file lacks."""
    zin = zipfile.ZipFile(template_path)
    paths = _sheet_paths(zin)
    old_header = read_template_header(zin, paths)
    referenced = referenced_columns(zin, paths)
    zin.close()
    new_header = new_layout_header(flat_path)
    full, problems, report = map_layouts(old_header, new_header, referenced, aliases)
    if problems:
        raise RuntimeError(
            "STOP: the report formulas read questions the new flat file no longer provides (or provides ambiguously):\n  "
            + "\n  ".join(problems)
            + "\nIf the client only re-worded a question, add 'old header,new header' to Column_Aliases.csv. "
              "If a question was dropped, the report needs a decision from the client before it can be built.")
    letters = {c: L(full[CI(c)]) for c in referenced}
    max_month, worth_rows, rd_rows = flat_sizes(flat_path)
    ae_capacity = max(AE_MIN_CAPACITY, math.ceil((max_month + AE_HEADROOM) / 100) * 100)
    worth_bound = max(WORTH_IT_MIN_BOUND, math.ceil(max(worth_rows.values(), default=0) * 1.5 / 1000) * 1000)
    report.update(referenced=len(referenced), ae_capacity=ae_capacity, worth_bound=worth_bound,
                  largest_month=max_month, worth_rows=worth_rows,
                  table_last_row=RD_HEADER_ROW + max(rd_rows, 1))
    if verbose:
        print(f"  Layout: template header {report['old_width']} columns -> flat file {report['new_width']} columns "
              f"(incl. Age absolute); {len(referenced)} columns read by the formulas, "
              f"{sum(1 for c in referenced if letters[c] != c)} of them moved")
        print(f"  Sizing: largest month {max_month} respondents -> AE tabs hold {ae_capacity}; "
              f"Worth it tables {worth_rows} rows -> formulas read to row {worth_bound}")
    remap_template(template_path, out_path, full, letters, new_header, worth_bound, ae_capacity + 2, verbose=verbose,
                   rd_last_row=report["table_last_row"])
    return report


def format_layout_report(report):
    lines = [f"Template header: {report['old_width']} columns  ->  flat file: {report['new_width']} columns (incl. Age absolute)",
             f"Columns in the same position: {report['unchanged']}   moved: {report['moved']}",
             f"Columns the report formulas read: {report['referenced']}"]
    lines.append(f"\nNEW in the flat file (not in the template; not used by any formula): {len(report['added'])}")
    lines += [f"  {c}  {h[:110]}" for c, h in report["added"][:60]]
    lines.append(f"\nREMOVED from the flat file (were in the template): {len(report['removed'])}")
    lines += [f"  {c}  {h[:110]}{'   <-- READ BY FORMULAS' if used else ''}" for c, h, used in report["removed"][:60]]
    return "\n".join(lines)


if __name__ == "__main__":
    import sys
    if len(sys.argv) not in (4, 5):
        print(__doc__.split("CLI:")[1].strip())
        sys.exit(1)
    rep = adapt_template(sys.argv[1], sys.argv[2], sys.argv[3], load_aliases(sys.argv[4] if len(sys.argv) == 5 else None))
    print("\n" + format_layout_report(rep) + f"\n\nWritten: {sys.argv[3]}")