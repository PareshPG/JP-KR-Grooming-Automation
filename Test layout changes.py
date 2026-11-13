import sys, re, zipfile, shutil, csv, traceback
sys.path.insert(0, "pipe")
from lxml import etree
from remap_template_layout import (adapt_template, load_aliases, _sheet_paths, read_template_header, REPORT_SHEETS)
from openpyxl.utils import column_index_from_string as CI
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
TPL = "sept2/JP_tpl.xlsx"
REF = re.compile(r"'Respondent Data'!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)")

def unwrap_iferror(t, heads):
    """Remove the IFERROR(<head>...,0) wrappers the adapter adds for blank table rows, so formulas can be compared
    with the original template's by meaning (other IFERROR(INDEX(..),"") uses are left alone)."""
    for head in heads:
        key = "IFERROR(" + head
        pos = 0
        while True:
            i = t.find(key, pos)
            if i < 0: break
            k = i + len("IFERROR("); depth = 1; q = False
            while depth:
                ch = t[k]
                if ch == "'": q = not q
                if not q:
                    if ch == "(": depth += 1
                    elif ch == ")": depth -= 1
                k += 1
            inner = t[i + len("IFERROR("):k - 1]
            if inner.endswith(",0"):
                t = t[:i] + inner[:-2] + t[k:]
                pos = i
            else:
                pos = i + 1
    return t

def formulas(path):
    zf = zipfile.ZipFile(path); paths = _sheet_paths(zf)
    hdr = read_template_header(zf, paths)
    out = {}
    for sh in REPORT_SHEETS:
        if sh == "Respondent Data": continue
        root = etree.fromstring(zf.read(paths[sh]))
        for c in root.iterfind(".//m:c", NS):
            f = c.find("m:f", NS)
            if f is None or not (f.text or "").strip(): continue
            def sub(m):
                a, b = m.group(1), m.group(3)
                ha = hdr[CI(a) - 1] if CI(a) - 1 < len(hdr) else "?" + a
                return f"RD[{ha}]@{m.group(2)}"
            t = unwrap_iferror(REF.sub(sub, f.text), ("INDEX(",))
            t = re.sub(r"\$([A-Z]+)\$(\d{3,})\b", r"$\1$N", t)   # local bounds
            t = re.sub(r"(RD\[[^\]]*\])@\d+", r"\1@N", t)
            t = re.sub(r"IFERROR\((VALUE\('AE\(WET\)'!\$Y\$3:\$Y\$[0-9N]+\)),0\)", r"\1", t)
            t = re.sub(r"\"?(April|September)\"?|\$T\$1", "MONTH", t)
            out[(sh, c.get("r"))] = t.replace("\r\n", "\n")
    return out, hdr

def run(name, flat, aliases_csv=None, expect_stop=False):
    print(f"\n===== {name} =====")
    out = f"stress/out_{name}.xlsx"
    al = load_aliases(aliases_csv) if aliases_csv else None
    try:
        rep = adapt_template(TPL, flat, out, al, verbose=False)
    except RuntimeError as e:
        msg = str(e)
        print("STOPPED:", msg[:700])
        print("RESULT:", "OK (stopped as expected)" if expect_stop else "UNEXPECTED STOP")
        return
    if expect_stop:
        print("RESULT: FAIL - should have stopped"); return
    print({k: (v if not isinstance(v, (list, dict)) else len(v)) for k, v in rep.items()})
    a, ha = formulas(TPL); b, hb = formulas(out)
    common = [k for k in a if k in b]
    alias_map = {}
    if al:
        alias_map = {o: n for o, n in al.items()} if isinstance(al, dict) else dict(al)
    diffs = []
    for k in common:
        x = a[k]
        for o, n in alias_map.items():
            x = x.replace(f"RD[{o}]", f"RD[{n}]")
        if x != b[k]: diffs.append(k)
    missing = [k for k in a if k not in b]
    print(f"formula cells compared {len(common)}; semantic diffs {len(diffs)}; template cells lost {len(missing)}; adapted has {len(b)-len(common)} extra")
    for k in diffs[:3]: print("  DIFF", k, "\n    A:", a[k][:200], "\n    B:", b[k][:200])
    print("RESULT:", "OK" if not diffs and not missing else "CHECK")
    # header check
    from run_monthly_pipeline_JP import validate_template_layout
    validate_template_layout(out, flat); print("validate_template_layout passed")

if __name__ == "__main__":
    run("v1_insert3", "stress/v1_insert3.xlsx")
    run("v2_drop_unref", "stress/v2_drop_unref.xlsx")
    run("v3_rename_noalias", "stress/v3_rename.xlsx", expect_stop=True)
    run("v3_rename_alias", "stress/v3_rename.xlsx", "stress/Column_Aliases.csv")
    run("v4_drop_ref", "stress/v4_drop_ref.xlsx", expect_stop=True)
    run("v5_bigmonth", "stress/v5_bigmonth.xlsx")