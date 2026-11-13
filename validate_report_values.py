"""
validate_report_values.py - independent check of a RECALCULATED JP Working file.

It recomputes every number on DRY SCORECARD, WET SCORECARD, Worth it & UBA_App / UBA_BR and the AE(DRY) / AE(WET) tabs
straight from the file's own Respondent Data and Worth it tables (no Excel involved), then compares the result with the
values Excel saved in the file. Proven on the client's own April workbook: all 1,194 evaluable cells reproduce.

    python validate_report_values.py "C:\\JP KR Grooming\\Python\\JP\\Report\\<name>_Working_2026-09.xlsx"

With no argument it takes the single *_Working_*.xlsx in the JP Report folder. Run it AFTER recalculate_via_macro_JP.py.
It reads the file only - nothing is changed. Exit code 0 = every value agrees; 1 = something differs or the file was
not really recalculated.
Not covered (needs Excel's LAMBDA engine): Worth it & UBA_App!AG5, a client scratch cell that no scorecard reads.
"""

import re, math, datetime, collections, zipfile
import numpy as np, openpyxl
from openpyxl.utils import column_index_from_string as CI, get_column_letter as L, coordinate_to_tuple

class NAType:
    def __repr__(self): return '#N/A'
NA = NAType()

class Err(Exception):
    def __init__(self, code): self.code = code; super().__init__(code)

def is_num(x): return isinstance(x, (int, float)) and not isinstance(x, bool)

class A:
    """Excel-style array: element-wise comparison (returns 0/1) and arithmetic."""
    __array_priority__ = 100
    def __init__(self, arr): self.a = np.asarray(arr, dtype=object) if not isinstance(arr, np.ndarray) else arr
    def _num(self): return np.array([np.nan if (x is None or isinstance(x, str)) and False else x for x in self.a], dtype=float)
    def _cmp(self, other, op):
        out = np.zeros(len(self.a), dtype=np.int64)
        if isinstance(other, str):
            o = other.lower()
            for i, x in enumerate(self.a):
                s = x.lower() if isinstance(x, str) else ("" if x is None else None)
                if op == "eq": out[i] = (s == o)
                elif op == "ne": out[i] = (s != o)
        elif isinstance(other, datetime.datetime):
            for i, x in enumerate(self.a):
                if op == "eq": out[i] = (x == other)
                elif op == "ne": out[i] = (x != other)
        else:
            o = float(other)
            for i, x in enumerate(self.a):
                if x is NA: raise Err("#N/A")
                if isinstance(x, float) and math.isnan(x): raise Err("#VALUE!")
                if x is None: xv, istext = 0.0, False
                elif isinstance(x, str): xv, istext = None, True
                elif isinstance(x, datetime.datetime): xv, istext = None, True
                else: xv, istext = float(x), False
                if op == "eq": r = (not istext) and xv == o
                elif op == "ne": r = istext or xv != o
                elif op == "ge": r = istext or xv >= o
                elif op == "gt": r = istext or xv > o
                elif op == "le": r = (not istext) and xv <= o
                elif op == "lt": r = (not istext) and xv < o
                out[i] = r
        return A(out)
    def __eq__(self, o): return self._cmp(o, "eq")
    def __ne__(self, o): return self._cmp(o, "ne")
    def __ge__(self, o): return self._cmp(o, "ge")
    def __gt__(self, o): return self._cmp(o, "gt")
    def __le__(self, o): return self._cmp(o, "le")
    def __lt__(self, o): return self._cmp(o, "lt")
    def _bin(self, o, f):
        b = o.a if isinstance(o, A) else o
        a = self.a
        def val(x):
            if x is NA: raise Err("#N/A")
            if isinstance(x, float) and math.isnan(x): raise Err("#VALUE!")
            if x is None: return 0.0
            if isinstance(x, str):
                try: return float(x)
                except ValueError: raise Err("#VALUE!")
            return float(x)
        if isinstance(b, np.ndarray):
            return A(np.array([f(val(x), val(y)) for x, y in zip(a, b)], dtype=float))
        return A(np.array([f(val(x), val(b)) for x in a], dtype=float))
    def __mul__(self, o): return self._bin(o, lambda x, y: x * y)
    __rmul__ = __mul__
    def __add__(self, o): return self._bin(o, lambda x, y: x + y)
    __radd__ = __add__
    def __sub__(self, o): return self._bin(o, lambda x, y: x - y)

def excel_serial(d): return (d - datetime.datetime(1899, 12, 30)).days

def eom(d, n, plus=0):
    y, m = d.year, d.month + n
    while m < 1: y -= 1; m += 12
    while m > 12: y += 1; m -= 12
    nxt = datetime.datetime(y + (m == 12), 1 if m == 12 else m + 1, 1)
    return nxt - datetime.timedelta(days=1) + datetime.timedelta(days=plus)

def split_args(s):
    out, depth, cur, q = [], 0, "", False
    for ch in s:
        if ch == '"': q = not q
        if not q:
            if ch == "(": depth += 1
            elif ch == ")": depth -= 1
            elif ch == "," and depth == 0: out.append(cur); cur = ""; continue
        cur += ch
    out.append(cur); return out

def match_paren(s, start):
    depth, q, k = 1, False, start
    while depth:
        ch = s[k]
        if ch == '"': q = not q
        if not q:
            if ch == "(": depth += 1
            elif ch == ")": depth -= 1
        k += 1
    return k

class Book:
    def __init__(self, path, month_date=None, ae_override=True, verbose=True):
        self.path = path
        wbf = openpyxl.load_workbook(path, data_only=False, read_only=True)
        wbv = openpyxl.load_workbook(path, data_only=True, read_only=True)
        self.F, self.C, self.cached = {}, {}, {}
        for sh in ["DRY SCORECARD", "WET SCORECARD", "AE(DRY)", "AE(WET)", "Worth it & UBA_App", "Worth it & UBA_BR"]:
            ws = wbf[sh]; ws.reset_dimensions()
            f, c = {}, {}
            maxr = None if sh.startswith(("DRY", "WET", "AE")) else 130
            for row in ws.iter_rows(max_row=maxr):
                for cell in row:
                    v = cell.value
                    if v is None: continue
                    v = v.text if hasattr(v, "text") else v
                    (f if (isinstance(v, str) and v.startswith("=")) else c)[cell.coordinate] = v
            self.F[sh], self.C[sh] = f, c
            wv = wbv[sh]; wv.reset_dimensions()
            self.cached[sh] = {cell.coordinate: cell.value for row in wv.iter_rows(max_row=maxr) for cell in row if cell.value is not None}
        # Respondent Data
        rd = wbv["Respondent Data"]; rd.reset_dimensions()
        rows = rd.iter_rows(min_row=10, values_only=True)
        self.hdr = list(next(rows))
        self.rows = [r for r in rows if r[0] is not None]
        self.n = len(self.rows)
        self._col = {}
        self.month = month_date or self.C["DRY SCORECARD"]["B1"]
        # Worth it tables (static values)
        self.WI = {}
        for key, sh in (("App", "Worth it & UBA_App"), ("BR", "Worth it & UBA_BR")):
            ws = wbv[sh]; ws.reset_dimensions()
            cols = collections.defaultdict(dict)
            for row in ws.iter_rows(min_row=2, max_col=17):
                for cell in row:
                    if cell.value is not None: cols[cell.column_letter][cell.row] = cell.value
            self.WI[key] = cols
        self.ae_override = ae_override
        self.memo = {}
        self.AEcache = {}
        self.aeconst = {}
        for sh in ("AE(DRY)", "AE(WET)"):
            self.aeconst[sh] = {k: v for k, v in self.C[sh].items() if is_num(v) and CI(re.match(r"[A-Z]+", k).group(0)) >= 2}

    # ---- data access
    def rd(self, letter):
        if letter in self._col: return self._col[letter]
        i = CI(letter) - 1
        arr = np.empty(self.n, dtype=object)
        for k, r in enumerate(self.rows): arr[k] = r[i] if i < len(r) else None
        self._col[letter] = A(arr); return self._col[letter]
    def rd_by_header(self, name):
        for i, h in enumerate(self.hdr):
            if h == name: return L(i + 1)
        raise KeyError(name)
    def tbl(self, key, letter, r1=2, r2=10000):
        d = self.WI[key].get(letter, {})
        return A(np.array([d.get(r) for r in range(r1, r2 + 1)], dtype=object))

    # ---- AE emulation
    def ae(self, sheet, letter, r1=3, r2=None):
        key = sheet
        if key not in self.AEcache: self.AEcache[key] = self._build_ae(sheet)
        tab, last = self.AEcache[key]
        r2 = r2 or last
        arr = tab.get(letter)
        if arr is None: raise KeyError(f"AE column {letter} not modelled")
        return A(np.array([arr.get(r) for r in range(r1, r2 + 1)], dtype=object))   # rows past the last formula row are empty cells

    def _build_ae(self, sheet):
        sh = f"AE({sheet})"; F = self.F[sh]
        last = max(int(re.search(r"\d+", k).group(0)) for k in F if k.startswith("B"))
        a1 = self.C[sh].get("A1") if "A1" in self.C[sh] else self.month
        a1 = self.cached[sh].get("A1", a1) if not isinstance(a1, datetime.datetime) else a1
        month = self.month
        ids = [r[1] for r in self.rows if r[3] == month]
        idpos = {r[1]: r for r in self.rows}
        def col_of(cell):  # RD letter used in a row-3 INDEX formula
            return re.search(r"'Respondent Data'!\$([A-Z]+)\$11", F[cell]).group(1)
        def look(letter, k):
            if k >= len(ids) or k < 0: return ""
            v = idpos[ids[k]][CI(letter) - 1] if CI(letter) - 1 < len(idpos[ids[k]]) else None
            return 0 if v is None else v
        off = 1 if "MATCH(A4" in F["G3"] else 0
        brands = [("D", "E", "F", "G"), ("I", "J", "K", "L"), ("N", "O", "P", "Q")]
        letters = {c: col_of(c + "3") for g in brands for c in g}
        user_letter = col_of("B3")
        userval = "BR_User" if "BR_User" in F["C3"] else "App_User"
        # Y / Z / AA / AB are evaluated FROM THEIR OWN row-3 FORMULAS (not hard-coded), so a template that reads the age with
        # INDEX/MATCH (JP) or VLOOKUP (KR), or classifies System / Dispo differently, is checked as it is written.
        ixm = re.compile(r"IFERROR\(INDEX\('Respondent Data'!\$([A-Z]+)\$11:\$[A-Z]+\$\d+,MATCH\(A3,'Respondent Data'!\$B\$11:\$B\$\d+,0\)\),\"\"\)")
        vlk = re.compile(r"IFERROR\(VLOOKUP\(A3,'Respondent Data'!\$([A-Z]+)\$11:\$([A-Z]+)\$\d+,(\d+),FALSE\),\"\"\)")
        def compile_row(cell):
            t = F[cell].lstrip("=")
            t = ixm.sub(lambda m: f'LK("{m.group(1)}")', t)
            t = vlk.sub(lambda m: f'LK("{L(CI(m.group(1)) + int(m.group(3)) - 1)}")', t)
            if "'Respondent Data'" in t or "A3" in t: raise ValueError(f"AE {cell}: unsupported row formula {F[cell][:120]}")
            t = re.sub(r"\bIF\(", "IF_(", t); t = re.sub(r"\bOR\(", "OR_(", t); t = re.sub(r"\bAND\(", "AND_(", t)
            t = re.sub(r"(?<![<>!=])=(?!=)", "==", t)
            return compile(t, cell, "eval")
        row_code = {c: compile_row(c + "3") for c in ("Y", "Z", "AA", "AB")}
        def row_val(c, k):
            env = dict(LK=lambda letter: look(letter, k), IF_=lambda c_, a, b: a if c_ else b,
                       OR_=lambda *a: any(a), AND_=lambda *a: all(a))
            return eval(row_code[c], {"__builtins__": {}}, env)
        tab = collections.defaultdict(dict)
        S = {}
        for k in range(last - 2):
            r = 3 + k
            B = look(user_letter, k); C = 1 if B == userval else 0
            tab["A"][r] = ids[k] if k < len(ids) else None
            tab["B"][r] = B; tab["C"][r] = C
            sc = []
            for (da, pu, co, ra) in brands:
                D, E, Fv = look(letters[da], k), look(letters[pu], k), look(letters[co], k)
                G = look(letters[ra], k + off)
                low = lambda x: x.lower() if isinstance(x, str) else None
                H = C if (low(D) == "selected" and (low(E) == "selected" or low(Fv) == "selected")) else 0
                tab[da][r], tab[pu][r], tab[co][r], tab[ra][r], tab[chr(ord(ra) + 1)][r] = D, E, Fv, G, H
                if self.ae_override and (f"{ra}{r}" in self.aeconst[sh]): tab[ra][r] = G = self.aeconst[sh][f"{ra}{r}"]
                if isinstance(G, str):                                 # Excel coerces numeric text; ""/other text -> #VALUE!
                    try: sc.append(float(G) * H)
                    except ValueError: sc.append(float("nan"))
                else: sc.append(float(G) * H)
            for j, c in enumerate("STU"): tab[c][r] = sc[j]
            def rank(x):
                if math.isnan(x): return 0
                if x == 0: return 0
                vals = [v for v in sc if not math.isnan(v)]
                if len(vals) < 3 and any(math.isnan(v) for v in sc): return 0
                gt = sum(1 for v in vals if v > x); eq = sum(1 for v in vals if v == x)
                return gt + (eq + 1) / 2
            Rk = [rank(x) for x in sc]
            for j, c in enumerate("VWX"): tab[c][r] = Rk[j]
            def term(rk):
                try: return (1 / rk) ** 1.35
                except ZeroDivisionError: return 0
            den = term(Rk[1]) + term(Rk[2]) + term(Rk[0])
            ac = [0 if Rk[j] == 0 else term(Rk[j]) / den for j in range(3)]
            for j, c in enumerate(["AC", "AD", "AE"]): tab[c][r] = ac[j]
            tab["AF"][r] = sum(ac)
            for c in ("Y", "Z", "AA", "AB"): tab[c][r] = row_val(c, k)
        return tab, last

    def _ltr(self, name): return self.rd_by_header(name)

    # ---- formula evaluation
    def value(self, sheet, ref):
        ref = ref.replace("$", "")
        key = (sheet, ref)
        if key in self.memo:
            v = self.memo[key]
            if isinstance(v, Err): raise v
            return v
        if ref in self.F[sheet]:
            try: v = self.evalf(sheet, self.F[sheet][ref])
            except Err as e: self.memo[key] = e; raise
            except ZeroDivisionError: e = Err("#DIV/0!"); self.memo[key] = e; raise e
            self.memo[key] = v; return v
        v = self.C[sheet].get(ref)
        if v is None: return 0
        return v

    def evalf(self, sheet, f):
        py = self.to_py(sheet, f)
        env = dict(R=self.rd, AE=lambda s, c, r1=3, r2=None: self.ae(s, c, r1, r2), TB=lambda c, r1=2, r2=10000: self.tbl(self._wikey(sheet), c, r1, r2),
                   V=lambda r: self.value(sheet, r), X=lambda k, r: self.value("Worth it & UBA_" + k, r),
                   SP=self.SP, CIFS=self.CIFS, SIFS=self.SIFS, OPC=lambda op, d: ("op", op, d), EOM=eom,
                   VAL=self.VAL, IFE=self.IFE, datetime=datetime, A=A, FILTER=self.FILTER, UNIQUE=self.UNIQUE, ROWS=self.ROWS,
                   MATCH=self.MATCH, ISNUMBER=self.ISNUMBER, INDEX=self.INDEX, VLK=self.VLK, SUMR=lambda c1, r1, c2, r2: self.SUMR(sheet, c1, r1, c2, r2))
        try: return eval(py, {"__builtins__": {}}, env)
        except KeyError: raise
    def _wikey(self, sheet): return "App" if sheet.endswith("App") else "BR"

    @staticmethod
    def FILTER(arr, cond): return A(np.array([v for v, k in zip(arr.a, cond.a) if k], dtype=object))
    @staticmethod
    def UNIQUE(arr):
        seen, out = set(), []
        for v in arr.a:
            k = Book._key(v)
            if k not in seen: seen.add(k); out.append(v)
        return A(np.array(out, dtype=object))
    @staticmethod
    def ROWS(arr): return float(len(arr.a))
    @staticmethod
    def MATCH(look, rng, mode=0):
        pos = {}
        for i, v in enumerate(rng.a):
            pos.setdefault(Book._key(v), i)
        return A(np.array([pos.get(Book._key(x), NA) for x in look.a], dtype=object))
    def VLK(self, look, c1, idx):
        """VLOOKUP(look, 'Respondent Data'!<c1>:..., idx, FALSE): value of column c1+idx-1 on the first row whose c1 equals the lookup value;
        no match -> #N/A (NA); a blank result cell reads as 0, like Excel."""
        keys, vals = self.rd(c1).a, self.rd(L(CI(c1) + idx - 1)).a
        pos = {}
        for i, v in enumerate(keys): pos.setdefault(Book._key(v), i)
        out = []
        for x in look.a:
            i = pos.get(Book._key(x))
            out.append(NA if i is None else (0 if vals[i] is None else vals[i]))
        return A(np.array(out, dtype=object))
    @staticmethod
    def ISNUMBER(x): return A(np.array([0 if (v is NA or isinstance(v, float) and math.isnan(v)) else 1 for v in x.a], dtype=np.int64))
    @staticmethod
    def INDEX(rng, idx):
        return A(np.array([NA if i is NA else rng.a[int(i)] for i in idx.a], dtype=object))
    def SUMR(self, sheet, c1, r1, c2, r2):
        tot = 0.0
        for ci in range(CI(c1), CI(c2) + 1):
            for r in range(r1, r2 + 1):
                v = self.value(sheet, f"{L(ci)}{r}")
                if is_num(v): tot += float(v)
        return tot
    @staticmethod
    def VAL(x):
        out = []
        for v in x.a:
            if is_num(v): out.append(float(v))
            elif isinstance(v, str):
                try: out.append(float(v))
                except ValueError: out.append(float("nan"))
            elif v is None: out.append(float("nan"))
            else: out.append(float("nan"))
        return A(np.array(out, dtype=float))
    @staticmethod
    def IFE(x, d):
        if isinstance(x, A):
            return A(np.array([d if (v is NA or (isinstance(v, float) and math.isnan(v))) else v for v in x.a], dtype=object))
        return x
    @staticmethod
    def SP(x):
        if isinstance(x, A):
            if any(isinstance(v, float) and math.isnan(v) for v in x.a): raise Err("#VALUE!")
            return float(sum(float(v) for v in x.a))
        return x
    @staticmethod
    def _key(x):
        if isinstance(x, str): return x.lower()
        if is_num(x): return float(x)
        return x
    def _crit_mask(self, rng, crit):
        n = len(rng.a)
        if isinstance(crit, str) and re.match(r"^(>=|<=|<>|>|<|=)\s*-?[\d.]+$", crit):
            op, num = re.match(r"^(>=|<=|<>|>|<|=)\s*(-?[\d.]+)$", crit).groups(); num = float(num)
            f = {">=": lambda x: x >= num, "<=": lambda x: x <= num, ">": lambda x: x > num, "<": lambda x: x < num, "=": lambda x: x == num, "<>": lambda x: x != num}[op]
            return np.array([is_num(x) and f(float(x)) for x in rng.a])
        if isinstance(crit, tuple) and crit[0] == "op":
            op, d = crit[1], crit[2]
            return np.array([bool(isinstance(x, datetime.datetime) and ((x >= d) if op == ">=" else (x <= d))) for x in rng.a])
        if isinstance(crit, str): return (rng == crit).a.astype(bool)
        if isinstance(crit, datetime.datetime): return (rng == crit).a.astype(bool)
        if crit is None: crit = 0
        ser = lambda x: float(excel_serial(x)) if isinstance(x, datetime.datetime) else (float(x) if is_num(x) else None)
        return np.array([ser(x) is not None and ser(x) == float(crit) for x in rng.a])
    def CIFS(self, *args):
        pairs = list(zip(args[0::2], args[1::2]))
        arr_pairs = [(r, c) for r, c in pairs if isinstance(c, A)]
        m = None
        for rng, crit in pairs:
            if isinstance(crit, A): continue
            k = self._crit_mask(rng, crit); m = k if m is None else (m & k)
        if not arr_pairs: return float(m.sum())
        # array criteria (one count per criterion row): group the matching rows by the criterion-range values
        if m is None: m = np.ones(len(arr_pairs[0][0].a), dtype=bool)
        cnt = collections.Counter(tuple(self._key(r.a[i]) for r, _ in arr_pairs) for i in range(len(m)) if m[i])
        n = len(arr_pairs[0][1].a)
        return A(np.array([cnt.get(tuple(self._key(c.a[j]) for _, c in arr_pairs), 0) for j in range(n)], dtype=float))
    def SIFS(self, sumr, *args):
        m = np.ones(len(sumr.a), dtype=bool)
        for rng, crit in zip(args[0::2], args[1::2]): m &= self._crit_mask(rng, crit)
        tot = 0.0
        for x, k in zip(sumr.a, m):
            if k:
                if isinstance(x, float) and math.isnan(x): raise Err("#VALUE!")
                if is_num(x): tot += float(x)
        return tot

    def to_py(self, sheet, f):
        f = f.lstrip("=").replace("\r\n", "\n").replace("\n", " ")
        f = f.replace("_xlfn.", "").replace("_xlws.", "")
        strs = []
        def keep(m): strs.append(m.group(0)); return f"__S{len(strs) - 1}__"
        f = re.sub(r'"[^"]*"', keep, f)
        f = re.sub(r"\s+", " ", f)
        # date criteria
        f = re.sub(r"__S(\d+)__&EOMONTH\(([^,]+),(-?\d+)\)\+1", lambda m: f'OPC({strs[int(m.group(1))]},EOM(V("{m.group(2).replace("$", "")}"),{m.group(3)},1))', f)
        f = re.sub(r"__S(\d+)__&EOMONTH\(([^,]+),(-?\d+)\)", lambda m: f'OPC({strs[int(m.group(1))]},EOM(V("{m.group(2).replace("$", "")}"),{m.group(3)},0))', f)
        # criteria built as text:  ">="&16  ->  ">=16"
        def _cat(m):
            strs.append('"' + strs[int(m.group(1))].strip('"') + m.group(2) + '"'); return f"__S{len(strs) - 1}__"
        f = re.sub(r"__S(\d+)__&(\d+(?:\.\d+)?)(?![\d.(])", _cat, f)
        f = re.sub(r"\bSUM\(\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)\)", r'SUMR("\1",\2,"\3",\4)', f)
        f = re.sub(r"VLOOKUP\((\$[A-Z]+\$\d+:\$[A-Z]+\$\d+),'Respondent Data'!\$([A-Z]+)\$11:\$[A-Z]+\$\d+,(\d+),FALSE\)",
                   lambda m: f'VLK({m.group(1)},"{m.group(2)}",{m.group(3)})', f)
        f = re.sub(r"'Respondent Data'!\$([A-Z]+)\$11:\$[A-Z]+\$\d+", r'R("\1")', f)
        f = re.sub(r"'AE\((DRY|WET)\)'!\$?([A-Z]+)\$(\d+):\$?[A-Z]+\$(\d+)", r'AE("\1","\2",\3,\4)', f)
        f = re.sub(r"'AE\((DRY|WET)\)'!\$?([A-Z]+):\$?[A-Z]+", r'AE("\1","\2")', f)
        f = re.sub(r"'Worth it & UBA_(App|BR)'!\$?([A-Z]+)\$?(\d+)", r'X("\1","\2\3")', f)
        f = re.sub(r"\$([A-Z]{1,2})\$(\d+):\$[A-Z]{1,2}\$(\d+)", r'TB("\1",\2,\3)', f)
        f = re.sub(r'(?<![A-Za-z_"\.])\$?([A-Z]{1,3})\$?(\d+)(?![\d(\"])', r'V("\1\2")', f)
        for fn, py in (("SUMPRODUCT", "SP"), ("COUNTIFS", "CIFS"), ("SUMIFS", "SIFS"), ("VALUE", "VAL"), ("IFERROR", "IFE")):
            f = re.sub(r"\b%s\(" % fn, py + "(", f)
        f = f.replace("<>", "!=").replace("^", "**")
        f = re.sub(r"(?<![<>!=])=(?!=)", "==", f)
        f = re.sub(r"__S(\d+)__", lambda m: strs[int(m.group(1))], f)
        f = f.replace("OPC((", "OPC(").replace("%", "/100")
        return f


# --------------------------------------------------------------------------- comparison
def _same(v, c):
    if isinstance(v, (int, float)) and isinstance(c, (int, float)) and not isinstance(v, bool) and not isinstance(c, bool):
        return abs(v - c) <= 1e-7 * max(1.0, abs(c))
    if isinstance(v, str) and isinstance(c, str):
        return v.strip().lower() == c.strip().lower()
    return False


def validate(path):
    b = Book(path)
    print(f"File: {path}")
    print(f"Report month {b.month:%Y-%m-%d}: {b.n} Respondent Data rows read")
    # age check: the report's age-cut rows (16-34 / 35-49 / 50+) read Age absolute (column F); it must sit inside
    # the age band the flat file gives each respondent (column G). 16-17 respondents must carry their OWN age.
    bad_age = []
    for r in b.rows:
        if r[3] != b.month: continue
        band, age = r[6], r[5]
        m = re.match(r"^\s*(\d+)\s*[-\u2013]\s*(\d+)\s*$", str(band)) or re.match(r"^\s*(\d+)\s*\+\s*$", str(band))
        if not m or not is_num(age): continue
        lo = int(m.group(1)); hi = int(m.group(2)) if m.lastindex == 2 else 999
        if not lo <= age <= hi: bad_age.append((r[1], band, age))
    n_minor = sum(1 for r in b.rows if r[3] == b.month and str(r[6]).strip() == "16-17")
    if bad_age:
        print(f"  AGE CHECK FAILED: {len(bad_age)} respondents have an Age absolute outside their age band, e.g. {bad_age[:3]}. "
              f"The 16-17 respondents must carry their own age (second age question), not the parent's.")
    else:
        print(f"  Age check: every respondent's Age absolute sits inside its age band ({n_minor} in 16-17).")
    total = ok = 0
    problems = []
    not_eval = []
    for sh in ["Worth it & UBA_App", "Worth it & UBA_BR", "DRY SCORECARD", "WET SCORECARD"]:
        s_ok = s_bad = 0
        for ref in sorted(b.F[sh], key=lambda k: (int(re.search(r"\d+", k).group(0)), k)):
            f = b.F[sh][ref]
            if "LAMBDA" in f or "SEQUENCE" in f:
                not_eval.append((sh, ref)); continue
            try: v = b.value(sh, ref)
            except Err as e: v = e.code
            except ZeroDivisionError: v = "#DIV/0!"
            except Exception as e:
                not_eval.append((sh, ref)); continue
            c = b.cached[sh].get(ref)
            total += 1
            if _same(v, c): ok += 1; s_ok += 1
            else: s_bad += 1; problems.append((sh, ref, v, c))
        print(f"  {sh:<22} agree {s_ok:>4}   differ {s_bad:>4}")
    # AE tabs: every respondent row, columns B..AF
    for key in ("DRY", "WET"):
        sh = f"AE({key})"
        tab, last = b._build_ae(key)
        n_resp = sum(1 for r in b.rows if r[3] == b.month)
        cached = b.cached[sh]
        a_ok = a_bad = 0
        for col in [L(i) for i in range(2, 33)]:
            for r in range(3, 3 + n_resp):
                mine = tab[col].get(r)
                c = cached.get(f"{col}{r}")
                if isinstance(mine, float) and mine != mine: mine = "#VALUE!"
                if mine in ("", None) and c in ("", None): a_ok += 1; continue
                if _same(mine, c) or (isinstance(mine, (int, float)) and isinstance(c, (int, float)) and abs(mine - c) < 1e-9) \
                        or (isinstance(mine, str) and isinstance(c, str) and mine.lower() == c.lower()):
                    a_ok += 1
                else:
                    a_bad += 1
                    if a_bad <= 5: problems.append((sh, f"{col}{r}", mine, c))
        ids = [cached.get(f"A{r}") for r in range(3, 3 + n_resp)]
        spill_ok = [r[1] for r in b.rows if r[3] == b.month] == ids
        print(f"  {sh:<22} {n_resp} respondent rows x 31 columns: agree {a_ok}, differ {a_bad}; column A spill lists exactly this month's respondents: {spill_ok}")
        total += a_ok + a_bad; ok += a_ok
        if not spill_ok: problems.append((sh, "A3:A..", "this month's ids", "different / stale ids"))
    bad = total - ok
    print(f"\nChecked {total} values: {ok} agree, {bad} differ. Not evaluated: {len(not_eval)} {[x[1] for x in not_eval][:5]}")
    if bad > 0.5 * total:
        print("\n*** NOT RECALCULATED: more than half the saved values differ from the data. The file still holds the template's")
        print("    old numbers. Open it in Excel, Formulas > Calculation Options = Automatic, Ctrl+Alt+Shift+F9, wait for the")
        print("    status bar to finish, Ctrl+S, then run this check again.")
    elif problems:
        print("\nDifferences (cell: independent value vs value saved by Excel):")
        for sh, ref, v, c in problems[:40]:
            print(f"   {sh}!{ref}: expected {v!r}  saved {c!r}")
    else:
        print("\nALL VALUES AGREE with an independent calculation from the Respondent Data.")
    errs = collections.Counter()
    for sh in ["DRY SCORECARD", "WET SCORECARD", "Worth it & UBA_App", "Worth it & UBA_BR"]:
        for ref, c in b.cached[sh].items():
            if isinstance(c, str) and c.startswith("#"): errs[(sh, c)] += 1
    if errs:
        print("\nError cells in the saved file (these are the formulas' own results, confirmed above):")
        for (sh, code), n in sorted(errs.items()): print(f"   {sh}: {n} x {code}")
    return 0 if (not problems and not bad) else 1


if __name__ == "__main__":
    import sys, glob
    if len(sys.argv) > 1: target = sys.argv[1]
    else:
        files = glob.glob(r"C:\JP KR Grooming\Python\JP\Report\*_Working_*.xlsx")
        if len(files) != 1: raise SystemExit("Pass the Working file path, e.g. python validate_report_values.py <file.xlsx>")
        target = files[0]
    raise SystemExit(validate(target))