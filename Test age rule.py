"""Tests the age rule in build_idxage.extract_id_age_pairs on small synthetic raw files:
the second (16-17) age question overrides the first only when it holds a real number.
Run from the folder that holds build_idxage.py:   python test_age_rule.py"""
import os, sys, tempfile, openpyxl
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_idxage import extract_id_age_pairs

ID, AGE, MINOR = "Internal Respondent Id", "age q", "child age q"

def make(rows, headers, path):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.append(headers)
    for r in rows: ws.append(r)
    wb.save(path)

def run(rows, headers=(ID, AGE, MINOR), minor=MINOR):
    p = os.path.join(tempfile.mkdtemp(), "raw.xlsx"); make(rows, list(headers), p)
    return dict(extract_id_age_pairs(p, age_header=AGE, minor_age_header=minor))

ok = True
def check(name, got, want):
    global ok
    good = got == want; ok &= good
    print(("PASS " if good else "FAIL ") + name + ("" if good else f"  got {got} want {want}"))

check("adult only (child column blank)", run([[1, 45, None]]), {1: 45})
check("child number overrides parent age", run([[2, 47, 17]]), {2: 17})
check("child age stored as text '16' overrides", run([[3, "48", "16"]]), {3: 16})
check("child column holds survey text -> parent column used", run([[4, 57, "설문조사"], [5, 45, "참여하기"]]), {4: 57, 5: 45})
check("child column holds a float 17.0", run([[6, 50, 17.0]]), {6: 17})
check("child number outside 10-17 is ignored", run([[7, 50, 40]]), {7: 50})
check("mixed month", run([[1, 30, None], [2, 44, 16], [3, 60, "x"]]), {1: 30, 2: 16, 3: 60})
check("child question missing from file -> parent ages (with a warning)", run([[1, 45], [2, 47]], headers=(ID, AGE), minor=MINOR), {1: 45, 2: 47})
check("both questions share one header -> 1st is adult, 2nd is child", run([[1, 45, 17]], headers=(ID, AGE, AGE), minor=AGE), {1: 17})
print("\nALL PASSED" if ok else "\nSOME FAILED"); sys.exit(0 if ok else 1)