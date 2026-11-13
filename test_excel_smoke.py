"""
Smoke test - run this BEFORE recalculate_excel.py on the real workbook.

Two separate things can go wrong with the Excel recalculation step, and this
isolates them from each other and from anything specific to the 11MB
workbook itself:

  1. xlwings/COM automation can't launch Excel at all (missing install,
     COM/security restrictions, licensing, etc.)
  2. Excel launches fine, but the installed VERSION doesn't support
     UNIQUE/FILTER - these dynamic-array functions only exist in Excel 365
     or Excel 2021+ (LTSC). Excel 2019 and older would fail on the exact
     same Worth it & UBA_App/UBA_BR formulas that LibreOffice failed on,
     just with #NAME? instead of #VALUE!.

This creates a tiny 6-row test file with the SAME formula shape used in the
real workbook (ROWS(UNIQUE(FILTER(...)))), recalculates it through Excel,
and checks the answer is right (expected: 4). Takes seconds, not minutes -
if this fails, there's no point running the full recalculation yet.

Run this first:
    pip install xlwings          (if not already installed)
    python test_excel_smoke.py
"""
import sys
import time


def build_test_file(path="test_unique_filter.xlsx"):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Test"
    ws["A1"] = "Month"; ws["B1"] = "Resp"; ws["C1"] = "Brand"
    data = [
        ("June", 1, "Braun"), ("June", 1, "Philips"), ("June", 2, "Braun"),
        ("June", 3, "Panasonic"), ("June", 4, "Braun"), ("June", 5, "Philips"),
    ]
    for i, row in enumerate(data, start=2):
        ws.cell(row=i, column=1, value=row[0])
        ws.cell(row=i, column=2, value=row[1])
        ws.cell(row=i, column=3, value=row[2])
    ws["E1"] = "June"
    # Same shape as the real formulas in Worth it & UBA_App/UBA_BR
    ws["F1"] = ('=ROWS(_xlfn.UNIQUE(_xlfn._xlws.FILTER($B$2:$B$7,'
                 '($A$2:$A$7=$E$1)*(($C$2:$C$7="Braun")+($C$2:$C$7="Philips")))))')
    wb.save(path)
    return path


def run_smoke_test():
    print("Step 1: checking xlwings is installed...")
    try:
        import xlwings as xw
    except ImportError:
        print("  FAILED: xlwings is not installed. Run: pip install xlwings")
        return False
    print(f"  OK: xlwings {xw.__version__}")

    print("\nStep 2: building test file with the real workbook's UNIQUE/FILTER shape...")
    path = build_test_file()
    print(f"  OK: {path} created")

    print("\nStep 3: launching Excel via COM automation...")
    t0 = time.time()
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
    except Exception as e:
        print(f"  FAILED: could not launch Excel. {type(e).__name__}: {e}")
        print("  This usually means either Excel isn't installed/licensed on this "
              "machine, or COM automation is blocked (common in locked-down "
              "enterprise environments, or when running as a service without "
              "'allow interacting with desktop').")
        return False
    print(f"  OK: Excel launched in {time.time() - t0:.1f}s")

    try:
        print(f"\nStep 4: checking Excel version...")
        try:
            version = app.version
            print(f"  Excel version: {version}")
            major = int(str(version).split(".")[0])
            if major < 16:
                print(f"  WARNING: version {version} looks old - UNIQUE/FILTER need "
                      f"Excel 365 or Excel 2021+ (LTSC). This test will tell us "
                      f"definitively either way.")
        except Exception as e:
            print(f"  Could not read version ({e}), continuing anyway - the formula "
                  f"test below is the real answer.")

        print("\nStep 5: opening test file, forcing full recalculation...")
        t0 = time.time()
        wb = app.books.open(path)
        app.api.CalculateFullRebuild()   # Application-level method, not Workbook
        elapsed = time.time() - t0
        print(f"  OK: recalculated in {elapsed:.1f}s")

        print("\nStep 6: reading back the result...")
        result = wb.sheets["Test"].range("F1").value
        wb.close()

        print(f"  F1 = {result!r} (expected: 4)")
        if result == 4:
            print("\nPASS: this machine's Excel correctly evaluates UNIQUE/FILTER. "
                  "Safe to run recalculate_excel.py on the real workbook.")
            return True
        else:
            print(f"\nFAIL: got {result!r} instead of 4. This Excel install may not "
                  f"fully support dynamic array formulas, or something else is off. "
                  f"Do not proceed to the full workbook until this is resolved - "
                  f"share this output and I'll help diagnose.")
            return False
    except Exception as e:
        import traceback
        print(f"\n  FAILED during recalculation: {type(e).__name__}: {e}")
        traceback.print_exc()
        return False
    finally:
        app.quit()


if __name__ == "__main__":
    success = run_smoke_test()
    sys.exit(0 if success else 1)
