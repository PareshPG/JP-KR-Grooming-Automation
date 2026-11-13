"""
Alternative recalculation path - launches AutoRecalcRunner.xlsm as an
ordinary new process (os.startfile, the same mechanism as double-clicking
the file) rather than instantiating Excel as a COM server. See
AutoRecalcRunner_VBA_source.txt for the reasoning and one-time setup.

UNTESTED on real hardware - this could not be verified in the environment
that built it (no Windows/Excel available). It represents a genuinely
different automation strategy from the xlwings/COM approach that was
confirmed blocked, worth trying specifically because it doesn't use
inter-process COM control at all - but it needs its first real test on
your machine, the same as every other Windows-specific piece of this
project. If it doesn't work, the manual Ctrl+Alt+Shift+F9 + Ctrl+S
fallback is completely unaffected and still works exactly as before.
"""
import os
import time
from pathlib import Path
from recalculate_excel import scan_for_errors, check_sample_sizes
from validate_ae_ratings import find_ae_rating_issues
from market_config import JP

RUNNER_PATH = r"C:\JP KR Grooming\Python\AutoRecalcRunner.xlsm"
LOG_PATH = r"C:\JP KR Grooming\Python\recalc_result.txt"
TIMEOUT_SECONDS = 3600   # a REAL full recalculation of this workbook takes minutes, not seconds (the Worth it / scorecard
                         # SUMPRODUCTs read ~1M-row ranges); 120s was only enough while the file was not being recalculated at all
POLL_INTERVAL_SECONDS = 2


def check_recalculated(working_path):
    """Prove the file was REALLY recalculated, not just opened and saved. A macro SUCCESS only means Excel opened
    and saved the file - on 27 Oct 2026 that happened while every cached value was still the template's old
    (April) figure. Two independent tests on the saved values:
      1. DRY SCORECARD 'Sample Size' (first one) == the number of App_User respondents of the report month,
         counted straight from Respondent Data;
      2. AE(DRY) A3 (first respondent of the FILTER spill) is a respondent of the report month.
    Raises RuntimeError if either fails."""
    import openpyxl
    wb = openpyxl.load_workbook(working_path, read_only=True, data_only=True)
    sc = wb["DRY SCORECARD"]; sc.reset_dimensions()
    month = sc["B1"].value
    sample = None
    for row in sc.iter_rows(min_row=1, max_row=40, min_col=2, max_col=3, values_only=True):
        if isinstance(row[0], str) and row[0].strip().lower() == "sample size":
            sample = row[1]
            break
    rd = wb["Respondent Data"]; rd.reset_dimensions()
    rows = rd.iter_rows(min_row=10, values_only=True)
    header = [str(h).strip() if h is not None else "" for h in next(rows)]
    i_id, i_mm, i_app = header.index("Resp id"), header.index("MM/YY"), header.index("App_User")
    month_ids, n_app = set(), 0
    for r in rows:
        if r[i_mm] == month and r[i_id] is not None:
            month_ids.add(r[i_id])
            n_app += 1 if r[i_app] == "App_User" else 0
    ae = wb["AE(DRY)"]; ae.reset_dimensions()
    first_id = ae["A3"].value
    problems = []
    if sample != n_app:
        problems.append(f"DRY SCORECARD sample size is {sample!r} but Respondent Data has {n_app} App_User respondents in {month}")
    if first_id not in month_ids:
        problems.append(f"AE(DRY)!A3 is {first_id!r}, which is not a respondent of {month}")
    if problems:
        raise RuntimeError(
            "NOT RECALCULATED - the file was saved, but its numbers are still the old ones:\n  - " + "\n  - ".join(problems)
            + "\nOpen the _Working file in Excel, check Formulas > Calculation Options = Automatic, press Ctrl+Alt+Shift+F9, "
              "wait until the status bar stops saying 'Calculate', then Ctrl+S. Do not deliver the file before this passes.")
    print(f"Recalculation check OK: DRY sample size {sample} = App_User respondents in {month:%Y-%m}; AE(DRY) lists this month's respondents.")


def recalculate_via_macro(runner_path=RUNNER_PATH, log_path=LOG_PATH,
                            timeout=TIMEOUT_SECONDS, wait_only=False):
    """Launch the macro runner as a plain process, wait for it to finish
    (detected via the log file it writes on completion), and report the
    result. Raises if the runner file is missing, if it times out, or if
    the macro itself reported failure."""
    runner = Path(runner_path)
    if not runner.exists():
        raise FileNotFoundError(
            f"Runner file not found: {runner}\n"
            f"Complete the one-time setup in AutoRecalcRunner_VBA_source.txt first."
        )

    log = Path(log_path)
    if wait_only:
        # Excel was already started by an earlier run and may still be calculating - do NOT launch it again
        # (a second Excel would open the same file). Just keep watching for the result file.
        print(f"Wait-only mode: not launching Excel; watching for {log.name} ...")
    else:
        if log.exists():
            log.unlink()  # clear any stale log from a previous run
        print(f"Launching {runner.name} (plain process, not COM automation)...")
        os.startfile(str(runner))  # Windows-only: opens the file with its default app, same as double-clicking

    print(f"Waiting up to {timeout // 60} min for it to finish. A full recalculation of this report takes several minutes -\n"
          f"  you can watch Excel's status bar ('Calculating ...'). Do not close Excel and do not start this script a second time.")
    waited = 0
    while not log.exists():
        time.sleep(POLL_INTERVAL_SECONDS)
        waited += POLL_INTERVAL_SECONDS
        if waited % 60 == 0:
            print(f"  ... still waiting ({waited // 60} min)")
        if waited >= timeout:
            raise TimeoutError(
                f"No result after {timeout // 60} minutes. Excel may still be calculating (look at EXCEL.EXE in Task Manager "
                f"and the status bar of the open Excel window), or it may be stuck on a dialog / macro-security prompt "
                f"(Trusted Location setup). If Excel is still working, leave it: when it finishes it saves the file, writes "
                f"{log.name} and closes - then run:  python recalculate_via_macro_JP.py --wait-only   "
                f"(or just  python validate_report_values.py)."
            )

    result = log.read_text().strip()
    print(f"Result: {result}")
    log.unlink()  # cleanup

    if result.startswith("FAIL"):
        raise RuntimeError(f"Macro reported failure: {result}")

    # Restore the validation the old xlwings-based path used to run
    # automatically - this was silently lost when switching to the
    # macro-based approach, since success here only proves the file
    # recalculated, not that the result is free of formula errors or
    # built on suspiciously small samples.
    working_file = result.split("SUCCESS: ", 1)[1].split(" recalculated")[0]
    working_path = Path(JP.report_folder) / working_file
    if working_path.exists():
        print()
        check_recalculated(str(working_path))   # stops here if the values are stale
        scan_for_errors(str(working_path))
        check_sample_sizes(str(working_path))
        find_ae_rating_issues(str(working_path))
    else:
        print(f"\nWARNING: could not locate {working_path} for post-recalc "
              f"validation (error scan / sample size check skipped) - the "
              f"macro reported success, but verify the file manually.")

    return result


if __name__ == "__main__":
    import sys
    recalculate_via_macro(wait_only="--wait-only" in sys.argv)