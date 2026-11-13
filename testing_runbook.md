# JP-KR Grooming Report Pipeline — Testing & Monthly Runbook

Three parts: (1) one-time setup checks for a new machine, (2) validating this specific
replacement of `run_monthly_pipeline.py`, (3) the repeatable steps for every future month.

---

## Part 1 — One-time machine setup (skip if already confirmed on this machine)

1. **Confirm Python is installed and on PATH**
   ```
   python --version
   ```
   Should print a version number. If not, install from python.org, checking "Add
   python.exe to PATH" during setup.

2. **Install dependencies**
   ```
   python -m pip install openpyxl xlwings lxml
   ```
   Confirm each imports cleanly:
   ```
   python -c "import openpyxl, xlwings, lxml; print('all OK')"
   ```

3. **Confirm all files are present** in the same folder as the client's files
   (`JP-June.xlsx`, `June-Rawdata.xlsx`, or whatever the current month's names are):
   - `safe_xlsx_writer.py`
   - `build_idxage.py`
   - `clean_ratings.py`
   - `build_jq_tables.py`
   - `set_report_month.py`
   - `recalculate_excel.py`
   - `run_monthly_pipeline.py`
   - `equity_color_coding.py`
   - `test_excel_smoke.py`
   - `debug_recalc_visible.py` *(debug tool only — not needed for normal runs)*

4. **Run the Excel smoke test**
   ```
   python test_excel_smoke.py
   ```
   Expect: `PASS: this machine's Excel correctly evaluates UNIQUE/FILTER.`
   The critical line is `F1 = 4.0 (expected: 4)`. If this fails, stop — nothing past
   this point will work until it passes.

---

## Part 2 — Validating this replacement of `run_monthly_pipeline.py`

You've just swapped in the latest version. Confirm it behaves identically to the last
fully-validated run before trusting it.

1. **Delete the old working file** so you get a genuinely fresh copy:
   ```
   del JP-June_Working.xlsx
   ```

2. **Run the prep steps only:**
   ```
   python run_monthly_pipeline.py --skip-recalc
   ```
   Enter `2026-06` at the month prompt.

3. **Check the output against these known-correct numbers** (this exact June data has
   been run and independently verified multiple times — any mismatch means something
   changed and is worth investigating before going further):

   | Check | Expected |
   |---|---|
   | Raw age pairs extracted | 300 |
   | IDxAGE rows written | 300 |
   | Rating cells fixed (total) | 12,178 |
   | — Braun / Philips / Panasonic | 2,361 / 1,943 / 2,318 |
   | — Gillette / Schick / Kai | 2,045 / 2,072 / 1,439 |
   | Messy patterns found | `1- Is terrible`, `10 - It is perfect in every way` |
   | Worth it & UBA_App table one → two | 5,559 → 4,184 rows, `[VALIDATED]` |
   | Worth it & UBA_BR table one → two | 5,128 → 3,914 rows, `[VALIDATED]` |
   | Integrity check | `0 unexpected parts lost, 10 sheet(s) changed` |
   | Original file hash | `88dd022338b0...` (unchanged) |
   | Respondent cross-check | `300 / 300 match exactly` |

4. **Open `JP-June_Working.xlsx` directly in Excel.** This is the real test — the
   script succeeding is necessary but not sufficient.
   - **No repair dialog appears** ("We found a problem with some content...").
   - If a yellow "External Data Connections have been disabled" bar appears, that's
     normal and expected — click "Enable Content" if you want live Power Query
     behavior, though it isn't required for viewing the report.

5. **Recalculate manually** (confirmed working; automated recalculation is blocked
   on this machine by security policy, not by anything wrong with the file — see
   Part 3, step 5 for why):
   - Press `Ctrl+Alt+Shift+F9` (full recalculation)
   - Press `Ctrl+S` (save)
   - Both should complete with **no dialogs at all**.

6. **Sanity-check the report itself.** Open `DRY SCORECARD` and `WET SCORECARD` and
   confirm the numbers look like June's report — sample sizes should read:

   | Segment | N |
   |---|---|
   | Dry REP | 213 |
   | Wet REP | 207 |
   | Dispo *(expect this one flagged — see note)* | 6 |

   The `Dispo` segment has a genuinely tiny base (n=6) and will show `#DIV/0!` in two
   cells (`Consideration to Purchase`, `Worth-it Normalized`) — this is expected,
   not a bug, and was present in the client's own original file before any of this
   automation existed. Treat that segment as directional only in any report built
   from this file.

7. **Apply equity color coding** — must run AFTER recalculation, not before, since it
   reads live calculated scores that don't exist until step 5 is done:
   ```
   python equity_color_coding.py JP-June_Working.xlsx
   ```
   Check the printed block summary against these known-correct June numbers:

   | Tab | Segment | POF | Single | Dual | Strength |
   |---|---|---|---|---|---|
   | DRY | Dry REP (rows 19-39) | 10 | 9 | – | 6 |
   | DRY | Dual (rows 60-80) | 8 | 9 | – | 7 |
   | DRY | Dry 16-34 (rows 101-121) | 9 | 6 | – | 14 |
   | DRY | Dry 35-49 (rows 142-162) | 13 | 8 | – | 8 |
   | DRY | Dry 50+ (rows 183-203) | 4 | 6 | – | 7 |
   | WET | Wet REP (rows 19-29) | 3 | 3 | – | 5 |
   | WET | Dual (rows 50-60) | 3 | 3 | 2 | 4 |
   | WET | System (rows 81-91) | 3 | 5 | – | 5 |
   | WET | Dispo (rows 112-122) | 6 | 6 | – | 6 |
   | WET | Wet -34 (rows 143-153) | 7 | 5 | – | 5 |
   | WET | Wet 35-49 (rows 174-184) | 4 | 4 | – | 5 |
   | WET | Wet 50+ (rows 205-215) | 4 | 4 | – | 5 |

   Total: 124 cells colored in DRY SCORECARD, 97 in WET SCORECARD (221 overall).
   Open the file and visually confirm the `Equity Attributes (Normalized)` sections
   (rows 19+ within each segment) show green/navy/blue/light-blue/red fills matching
   the `Instructions` tab's legend in the `Equity_ownership_calculator` file.

If all eight steps check out, the replacement is confirmed working and you're done —
no need to re-run everything again unless you change the scripts.

---

## Part 3 — Monthly runbook (repeat every month going forward)

1. **Get the two files from the client**: the updated report workbook (e.g.
   `JP-July.xlsx`) and the raw data export (e.g. `July-Rawdata.xlsx`). Place both in
   the same folder as the nine scripts.

2. **Update the default paths** (optional, saves typing) — open `run_monthly_pipeline.py`
   and edit these two lines near the top to the new month's filenames:
   ```python
   DEFAULT_CLIENT_FILE = r"C:\JP KR Grooming\Python\JP-July.xlsx"
   DEFAULT_RAW_DATA_FILE = r"C:\JP KR Grooming\Python\July-Rawdata.xlsx"
   ```
   Or skip this and pass them on the command line instead:
   ```
   python run_monthly_pipeline.py --client-file "JP-July.xlsx" --raw-data-file "July-Rawdata.xlsx"
   ```

3. **Run it:**
   ```
   python run_monthly_pipeline.py
   ```
   Enter the new month at the prompt, e.g. `2026-07`.

4. **Read the console output.** A few things are worth actually reading, not just
   scrolling past:
   - **`Report month: ... -> exact stamped date found in Respondent Data: ...`** —
     if this raises an error instead ("No rows found in Respondent Data for..."),
     the client's file doesn't have that month's data yet — stop and check with them
     before proceeding.
   - **Any `WARNING: column ... isn't one of the known brands`** — this means a
     column matched the Rating question pattern but isn't Braun/Philips/Panasonic/
     Gillette/Schick/Kai. It won't be touched automatically; check whether it's
     meaningful (a new brand added?) or the same harmless `[[ABA_7_SP]]` artifact
     seen in June.
   - **`Integrity check: 0 unexpected parts lost, N sheet(s) changed`** — should
     always read exactly this. If a number appears where "0 unexpected" should be,
     stop and don't proceed to Excel — something has gone wrong with the write step.

5. **Recalculate manually** (this step will very likely fail automatically and print
   clear instructions — that's expected on this machine, confirmed to be a security
   policy boundary, not a bug):
   - Open the `_Working` file
   - `Ctrl+Alt+Shift+F9`, then `Ctrl+S`

6. **Apply equity color coding** — must run after the manual recalc above, since it
   reads live calculated values:
   ```
   python equity_color_coding.py JP-July_Working.xlsx
   ```
   Reads the printed block summary — should show one block per scorecard segment
   (5 for DRY, 7 for WET this template), each with a POF/Single/Dual/Shared/Strength
   breakdown. No fixed numbers to check against for a new month (the breakdown will
   differ from June's), but the structure — 5 DRY blocks, 7 WET blocks — should stay
   the same unless the client's template itself changes.

7. **Spot-check the new month's numbers** the same way as Part 2, step 6 — sample
   sizes should be in a similar range to prior months (roughly 300 respondents split
   across DRY/WET categories), and watch for any segment with a suspiciously small N
   the way `Dispo` was flagged this month. Also glance at the colored cells — mostly
   green/red/light-blue is normal; if a block comes back with zero colored cells at
   all, something's likely wrong with that segment's data.

8. **Deliver `JP-July_Working.xlsx`** (or whatever that month is named) as the
   finished report. The original client file stays untouched in the same folder,
   available for reference.

---

## Quick troubleshooting index

| Symptom | Likely cause / fix |
|---|---|
| `pip` not recognized | Use `python -m pip install ...` instead |
| Repair dialog on open | Re-run from scratch — should not happen with the current scripts; if it does, stop and report the exact repair-log message (File → the repair popup names the specific part) |
| Save As dialog / sharing violation during auto-recalc | Expected on this machine — use the manual `Ctrl+Alt+Shift+F9` + `Ctrl+S` fallback |
| `Client file not found` / `Raw data file not found` | Check the path is correct and, if on a network drive, that it's actually connected |
| `No rows found in Respondent Data for YYYY-MM` | That month's data isn't in the client's file yet |
| Script exits during `Post-save validation` with a mismatch warning | Some current-month respondents in `Respondent Data` have no matching row in the raw data file, or vice versa — worth a manual look before trusting the output |
| Colored cells look wrong, or a block shows zero colored cells | Almost always means `equity_color_coding.py` ran before the manual recalc step — it reads live calculated values, which are blank until Excel has recalculated. Redo the recalc, then re-run the color script |
