# JP-KR Grooming Report Pipeline — Complete Runbook (JP + KO)

Everything needed to run both markets' monthly reports, end to end, on this machine.

---

## Part 1 — Folder & file structure

All 17 Python scripts live in **one folder**, regardless of market — Python resolves
imports by looking next to the script you run, so splitting them up would break
every shared import.

```
C:\JP KR Grooming\Python\
│
│  --- JP-specific entry points ---
├── python run_monthly_pipeline_JP.py --skip-recalc
├── recalculate_via_macro_JP.py
├── equity_color_coding_JP.py
├── debug_recalc_visible_JP.py         (debug tool only)
├── AutoRecalcRunner.xlsm              (JP's macro file — NOT a .py, lives here directly)
│
│  --- KO-specific entry points ---
├── run_monthly_pipeline_KO.py
├── recalculate_via_macro_KO.py
├── equity_color_coding_KO.py
│
│  --- shared modules (both markets import these; never rename or move) ---
├── pipeline_config.py
├── market_config.py
├── safe_xlsx_writer.py
├── build_idxage.py
├── clean_ratings.py
├── build_jq_tables.py
├── set_report_month.py
├── recalculate_excel.py
├── equity_color_coding.py
├── test_excel_smoke.py                (debug tool only)
│
│  --- JP's data folders ---
├── Report\        (client's monthly JP report file — exactly one file, any name)
├── RawData\        (JP raw survey export — exactly one file, any name)
├── Formatting\     (reference only — Equity_ownership_calculator, not read by any script)
│
└── KO\             (everything Korea-specific lives under here)
    ├── AutoRecalcRunner_KO.xlsm       (KO's macro file — NOT inside KO\Report\)
    ├── Report\      (client's monthly KO report file)
    ├── RawData\     (KO raw survey export)
    └── Formatting\  (reference only)
```

**Never place `AutoRecalcRunner.xlsm` / `AutoRecalcRunner_KO.xlsm` inside their
respective `Report\` folders** — both pipelines scan `Report\` expecting exactly one
file, and the macro file would be silently counted as a second candidate.

---

## Part 2 — One-time machine setup

1. **Python + dependencies**
   ```
   python --version
   python -m pip install openpyxl xlwings lxml
   python -c "import openpyxl, xlwings, lxml; print('all OK')"
   ```

2. **Excel smoke test** (shared — covers both markets, since it just checks this
   machine's Excel/COM capability generally, not market-specific data)
   ```
   python test_excel_smoke.py
   ```
   Expect `PASS`, with `F1 = 4.0 (expected: 4)` as the critical line.

3. **Set up both macro runner files** — `AutoRecalcRunner.xlsm` (JP) and
   `AutoRecalcRunner_KO.xlsm` (KO). Full VBA source, exact paste locations, and the
   "how to safely edit without it auto-running and closing on you" technique are in
   `AutoRecalcRunner_VBA_source.txt` (JP) and `AutoRecalcRunner_KO_VBA_source.txt` (KO).

   Quick summary of the safe-editing method that actually worked on this machine
   (Shift-bypass and untrusting the folder both failed — this did work): open a
   blank workbook, `Alt+F11`, `Ctrl+G`, type `Application.EnableEvents = False`,
   Enter — *then* open the runner file via File → Open from that same session.

   Confirm Trusted Locations covers both: File → Options → Trust Center → Trust
   Center Settings → Trusted Locations. If `C:\JP KR Grooming\Python` is trusted
   with "Subfolders of this location are also trusted" checked, `KO\` is already
   covered automatically.

4. **Test each runner in isolation** before relying on it as part of a real run:
   ```
   python recalculate_via_macro_JP.py
   python recalculate_via_macro_KO.py
   ```
   (Each will fail with "no *_Working.xlsx file found" if `Report\` is empty right
   now — that's expected; the real test happens as part of Part 3/4 below.)

---

## Part 3 — Monthly workflow: JP

**Confirmed reliable — five-plus consecutive clean runs on this machine.**

1. Drop the client's report file into `Report\`, raw data into `RawData\` — any
   filenames, exactly one file each. No need to delete last month's `_Working` file
   first; it's overwritten automatically. Just make sure nothing is open in Excel.

2. ```
   python run_monthly_pipeline_JP.py --skip-recalc
   ```
   Enter the report month when prompted (e.g. `2026-07`).

   **Read the output, don't just scroll past:**
   - `Integrity check: 0 unexpected parts lost, N sheet(s) changed` — should always
     read exactly this
   - Any `WARNING: column ... isn't one of the known brands` — check whether it's
     the familiar harmless `[[ABA_7_SP]]` artifact or something genuinely new
   - Respondent cross-check should show an exact match

3. ```
   python recalculate_via_macro_JP.py
   ```
   Expect `Result: SUCCESS: <filename> recalculated and saved at ...` naming the
   file you expect, followed by the error scan and sample-size check running
   automatically. If it fails for any reason: open the `_Working` file in `Report\`,
   `Ctrl+Alt+Shift+F9`, `Ctrl+S` — proven reliable as a fallback every time it's
   been needed.

4. ```
   python equity_color_coding_JP.py
   ```

5. **Open the file and actually look at it** — no repair dialog, `DRY SCORECARD`/
   `WET SCORECARD` numbers look plausible, colors match the legend in the
   `Equity_ownership_calculator` Instructions tab. Nothing in the console output
   substitutes for this.

6. Deliver the `_Working` file from `Report\`.

**June reference numbers** (this exact data, useful for confirming a script change
didn't break anything — not applicable to a genuinely new month's real data):

| Check | Value |
|---|---|
| Respondents | 300 |
| Rating cells fixed | 12,178 (Braun 2361, Philips 1943, Panasonic 2318, Gillette 2045, Schick 2072, Kai 1439) |
| J:Q (App / BR) | 5559→4184 / 5128→3914 |
| Error scan | 89 total (WET SCORECARD 2× `#DIV/0!`, AE(WET) 87× `#VALUE!`) |
| Sample-size flag | WET SCORECARD / Dispo, n=6 |
| Colored cells | DRY 124, WET 97 (221 total) |

---

## Part 4 — Monthly workflow: KO

**Confirmed reliable — clean on the first real-hardware attempt.** Same structure as
JP throughout, just the KO-suffixed scripts and the `KO\` folder tree.

1. Drop the client's report file into `KO\Report\`, raw data into `KO\RawData\`.

2. ```
   python run_monthly_pipeline_KO.py --skip-recalc
   ```
   Enter the report month when prompted. Same output-reading guidance as JP step 2
   — additionally, KO's known unmapped-column artifact appears as column `OZ`
   rather than `PD` (different position, same `[[ABA_7_SP]]` placeholder, not a
   new issue).

3. ```
   python recalculate_via_macro_KO.py
   ```
   Same expectations and fallback as JP step 3. Note: KO's `AE(WET)` error count
   will vary somewhat between scans depending on recalculation state (its
   respondent list is driven by a live `FILTER` formula off cell A1) — this is
   expected, not a bug; `DRY SCORECARD`/`WET SCORECARD` error cells stay stable
   regardless.

4. ```
   python equity_color_coding_KO.py
   ```

5. Open the file and look at it — same as JP step 5.

6. Deliver the `_Working` file from `KO\Report\`.

**June reference numbers:**

| Check | Value |
|---|---|
| Respondents | 300 |
| Rating cells fixed | 8,957 (Gillette 2038, Schick 1021, Dorco 1558, Braun 1457, Philips 1780, Panasonic 1103) |
| J:Q (App / BR) | 4105→3126 / 5846→3896 |
| Error scan (post-recalc) | 550 total (DRY SCORECARD 1× `#DIV/0!`, WET SCORECARD 9× `#VALUE!`, AE(WET) 540× `#VALUE!`) |
| Sample-size flag | none this month |
| Colored cells | DRY 150, WET 101 (251 total) |

---

## Part 5 — Troubleshooting index (both markets)

| Symptom | Likely cause / fix |
|---|---|
| `pip` not recognized | Use `python -m pip install ...` instead |
| Repair dialog on open | Should not happen with current scripts — if it does, report the exact repair-log message |
| `Client file not found` / `Expected exactly one file, found N` | Check the correct `Report\` (JP) or `KO\Report\` (KO) folder — a stray file, including an old `_Working.xlsx` sitting alongside a new original, is the usual cause |
| `No rows found in Respondent Data for YYYY-MM` | That month's data isn't in the client's file yet |
| A Save As dialog with a garbled filename appears | This is the internal `xlwings` COM attempt inside `run_monthly_pipeline_JP/KO.py` when run *without* `--skip-recalc` — always use `--skip-recalc`, then the separate `recalculate_via_macro_*.py` script |
| `recalculate_via_macro_*.py` times out or fails | Check Task Manager for a lingering `EXCEL.EXE`; check the runner `.xlsm` is set up and trusted; fall back to manual `Ctrl+Alt+Shift+F9` + `Ctrl+S` |
| `recalculate_via_macro_*.py` reports "Found N files matching *_Working.xlsx" | Remove the extra file(s) named in the message |
| Colored cells look wrong, or zero cells colored | `equity_color_coding_*.py` ran before recalculation completed — recalculate first, then re-run color coding (safe to re-run; fully idempotent, proven byte-identical on repeat runs) |
| Need to recalculate a specific file manually, bypassing auto-discovery | Open the relevant `AutoRecalcRunner*.xlsm`, `Alt+F8` → `RecalcWithFilePicker` → `Run` |
