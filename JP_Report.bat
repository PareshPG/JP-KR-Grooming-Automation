@echo off
setlocal

cd /d "C:\JP KR Grooming\Python"

echo ============================================================
echo  JP Monthly Report Pipeline
echo ============================================================
echo.
echo  Before continuing, confirm:
echo    - JP\Flat File\  has this quarter's flat file
echo    - JP\RawData\    has this quarter's raw data file
echo    - JP\Report\     has exactly one template (no leftover _Working file)
echo.
pause

echo.
echo ============================================================
echo  STEP 1: Build the report (you will be asked for the month)
echo ============================================================
python run_monthly_pipeline_JP.py
if errorlevel 1 (
    echo.
    echo [STOPPED] run_monthly_pipeline_JP.py reported an error above.
    echo Recalculation was NOT started. See SOP Section 8 (Troubleshooting).
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  STEP 2: Recalculate
echo ============================================================
python recalculate_via_macro_JP.py
if errorlevel 1 (
    echo.
    echo [STOPPED] recalculate_via_macro_JP.py reported an error above.
    echo See SOP Section 8 (Troubleshooting).
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  DONE. Next:
echo    1. Open the working file in JP\Report\ and do the final
echo       check (SOP Section 5.6).
echo    2. Archive it out of JP\Report\ (SOP Section 5.7) BEFORE
echo       running this again for the next month - the next run
echo       will auto-delete any _Working file still sitting there.
echo ============================================================
pause
