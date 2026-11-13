"""
Shared configuration for the monthly pipeline - folder locations and the
"find the one file in this folder" helper, used by both
run_monthly_pipeline.py and equity_color_coding.py. Kept in one place so
the two scripts (deliberately kept separate, not merged) can't drift out
of sync on where files actually live.
"""
from pathlib import Path

# Each folder must contain exactly one file (any name) for the current
# month - no filenames to update in code each month, just replace the
# file in the folder.
REPORT_FOLDER = r"C:\JP KR Grooming\Python\Report"
RAW_DATA_FOLDER = r"C:\JP KR Grooming\Python\RawData"

# Reference only - holds Equity_ownership_calculator_-_Raw.xlsm for the
# user's own audit trail. No script reads from this folder: the
# classification formula and color table it documents were extracted once,
# verified against the file's own real output, and now live hardcoded in
# equity_color_coding.py. Defined here for completeness/documentation of
# the folder convention, not because anything opens it.
FORMATTING_FOLDER = r"C:\JP KR Grooming\Python\Formatting"


def find_single_file_in_folder(folder_path, extensions=(".xlsx", ".xlsm"), name_filter=None):
    """Look in a folder for exactly one file matching the given extensions
    and return its path. Excel lock files (~$whatever.xlsx, created while
    a file is open) are always excluded - otherwise a folder with one real
    file open in Excel would look like it has two files and fail
    confusingly. name_filter, if given, is a callable(filename) -> bool
    for narrowing further (e.g. only "*_Working.xlsx" files).
    Fails clearly if the folder is missing, empty, or has more than one
    candidate - silently picking the wrong file is worse than an error."""
    folder = Path(folder_path)
    if not folder.exists():
        raise FileNotFoundError(
            f"Folder not found: {folder}\n"
            f"Check the path is correct and this machine can reach it."
        )
    if not folder.is_dir():
        raise NotADirectoryError(f"Expected a folder, not a file: {folder}")

    candidates = [
        f for f in folder.iterdir()
        if f.is_file() and f.suffix.lower() in extensions and not f.name.startswith("~$")
    ]
    if name_filter is not None:
        candidates = [f for f in candidates if name_filter(f.name)]

    if not candidates:
        extra = " matching the expected name pattern" if name_filter else ""
        raise FileNotFoundError(
            f"No file found in {folder}{extra} (looked for {', '.join(extensions)}). "
            f"Make sure this month's file has been placed there."
        )
    if len(candidates) > 1:
        names = ", ".join(f.name for f in candidates)
        raise RuntimeError(
            f"Expected exactly one file in {folder}, found {len(candidates)}: {names}\n"
            f"Remove the extra file(s) so there's no ambiguity about which one to use."
        )
    return candidates[0]
