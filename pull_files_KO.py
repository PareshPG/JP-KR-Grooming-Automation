"""
Pulls this month's flat file and raw data file from the OneDrive-synced
SharePoint location into the local pipeline folders - KOREA.

Run this BEFORE the pipeline scripts each month. It does NOT run any
pipeline step itself; it only handles getting the right files into the
right local folders, cleanly (each destination folder is emptied before
the new file is copied in, per the explicit instruction that these
folders must never contain more than the current month's file).

FOLDER LAYOUT THIS ASSUMES (see market_config.py for the single source of
truth on these paths):
  OneDrive (source, read-only - never written to):
    .../Data Compass - BHT/Flat data/<Month Year>/   <- new subfolder each month
    .../Data Compass - BHT/Raw data/                  <- no month subfolders

  Local (destination):
    C:\\JP KR Grooming\\Python\\JP\\Flat File\\   <- cleaned + refreshed each run
    C:\\JP KR Grooming\\Python\\JP\\RawData\\     <- cleaned + refreshed each run
    C:\\JP KR Grooming\\Python\\JP\\Report\\      <- NOT touched by this script at all
"""
import shutil
import re
from datetime import datetime
from pathlib import Path

ONEDRIVE_FLAT_DATA_ROOT = (
    r"C:\Users\lad.p.1\OneDrive - Procter and Gamble\Data Compass - BHT\Flat data"
)
ONEDRIVE_RAW_DATA_ROOT = (
    r"C:\Users\lad.p.1\OneDrive - Procter and Gamble\Data Compass - BHT\Raw data"
)

LOCAL_FLAT_FILE_FOLDER = r"C:\JP KR Grooming\Python\KO\Flat File"
LOCAL_RAW_DATA_FOLDER = r"C:\JP KR Grooming\Python\KO\RawData"

FLAT_FILE_KEYWORD = "KO"
RAW_DATA_KEYWORD = "Korea"

MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July",
               "August", "September", "October", "November", "December"]


def parse_month_folder(month_folder_name):
    """Parse a OneDrive folder name like 'Jul 2026' into (month, year).
    Accepts common abbreviated or full month name + 4-digit year, separated
    by a space - matches every example seen so far. Raises a clear error
    on anything else, rather than guessing."""
    m = re.match(r"^([A-Za-z]+)\s+(\d{4})$", month_folder_name.strip())
    if not m:
        raise ValueError(
            f"Couldn't parse month folder name {month_folder_name!r} - "
            f"expected a format like 'Jul 2026' (month name + space + 4-digit year)."
        )
    month_str, year_str = m.group(1), m.group(2)
    try:
        parsed = datetime.strptime(month_str[:3], "%b")
    except ValueError:
        raise ValueError(
            f"Couldn't recognize {month_str!r} as a month name in {month_folder_name!r}."
        )
    return parsed.month, int(year_str)


def expected_data_month(folder_month, folder_year):
    """The OneDrive folder is named ONE MONTH AHEAD of the data it
    contains - a 'Jul 2026' folder holds June 2026's report data. This
    computes that actual data month, handling the December -> January
    year rollover correctly."""
    if folder_month == 1:
        return 12, folder_year - 1
    return folder_month - 1, folder_year


def check_filename_matches_month(filename, month, year, label):
    """Verifies filename contains the expected month name (e.g. 'June')
    for the flat file, or month+2-digit-year (e.g. 'June26') for the raw
    file - both patterns confirmed consistent across every file seen in
    this project so far. Raises clearly rather than silently proceeding
    on a mismatch, since a wrong-month file processed silently produces a
    wrong report with no visible error until someone reads the numbers."""
    month_name = MONTH_NAMES[month - 1]
    year_2digit = str(year)[-2:]
    full_pattern = month_name
    abbreviated_pattern = f"{month_name}{year_2digit}"
    if full_pattern.lower() in filename.lower() or abbreviated_pattern.lower() in filename.lower():
        return
    raise ValueError(
        f"{label} filename doesn't appear to match the expected data month "
        f"({month_name} {year}): {filename!r}\n"
        f"Expected to find '{full_pattern}' or '{abbreviated_pattern}' somewhere "
        f"in the filename. This usually means the wrong month's file is "
        f"present in the source folder - confirm with the client before proceeding."
    )


def find_file_with_keyword(folder_path, keyword, extensions=(".xlsx", ".xlsm")):
    """Find exactly one file in folder_path whose name contains keyword
    (case-insensitive). Fails loudly - never guesses - if zero or more
    than one match is found, same principle as find_single_file_in_folder
    in pipeline_config.py. Excel lock files (~$...) are always excluded."""
    folder = Path(folder_path)
    if not folder.exists():
        raise FileNotFoundError(
            f"OneDrive folder not found: {folder}\n"
            f"Check the path is correct, the month folder name is spelled "
            f"exactly as it appears on OneDrive, and this folder has finished "
            f"syncing (a cloud-only file can look present but not be readable yet)."
        )
    if not folder.is_dir():
        raise NotADirectoryError(f"Expected a folder, not a file: {folder}")

    candidates = [
        f for f in folder.iterdir()
        if f.is_file()
        and f.suffix.lower() in extensions
        and not f.name.startswith("~$")
        and keyword.lower() in f.name.lower()
    ]

    if not candidates:
        all_files = [f.name for f in folder.iterdir() if f.is_file()]
        raise FileNotFoundError(
            f"No file containing '{keyword}' found in {folder}.\n"
            f"Files actually present there: {all_files or '(none)'}"
        )
    if len(candidates) > 1:
        names = ", ".join(f.name for f in candidates)
        raise RuntimeError(
            f"Expected exactly one file containing '{keyword}' in {folder}, "
            f"found {len(candidates)}: {names}\n"
            f"Remove the extra file(s) on OneDrive, or narrow the keyword."
        )
    return candidates[0]


def clean_and_copy(source_file, dest_folder_path):
    """Empty dest_folder_path completely, then copy source_file into it.
    The 'empty it first' step is deliberate and explicit per instruction -
    these folders must never hold more than the current month's single file,
    otherwise the pipeline's own 'exactly one file' folder-discovery would
    immediately fail with a false ambiguity error."""
    dest = Path(dest_folder_path)
    dest.mkdir(parents=True, exist_ok=True)

    existing = [f for f in dest.iterdir() if f.is_file()]
    for f in existing:
        f.unlink()
    if existing:
        print(f"  Cleared {len(existing)} existing file(s) from {dest}")

    dest_file = dest / source_file.name
    shutil.copy2(source_file, dest_file)
    print(f"  Copied {source_file.name} -> {dest}")
    return dest_file


def pull_files_for_month(month_folder_name):
    """month_folder_name: exactly as it appears on OneDrive, e.g. 'Jul 2026'.
    Note this folder is named ONE MONTH AHEAD of the data it contains -
    'Jul 2026' holds June 2026's report data. Both files found are checked
    against that expected data month before copying, to catch a stale or
    wrong-month file with a clear error instead of silently processing it."""
    folder_month, folder_year = parse_month_folder(month_folder_name)
    data_month, data_year = expected_data_month(folder_month, folder_year)
    print(f"Pulling files for: {month_folder_name} "
          f"(expected data month: {MONTH_NAMES[data_month-1]} {data_year})")
    print()

    print("Flat file:")
    flat_month_folder = Path(ONEDRIVE_FLAT_DATA_ROOT) / month_folder_name
    flat_file = find_file_with_keyword(flat_month_folder, FLAT_FILE_KEYWORD)
    print(f"  Found: {flat_file}")
    check_filename_matches_month(flat_file.name, data_month, data_year, "Flat file")
    clean_and_copy(flat_file, LOCAL_FLAT_FILE_FOLDER)
    print()

    print("Raw data file:")
    raw_file = find_file_with_keyword(ONEDRIVE_RAW_DATA_ROOT, RAW_DATA_KEYWORD)
    print(f"  Found: {raw_file}")
    check_filename_matches_month(raw_file.name, data_month, data_year, "Raw data file")
    clean_and_copy(raw_file, LOCAL_RAW_DATA_FOLDER)
    print()

    print("File pull complete. Report template folder was not touched.")


if __name__ == "__main__":
    month = input("Enter the OneDrive month folder name exactly as it appears (e.g. 'Jul 2026'): ").strip()
    pull_files_for_month(month)
