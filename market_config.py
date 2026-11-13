"""
Market-specific configuration - the single source of truth for everything
that differs between JP and KO (and any future market). Every other script
either imports these values directly, or accepts them as parameters with
JP's values as the default (so existing JP behavior is completely
unchanged unless a script is explicitly called with the KO profile).

Verified directly against both real files before being encoded here - see
conversation history for the cross-checks (e.g. KO's age question was
confirmed by matching 300/300 respondents against the Age sheet's existing
ground-truth values, not assumed from column position alone).
"""
from dataclasses import dataclass, field
from typing import List, Dict


@dataclass
class MarketConfig:
    name: str
    report_folder: str
    raw_data_folder: str
    formatting_folder: str

    # Age lookup sheet - JP calls it "IDxAGE" with the age duplicated into
    # both columns B and C; KO calls it "Age" with a single column B only.
    age_sheet_name: str
    age_sheet_has_duplicate_column: bool

    # Raw data file: the exact header text of the age question, and its
    # column position as a hint only (headers are still located by text
    # match, not by this position, in case the layout shifts again).
    age_question_header: str
    # Second age question, asked of 16-17 year-olds (the first one then holds the PARENT's age).
    # Where it holds a number it overrides age_question_header for that respondent.
    minor_age_question_header: str = ""

    # Known brand names for Rating-column detection in Respondent Data -
    # only brands in this list get their rating cells cleaned; anything
    # else matching the rating-question pattern is reported but left
    # untouched (e.g. JP's and KO's shared [[ABA_7_SP]] artifact).
    known_brands: List[str] = field(default_factory=list)

    # Raw-file brand names (local language, as a raw column header STARTS with them) -> English brand name, used by the
    # flat-vs-raw alignment check (flat_raw_alignment.py). Empty = that check is skipped for the market.
    raw_brand_names: Dict[str, str] = field(default_factory=dict)

    # Sheets recalculate_excel.py's error scan / sample-size check look at -
    # only differs from the sheet-name substitution above (IDxAGE -> Age).
    in_scope_sheets: List[str] = field(default_factory=list)

    # Runner file + result log used by the macro-based recalculation
    # approach - kept separate per market so a KO run can never be
    # confused with a JP run mid-flight.
    runner_path: str = ""
    recalc_log_path: str = ""
    # Only used by markets on the new flat-file + fixed-template
    # architecture (currently JP only) - empty string for markets still on
    # the older combined-file architecture.
    flat_file_folder: str = ""


JP = MarketConfig(
    name="JP",
    report_folder=r"C:\JP KR Grooming\Python\JP\Report",
    raw_data_folder=r"C:\JP KR Grooming\Python\JP\RawData",
    formatting_folder=r"C:\JP KR Grooming\Python\Formatting",
    flat_file_folder=r"C:\JP KR Grooming\Python\JP\Flat File",
    age_sheet_name="IDxAGE",
    age_sheet_has_duplicate_column=True,
    age_question_header="あなたの年齢は？",
    minor_age_question_header="あなたは何歳ですか？",
    known_brands=["Braun", "Philips", "Panasonic", "Gillette", "Schick", "Kai"],
    raw_brand_names={"ジレット": "Gillette", "シック": "Schick", "カイ": "Kai", "貝印": "Kai", "ブラウン": "Braun",
                     "フィリップス": "Philips", "パナソニック": "Panasonic"},
    in_scope_sheets=[
        "DRY SCORECARD", "AE(DRY)", "Worth it & UBA_App", "Respondent Data",
        "WET SCORECARD", "AE(WET)", "Worth it & UBA_BR", "IDxAGE",
    ],
    runner_path=r"C:\JP KR Grooming\Python\AutoRecalcRunner.xlsm",
    recalc_log_path=r"C:\JP KR Grooming\Python\recalc_result.txt",
)

KO = MarketConfig(
    name="KO",
    report_folder=r"C:\JP KR Grooming\Python\KO\Report",
    raw_data_folder=r"C:\JP KR Grooming\Python\KO\RawData",
    formatting_folder=r"C:\JP KR Grooming\Python\KO\Formatting",
    flat_file_folder=r"C:\JP KR Grooming\Python\KO\Flat File",
    # The KR template reads respondent ages from the tab called "Sheet1" (Respondent Data's old age formula was
    # XLOOKUP(id, Sheet1!A:A, Sheet1!B:B)). The template also carries an "IDxAGE" tab that nothing reads (a stale copy
    # with scratch formulas of its own) - the pipeline leaves it alone.
    age_sheet_name="Sheet1",
    age_sheet_has_duplicate_column=False,
    age_question_header="당신의 나이는 몇 살입니까?",
    minor_age_question_header="당신의 나이는 몇 세입니까?",
    known_brands=["Braun", "Philips", "Panasonic", "Gillette", "Schick", "Dorco"],
    raw_brand_names={"질레트": "Gillette", "쉬크": "Schick", "도르코": "Dorco", "도루코": "Dorco", "브라운": "Braun",
                     "필립스": "Philips", "파나소닉": "Panasonic"},
    in_scope_sheets=[
        "DRY SCORECARD", "AE(DRY)", "Worth it & UBA_App", "Respondent Data",
        "WET SCORECARD", "AE(WET)", "Worth it & UBA_BR", "Sheet1",
    ],
    runner_path=r"C:\JP KR Grooming\Python\KO\AutoRecalcRunner_KO.xlsm",
    recalc_log_path=r"C:\JP KR Grooming\Python\KO\recalc_result.txt",
)

MARKETS = {"JP": JP, "KO": KO}