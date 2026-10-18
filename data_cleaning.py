"""
data_cleaning.py
-----------------
Loads student_performance_raw.xlsx, fixes every category of data-quality issue,
and saves a clean copy as student_performance_cleaned.xlsx.

Returns a CleaningResult dataclass so other modules can use the stats without
re-running the pipeline.
"""

import re
import sys
import warnings
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

warnings.filterwarnings("ignore")

# ── Paths ──────────────────────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parents[1]
RAW_FILE     = ROOT / "data" / "student_performance_raw.xlsx"
CLEANED_FILE = ROOT / "data" / "student_performance_cleaned.xlsx"

# ── Valid ranges ───────────────────────────────────────────────────────────────
AGE_MIN, AGE_MAX = 15, 40
ATTEND_MIN, ATTEND_MAX = 0, 100
ASSIGN_MIN, ASSIGN_MAX = 0, 20
INTERN_MIN, INTERN_MAX = 0, 30
FINAL_MIN,  FINAL_MAX  = 0, 50
TOTAL_MIN,  TOTAL_MAX  = 0, 100

# ── Canonical mappings ─────────────────────────────────────────────────────────
DEPT_MAP = {
    "cse": "CSE", "ece": "ECE", "me": "ME",
    "ce": "CE",   "eee": "EEE",
}
GENDER_MAP = {
    "male": "Male", "m": "Male",
    "female": "Female", "f": "Female",
}
CITY_MAP = {
    "bombay":    "Mumbai",
    "new delhi": "Delhi",
    "bangalore": "Bengaluru",
    "calcutta":  "Kolkata",
}
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# ── Result container ───────────────────────────────────────────────────────────
@dataclass
class CleaningResult:
    original_rows: int = 0
    final_rows: int = 0
    original_cols: int = 0
    missing_before: int = 0
    missing_after: int = 0
    duplicates_removed: int = 0
    invalid_values: int = 0
    corrected_categorical: int = 0
    invalid_emails: int = 0
    log: list[dict] = field(default_factory=list)
    df_cleaned: pd.DataFrame = field(default_factory=pd.DataFrame)

    def add_log(self, step: str, column: str, action: str, count: int) -> None:
        self.log.append({"Step": step, "Column": column,
                         "Action": action, "Records_Affected": count})

    def print_summary(self) -> None:
        print("\n" + "─" * 60)
        print("  DATA CLEANING SUMMARY")
        print("─" * 60)
        print(f"  Original rows          : {self.original_rows}")
        print(f"  Final rows             : {self.final_rows}")
        print(f"  Columns                : {self.original_cols}")
        print(f"  Missing values (before): {self.missing_before}")
        print(f"  Missing values (after) : {self.missing_after}")
        print(f"  Duplicates removed     : {self.duplicates_removed}")
        print(f"  Invalid values flagged : {self.invalid_values}")
        print(f"  Categorical corrections: {self.corrected_categorical}")
        print(f"  Invalid emails flagged : {self.invalid_emails}")
        print("─" * 60)


# ══════════════════════════════════════════════════════════════════════════════
#  A. LOAD
# ══════════════════════════════════════════════════════════════════════════════

def load_dataset(path: Path = RAW_FILE) -> pd.DataFrame:
    if not path.exists():
        sys.exit(f"ERROR: Raw dataset not found at {path}\n"
                 "       Run generate_dataset.py first.")
    df = pd.read_excel(path, engine="openpyxl")
    print(f"\n[A] Dataset loaded: {df.shape[0]} rows × {df.shape[1]} columns")
    print(f"    Columns: {list(df.columns)}")
    return df


# ══════════════════════════════════════════════════════════════════════════════
#  B. MISSING VALUES
# ══════════════════════════════════════════════════════════════════════════════

def handle_missing_values(df: pd.DataFrame, res: CleaningResult) -> pd.DataFrame:
    """Fill missing values using domain-appropriate strategies."""
    print("\n[B] Handling missing values …")

    # Numeric columns → median (robust to outliers)
    numeric_fill = {
        "Age":                   ("median", lambda s: round(s.median())),
        "Attendance_Percentage": ("median", lambda s: round(s.median(), 1)),
        "Assignment_Marks":      ("median", lambda s: round(s.median(), 1)),
        "Internal_Marks":        ("median", lambda s: round(s.median(), 1)),
        "Final_Exam_Marks":      ("median", lambda s: round(s.median(), 1)),
        "Total_Marks":           ("median", lambda s: round(s.median(), 1)),
    }
    for col, (method, fn) in numeric_fill.items():
        n_miss = df[col].isna().sum()
        if n_miss:
            fill_val = fn(df[col].dropna())
            df[col] = df[col].fillna(fill_val)
            res.add_log("Missing Values", col,
                        f"Filled {n_miss} NaN with {method} ({fill_val})", n_miss)
            print(f"    {col}: filled {n_miss} NaN → {method} = {fill_val}")

    # Categorical columns → mode or explicit label
    cat_fill = {
        "Gender":     ("mode",    None),
        "Department": ("mode",    None),
        "Grade":      ("mode",    None),
        "City":       ("Unknown", None),
        "Email":      ("flag",    None),  # keep as NaN; flagged as invalid later
    }
    for col, (strategy, _) in cat_fill.items():
        n_miss = df[col].isna().sum()
        if not n_miss:
            continue
        if strategy == "mode":
            mode_val = df[col].dropna().mode()
            if len(mode_val):
                df[col] = df[col].fillna(mode_val.iloc[0])
                res.add_log("Missing Values", col,
                            f"Filled {n_miss} NaN with mode ({mode_val.iloc[0]})",
                            n_miss)
                print(f"    {col}: filled {n_miss} NaN → mode = {mode_val.iloc[0]}")
        elif strategy == "Unknown":
            df[col] = df[col].fillna("Unknown")
            res.add_log("Missing Values", col,
                        f"Filled {n_miss} NaN with 'Unknown'", n_miss)
            print(f"    {col}: filled {n_miss} NaN → 'Unknown'")
        elif strategy == "flag":
            # Retain NaN – will be identified in email validation step
            print(f"    {col}: {n_miss} NaN retained (flagged in email step)")

    # Student_ID and Student_Name: do NOT fill – keep unique identifiers intact
    for col in ("Student_ID", "Student_Name"):
        n = df[col].isna().sum()
        if n:
            print(f"    {col}: {n} NaN left unfilled (must remain unique)")

    return df


# ══════════════════════════════════════════════════════════════════════════════
#  C. DUPLICATE REMOVAL
# ══════════════════════════════════════════════════════════════════════════════

def remove_duplicates(df: pd.DataFrame, res: CleaningResult) -> pd.DataFrame:
    """Remove exact duplicate rows; report duplicate Student_IDs separately."""
    print("\n[C] Removing duplicates …")

    before = len(df)
    df = df.drop_duplicates()
    removed = before - len(df)
    res.duplicates_removed = removed
    res.add_log("Duplicates", "All columns", f"Removed {removed} exact duplicates",
                removed)
    print(f"    Exact duplicates removed: {removed}")

    # Report duplicate Student_IDs (informational – do not auto-delete)
    dup_ids = df[df["Student_ID"].duplicated(keep=False)]["Student_ID"].unique()
    if len(dup_ids):
        print(f"    ⚠  Duplicate Student_IDs found (kept for review): {list(dup_ids)}")
    else:
        print("    No duplicate Student_IDs found after deduplication.")

    return df


# ══════════════════════════════════════════════════════════════════════════════
#  D. STANDARDIZE TEXT
# ══════════════════════════════════════════════════════════════════════════════

def _map_value(val: str | float, mapping: dict[str, str]) -> str:
    """Strip, lower, look up in mapping; return Title-cased original if not found."""
    if pd.isna(val):
        return val
    stripped = str(val).strip()
    return mapping.get(stripped.lower(), stripped.upper()
                       if len(stripped) <= 3 else stripped.title())


def standardize_text(df: pd.DataFrame, res: CleaningResult) -> pd.DataFrame:
    """Strip whitespace, normalize Department / Gender / City."""
    print("\n[D] Standardizing text fields …")

    # Strip all string columns
    str_cols = df.select_dtypes("object").columns
    df[str_cols] = df[str_cols].apply(
        lambda col: col.str.strip() if col.dtype == "object" else col
    )

    # Department
    before = df["Department"].copy()
    df["Department"] = df["Department"].apply(
        lambda v: _map_value(v, DEPT_MAP) if pd.notna(v) else v
    )
    # Force uppercase for 2-3 char codes
    df["Department"] = df["Department"].apply(
        lambda v: v.upper() if pd.notna(v) and len(str(v)) <= 3 else v
    )
    changed_dept = (before != df["Department"]).sum()
    res.corrected_categorical += changed_dept
    res.add_log("Standardize", "Department",
                f"Normalized {changed_dept} values", changed_dept)
    print(f"    Department: normalized {changed_dept} values")

    # Gender
    before = df["Gender"].copy()
    df["Gender"] = df["Gender"].apply(
        lambda v: _map_value(v, GENDER_MAP) if pd.notna(v) else v
    )
    changed_gender = (before != df["Gender"]).sum()
    res.corrected_categorical += changed_gender
    res.add_log("Standardize", "Gender",
                f"Normalized {changed_gender} values", changed_gender)
    print(f"    Gender    : normalized {changed_gender} values")

    # City
    before = df["City"].copy()
    df["City"] = df["City"].apply(
        lambda v: CITY_MAP.get(str(v).strip().lower(), str(v).strip().title())
        if pd.notna(v) else v
    )
    changed_city = (before != df["City"]).sum()
    res.corrected_categorical += changed_city
    res.add_log("Standardize", "City",
                f"Normalized {changed_city} values", changed_city)
    print(f"    City      : normalized {changed_city} values")

    # Student_Name → Title case
    df["Student_Name"] = df["Student_Name"].str.title()

    return df


# ══════════════════════════════════════════════════════════════════════════════
#  E. VALIDATE NUMERICAL VALUES
# ══════════════════════════════════════════════════════════════════════════════

def validate_numerics(df: pd.DataFrame, res: CleaningResult) -> pd.DataFrame:
    """Flag out-of-range values by setting them to NaN (documented action)."""
    print("\n[E] Validating numerical values …")

    checks = [
        ("Age",                   AGE_MIN,    AGE_MAX),
        ("Attendance_Percentage", ATTEND_MIN, ATTEND_MAX),
        ("Assignment_Marks",      ASSIGN_MIN, ASSIGN_MAX),
        ("Internal_Marks",        INTERN_MIN, INTERN_MAX),
        ("Final_Exam_Marks",      FINAL_MIN,  FINAL_MAX),
        ("Total_Marks",           TOTAL_MIN,  TOTAL_MAX),
    ]
    total_invalid = 0
    for col, lo, hi in checks:
        mask = (df[col] < lo) | (df[col] > hi)
        n = mask.sum()
        if n:
            df.loc[mask, col] = np.nan   # flag → NaN (do not invent a value)
            total_invalid += n
            res.add_log("Validate Numerics", col,
                        f"Flagged {n} out-of-range values → NaN", n)
            print(f"    {col}: {n} out-of-range values set to NaN "
                  f"(valid range {lo}–{hi})")

    # Re-fill flagged NaN values with the column median (post-flag fill)
    numeric_cols_to_refill = [
        "Age", "Attendance_Percentage", "Assignment_Marks",
        "Internal_Marks", "Final_Exam_Marks",
    ]
    for col in numeric_cols_to_refill:
        n_nan = df[col].isna().sum()
        if n_nan:
            med = round(df[col].median(), 1)
            df[col] = df[col].fillna(med)
            res.add_log("Post-flag Fill", col,
                        f"Refilled {n_nan} NaN with median ({med})", n_nan)

    # Recalculate Total_Marks to ensure consistency
    df["Total_Marks"] = (
        df["Assignment_Marks"] + df["Internal_Marks"] + df["Final_Exam_Marks"]
    ).round(1)

    res.invalid_values += total_invalid
    print(f"    Total invalid numeric values flagged: {total_invalid}")
    return df


# ══════════════════════════════════════════════════════════════════════════════
#  F. VALIDATE EMAIL
# ══════════════════════════════════════════════════════════════════════════════

def validate_emails(df: pd.DataFrame, res: CleaningResult) -> pd.DataFrame:
    """Flag malformed emails by replacing them with NaN."""
    print("\n[F] Validating email addresses …")
    # Include NaN as "invalid" for counting purposes
    invalid_mask = df["Email"].isna() | df["Email"].apply(
        lambda e: not EMAIL_RE.match(str(e)) if pd.notna(e) else True
    )
    n_invalid = invalid_mask.sum()
    df.loc[invalid_mask, "Email"] = np.nan
    res.invalid_emails = n_invalid
    res.invalid_values += n_invalid
    res.add_log("Validate Email", "Email",
                f"Flagged {n_invalid} invalid/missing emails → NaN", n_invalid)
    print(f"    Invalid / missing emails flagged: {n_invalid}")
    return df


# ══════════════════════════════════════════════════════════════════════════════
#  G. RECALCULATE GRADE
# ══════════════════════════════════════════════════════════════════════════════

def recalculate_grades(df: pd.DataFrame, res: CleaningResult) -> pd.DataFrame:
    """Recompute Grade from the cleaned Total_Marks."""
    def assign_grade(total: float) -> str:
        if pd.isna(total):
            return "Unknown"
        if total >= 90: return "A+"
        if total >= 80: return "A"
        if total >= 70: return "B"
        if total >= 60: return "C"
        if total >= 50: return "D"
        return "F"

    df["Grade"] = df["Total_Marks"].apply(assign_grade)
    res.add_log("Grade Recalc", "Grade", "Recomputed grades from cleaned Total_Marks",
                len(df))
    return df


# ══════════════════════════════════════════════════════════════════════════════
#  H. SAVE CLEANED DATASET
# ══════════════════════════════════════════════════════════════════════════════

_HEADER_FILL   = PatternFill("solid", fgColor="1F4E79")
_HEADER_FONT   = Font(bold=True, color="FFFFFF", size=11)
_ALT_ROW_FILL  = PatternFill("solid", fgColor="D9E1F2")

def save_cleaned(df: pd.DataFrame) -> None:
    """Write the cleaned DataFrame to Excel with professional formatting."""
    CLEANED_FILE.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(CLEANED_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Cleaned_Data")
        ws = writer.sheets["Cleaned_Data"]

        # Column widths
        for col_cells in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col_cells) + 4
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len, 28)

        # Header style
        for cell in ws[1]:
            cell.font      = _HEADER_FONT
            cell.fill      = _HEADER_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Alternating row colour
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            if row_idx % 2 == 0:
                for cell in row:
                    cell.fill = _ALT_ROW_FILL

        # Auto-filter & freeze header
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes    = "A2"

    print(f"\n[G] Cleaned dataset saved → {CLEANED_FILE}")


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN PIPELINE
# ══════════════════════════════════════════════════════════════════════════════

def clean_data(raw_path: Path = RAW_FILE) -> CleaningResult:
    print("=" * 60)
    print("  Data Cleaning Pipeline")
    print("=" * 60)

    res = CleaningResult()
    df  = load_dataset(raw_path)

    res.original_rows = len(df)
    res.original_cols = len(df.columns)
    res.missing_before = int(df.isna().sum().sum())

    df = handle_missing_values(df, res)
    df = remove_duplicates(df, res)
    df = standardize_text(df, res)
    df = validate_numerics(df, res)
    df = validate_emails(df, res)
    df = recalculate_grades(df, res)

    res.missing_after = int(df.isna().sum().sum())
    res.final_rows    = len(df)
    res.df_cleaned    = df.reset_index(drop=True)

    save_cleaned(res.df_cleaned)
    res.print_summary()

    print("=" * 60)
    return res


if __name__ == "__main__":
    clean_data()
