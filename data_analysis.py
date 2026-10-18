"""
data_analysis.py
-----------------
Loads student_performance_cleaned.xlsx and computes all required statistics.
Returns an AnalysisResult dataclass used by generate_report.py and main.py.
"""

from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd

ROOT        = Path(__file__).resolve().parents[1]
CLEANED_FILE = ROOT / "data" / "student_performance_cleaned.xlsx"


# ── Result container ───────────────────────────────────────────────────────────

@dataclass
class AnalysisResult:
    # Overall stats
    total_students: int = 0
    avg_age: float = 0.0
    avg_attendance: float = 0.0
    avg_assignment: float = 0.0
    avg_internal: float = 0.0
    avg_final: float = 0.0
    avg_total: float = 0.0
    max_total: float = 0.0
    min_total: float = 0.0
    low_attendance_count: int = 0

    # Aggregations (DataFrames)
    dept_counts: pd.DataFrame        = field(default_factory=pd.DataFrame)
    dept_avg_marks: pd.DataFrame     = field(default_factory=pd.DataFrame)
    dept_avg_attend: pd.DataFrame    = field(default_factory=pd.DataFrame)
    dept_max_marks: pd.DataFrame     = field(default_factory=pd.DataFrame)
    grade_dist: pd.DataFrame         = field(default_factory=pd.DataFrame)
    attend_bins: pd.Series           = field(default_factory=pd.Series)
    correlation: float | None        = None

    # Raw cleaned DataFrame (for report sheet)
    df: pd.DataFrame                 = field(default_factory=pd.DataFrame)

    def print_summary(self) -> None:
        print("\n" + "─" * 60)
        print("  DATA ANALYSIS SUMMARY")
        print("─" * 60)
        print(f"  Total students          : {self.total_students}")
        print(f"  Average age             : {self.avg_age:.1f} years")
        print(f"  Average attendance      : {self.avg_attendance:.1f}%")
        print(f"  Average assignment marks: {self.avg_assignment:.1f}")
        print(f"  Average internal marks  : {self.avg_internal:.1f}")
        print(f"  Average final marks     : {self.avg_final:.1f}")
        print(f"  Average total marks     : {self.avg_total:.1f}")
        print(f"  Highest total marks     : {self.max_total:.1f}")
        print(f"  Lowest total marks      : {self.min_total:.1f}")
        print(f"  Students attendance<75% : {self.low_attendance_count}")
        if self.correlation is not None:
            print(f"  Attendance-Marks corr   : {self.correlation:.3f}")
        print("─" * 60)


# ── Loader ─────────────────────────────────────────────────────────────────────

def load_cleaned(path: Path = CLEANED_FILE) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Cleaned dataset not found: {path}")
    df = pd.read_excel(path, engine="openpyxl")
    print(f"[Analysis] Loaded cleaned dataset: {df.shape[0]} rows")
    return df


# ── Analysis functions ─────────────────────────────────────────────────────────

def overall_stats(df: pd.DataFrame, res: AnalysisResult) -> None:
    """Scalar statistics for the whole dataset."""
    res.total_students    = len(df)
    res.avg_age           = round(df["Age"].mean(), 1)
    res.avg_attendance    = round(df["Attendance_Percentage"].mean(), 1)
    res.avg_assignment    = round(df["Assignment_Marks"].mean(), 1)
    res.avg_internal      = round(df["Internal_Marks"].mean(), 1)
    res.avg_final         = round(df["Final_Exam_Marks"].mean(), 1)
    res.avg_total         = round(df["Total_Marks"].mean(), 1)
    res.max_total         = round(df["Total_Marks"].max(), 1)
    res.min_total         = round(df["Total_Marks"].min(), 1)
    res.low_attendance_count = int((df["Attendance_Percentage"] < 75).sum())


def department_analysis(df: pd.DataFrame, res: AnalysisResult) -> None:
    """Department-level aggregations."""
    res.dept_counts = (
        df.groupby("Department")
        .size()
        .reset_index(name="Student_Count")
        .sort_values("Student_Count", ascending=False)
    )

    res.dept_avg_marks = (
        df.groupby("Department")[
            ["Assignment_Marks", "Internal_Marks", "Final_Exam_Marks", "Total_Marks"]
        ]
        .mean()
        .round(2)
        .reset_index()
    )

    res.dept_avg_attend = (
        df.groupby("Department")["Attendance_Percentage"]
        .mean()
        .round(2)
        .reset_index(name="Avg_Attendance")
    )

    res.dept_max_marks = (
        df.groupby("Department")["Total_Marks"]
        .max()
        .round(2)
        .reset_index(name="Max_Total_Marks")
    )


def grade_analysis(df: pd.DataFrame, res: AnalysisResult) -> None:
    """Grade distribution counts and percentages."""
    counts = df["Grade"].value_counts().reset_index()
    counts.columns = ["Grade", "Count"]
    counts["Percentage"] = (counts["Count"] / len(df) * 100).round(2)
    res.grade_dist = counts


def attendance_bins(df: pd.DataFrame, res: AnalysisResult) -> None:
    """Bucket attendance into bands."""
    bins   = [0, 50, 60, 75, 85, 100]
    labels = ["<50%", "50-60%", "60-75%", "75-85%", "85-100%"]
    res.attend_bins = (
        pd.cut(df["Attendance_Percentage"], bins=bins, labels=labels, include_lowest=True)
        .value_counts()
        .sort_index()
    )


def attendance_marks_correlation(df: pd.DataFrame, res: AnalysisResult) -> None:
    """Pearson correlation between Attendance_Percentage and Total_Marks."""
    valid = df[["Attendance_Percentage", "Total_Marks"]].dropna()
    if len(valid) >= 3:
        res.correlation = round(valid["Attendance_Percentage"].corr(valid["Total_Marks"]), 4)
    else:
        res.correlation = None


# ── Main entry point ───────────────────────────────────────────────────────────

def analyse(df: pd.DataFrame | None = None) -> AnalysisResult:
    print("=" * 60)
    print("  Data Analysis")
    print("=" * 60)

    if df is None:
        df = load_cleaned()

    res = AnalysisResult(df=df)

    overall_stats(df, res)
    department_analysis(df, res)
    grade_analysis(df, res)
    attendance_bins(df, res)
    attendance_marks_correlation(df, res)

    res.print_summary()

    # Department details
    print("\n  Department Counts:")
    print(res.dept_counts.to_string(index=False))
    print("\n  Grade Distribution:")
    print(res.grade_dist.to_string(index=False))

    print("=" * 60)
    return res


if __name__ == "__main__":
    analyse()
