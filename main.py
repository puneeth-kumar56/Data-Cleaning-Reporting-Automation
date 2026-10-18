"""
main.py
--------
Single entry point for the Data Cleaning & Reporting Automation project.

Run this file to execute every stage automatically:
  1. Generate raw dataset
  2. Clean and validate the data
  3. Analyse the cleaned data
  4. Create charts and Excel reports

Usage:
    python main.py
"""

import io
import sys
import time
from pathlib import Path

# Force UTF-8 output so Unicode symbols render on Windows terminals
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
else:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

# ── Make src/ importable when running from the project root ───────────────────
ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from generate_dataset import generate_dataset, save_dataset, OUTPUT_FILE as RAW_FILE
from data_cleaning    import clean_data
from data_analysis    import analyse
from generate_report  import generate_reports


# ── Verification helpers ──────────────────────────────────────────────────────

def verify_outputs(clean_res, ana, chart_paths: dict) -> None:
    """Run basic sanity checks and print a pass/fail table."""
    print("\n" + "=" * 60)
    print("  VERIFICATION CHECKS")
    print("=" * 60)

    cleaned_path = ROOT / "data"  / "student_performance_cleaned.xlsx"
    report_path  = ROOT / "reports" / "data_cleaning_report.xlsx"
    summary_path = ROOT / "reports" / "data_quality_summary.txt"

    checks = [
        ("Raw dataset has ≥ 100 records",
         clean_res.original_rows >= 100),
        ("Duplicates were removed",
         clean_res.duplicates_removed > 0),
        ("Missing values reduced",
         clean_res.missing_after < clean_res.missing_before),
        ("Department names standardised",
         ana.df["Department"].str.upper().eq(ana.df["Department"]).all()),
        ("Cleaned dataset saved",
         cleaned_path.exists()),
        ("Excel report saved",
         report_path.exists()),
        ("Text summary saved",
         summary_path.exists()),
        ("Charts generated",
         len(chart_paths) == 6),
        ("Cleaned dataset non-empty",
         clean_res.final_rows > 0),
        ("Grade column present",
         "Grade" in ana.df.columns),
    ]

    all_passed = True
    for label, result in checks:
        status = "✔  PASS" if result else "✘  FAIL"
        if not result:
            all_passed = False
        print(f"  {status}  {label}")

    print("─" * 60)
    print("  All checks PASSED ✔" if all_passed else "  Some checks FAILED – review output above.")
    print("=" * 60)


# ── Main pipeline ──────────────────────────────────────────────────────────────

def main() -> None:
    t_start = time.time()

    print("\n" + "=" * 60)
    print("  DATA CLEANING & REPORTING AUTOMATION")
    print("  Student Academic Performance")
    print("=" * 60 + "\n")

    # ── Step 1: Generate dataset (skip if already exists) ──────────────────────
    print("STEP 1/4 ─ Dataset Generation")
    print("-" * 40)
    if RAW_FILE.exists():
        print(f"  Raw dataset already exists → {RAW_FILE}")
        print("  Delete it to regenerate, or continue with existing file.")
        import pandas as pd
        raw_df = pd.read_excel(RAW_FILE, engine="openpyxl")
    else:
        raw_df = generate_dataset()
        save_dataset(raw_df)

    # ── Step 2: Clean and validate ─────────────────────────────────────────────
    print("\nSTEP 2/4 ─ Data Cleaning & Validation")
    print("-" * 40)
    clean_res = clean_data()

    # ── Step 3: Analyse ────────────────────────────────────────────────────────
    print("\nSTEP 3/4 ─ Data Analysis")
    print("-" * 40)
    ana = analyse(clean_res.df_cleaned)

    # ── Step 4: Generate reports and charts ────────────────────────────────────
    print("\nSTEP 4/4 ─ Report & Chart Generation")
    print("-" * 40)
    chart_paths, summary_text = generate_reports(clean_res, ana)

    # ── Verification ──────────────────────────────────────────────────────────
    verify_outputs(clean_res, ana, chart_paths)

    # ── Final summary ─────────────────────────────────────────────────────────
    elapsed = round(time.time() - t_start, 1)
    print(f"\n  ✔  Pipeline completed in {elapsed}s")
    print("\n  OUTPUT FILES")
    outputs = [
        ROOT / "data"    / "student_performance_raw.xlsx",
        ROOT / "data"    / "student_performance_cleaned.xlsx",
        ROOT / "reports" / "data_cleaning_report.xlsx",
        ROOT / "reports" / "data_quality_summary.txt",
    ]
    for p in outputs:
        status = "✔" if p.exists() else "✘ MISSING"
        print(f"  {status}  {p.relative_to(ROOT)}")
    charts_dir = ROOT / "reports" / "charts"
    print(f"  ✔  {len(list(charts_dir.glob('*.png')))} charts in reports/charts/")
    print()


if __name__ == "__main__":
    main()
