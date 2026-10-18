"""
generate_dataset.py
--------------------
Generates a synthetic student academic performance dataset with intentional
data quality issues (missing values, duplicates, inconsistent text, invalid
values, bad emails) and saves it as data/student_performance_raw.xlsx.
"""

import os
import random
import numpy as np
import pandas as pd
from pathlib import Path

# ── Reproducibility ────────────────────────────────────────────────────────────
SEED = 42
random.seed(SEED)
np.random.seed(SEED)

# ── Constants ──────────────────────────────────────────────────────────────────
NUM_RECORDS = 120          # generate extra so duplicates still give ≥ 100 unique
NUM_DUPLICATES = 10        # rows to copy as exact duplicates
OUTPUT_DIR = Path(__file__).resolve().parents[1] / "data"
OUTPUT_FILE = OUTPUT_DIR / "student_performance_raw.xlsx"

DEPARTMENTS_CLEAN = ["CSE", "ECE", "ME", "CE", "EEE"]
# Dirty variants introduced on purpose
DEPARTMENTS_DIRTY = [
    "CSE", "cse", "Cse", "CSE ", " CSE",
    "ECE", "ece", "Ece", " ECE",
    "ME", "me", "Me",
    "CE", "ce", "Ce",
    "EEE", "eee", "Eee",
]

GENDERS_CLEAN = ["Male", "Female"]
GENDERS_DIRTY = [
    "Male", "male", "MALE", "M",
    "Female", "female", "FEMALE", "F",
]

CITIES_CLEAN = [
    "Mumbai", "Delhi", "Bengaluru", "Chennai",
    "Hyderabad", "Pune", "Kolkata", "Ahmedabad",
]
CITIES_DIRTY = [
    "Mumbai", "mumbai", "MUMBAI", "Bombay",
    "Delhi", "delhi", "DELHI", "New Delhi",
    "Bengaluru", "Bangalore", "bangalore", "BENGALURU",
    "Chennai", "chennai", "CHENNAI",
    "Hyderabad", "hyderabad", "HYDERABAD",
    "Pune", "pune", "PUNE",
    "Kolkata", "kolkata", "KOLKATA", "Calcutta",
    "Ahmedabad", "ahmedabad", "AHMEDABAD",
]

FIRST_NAMES = [
    "Aarav", "Aditya", "Akash", "Amit", "Anjali", "Ankit", "Ananya", "Arjun",
    "Ayesha", "Deepak", "Divya", "Ishaan", "Kavya", "Kiran", "Kunal",
    "Lakshmi", "Manish", "Meera", "Mohit", "Nisha", "Pooja", "Priya",
    "Rahul", "Raj", "Ravi", "Rohit", "Sachin", "Sanya", "Shreya", "Sid",
    "Sneha", "Sonia", "Suresh", "Tanya", "Vijay", "Vikas", "Vishal", "Zara",
]
LAST_NAMES = [
    "Sharma", "Verma", "Gupta", "Kumar", "Singh", "Patel", "Shah", "Mehta",
    "Joshi", "Rao", "Nair", "Iyer", "Reddy", "Kapoor", "Malhotra", "Desai",
    "Agarwal", "Bose", "Chatterjee", "Das",
]


# ── Helper functions ──────────────────────────────────────────────────────────

def make_student_name() -> str:
    return f"{random.choice(FIRST_NAMES)} {random.choice(LAST_NAMES)}"


def make_email(name: str, student_id: int) -> str:
    """Generate an email; ~15 % will be intentionally malformed."""
    local = name.lower().replace(" ", ".") + str(student_id)
    domain = random.choice(["gmail.com", "yahoo.com", "outlook.com", "edu.in"])
    valid_email = f"{local}@{domain}"

    roll = random.random()
    if roll < 0.05:
        return local + domain          # missing @
    elif roll < 0.10:
        return f"{local}@@{domain}"   # double @
    elif roll < 0.15:
        return f"{local}@"            # no domain
    else:
        return valid_email


def make_grade(total: float | None) -> str:
    """Map total marks (out of 100) to a letter grade."""
    if total is None or np.isnan(total):
        return random.choice(["A", "B", "C", "D", "F", ""])
    if total >= 90:
        return "A+"
    elif total >= 80:
        return "A"
    elif total >= 70:
        return "B"
    elif total >= 60:
        return "C"
    elif total >= 50:
        return "D"
    else:
        return "F"


def make_marks() -> tuple:
    """Return (assignment, internal, final, total) with occasional anomalies."""
    assignment = round(random.uniform(5, 20), 1)    # max 20
    internal   = round(random.uniform(10, 30), 1)   # max 30
    final_exam = round(random.uniform(20, 50), 1)   # max 50
    total      = round(assignment + internal + final_exam, 1)

    # Introduce a few out-of-range values
    roll = random.random()
    if roll < 0.04:
        assignment = round(random.uniform(21, 30), 1)   # > max 20
    elif roll < 0.06:
        internal = round(random.uniform(-5, 0), 1)      # negative

    return assignment, internal, final_exam, total


# ── Main generator ────────────────────────────────────────────────────────────

def generate_dataset() -> pd.DataFrame:
    """Build and return the raw DataFrame."""
    records = []

    for i in range(1, NUM_RECORDS + 1):
        student_id = f"STU{1000 + i:04d}"
        name = make_student_name()

        # Age: mostly 17-24 but a few invalid values
        if random.random() < 0.04:
            age: int | None = random.choice([-3, 0, 150, 200])
        else:
            age = random.randint(17, 24)

        # Attendance: mostly 40-100 but a few > 100
        if random.random() < 0.04:
            attendance: float | None = round(random.uniform(101, 120), 1)
        else:
            attendance = round(random.uniform(40, 100), 1)

        dept   = random.choice(DEPARTMENTS_DIRTY)
        gender = random.choice(GENDERS_DIRTY)
        city   = random.choice(CITIES_DIRTY)
        sem    = random.randint(1, 8)

        assignment, internal, final_exam, total = make_marks()
        grade = make_grade(total)
        email = make_email(name, 1000 + i)

        records.append({
            "Student_ID":            student_id,
            "Student_Name":          name,
            "Age":                   age,
            "Gender":                gender,
            "Department":            dept,
            "Semester":              sem,
            "Attendance_Percentage": attendance,
            "Assignment_Marks":      assignment,
            "Internal_Marks":        internal,
            "Final_Exam_Marks":      final_exam,
            "Total_Marks":           total,
            "Grade":                 grade,
            "Email":                 email,
            "City":                  city,
        })

    df = pd.DataFrame(records)

    # ── Introduce missing values ───────────────────────────────────────────────
    miss_config = {
        "Age":                   0.07,
        "Gender":                0.06,
        "Attendance_Percentage": 0.08,
        "Assignment_Marks":      0.05,
        "Internal_Marks":        0.05,
        "Final_Exam_Marks":      0.06,
        "Total_Marks":           0.05,
        "Grade":                 0.06,
        "Email":                 0.07,
        "City":                  0.05,
        "Department":            0.04,
    }
    for col, frac in miss_config.items():
        n_miss = max(1, int(len(df) * frac))
        idx = np.random.choice(df.index, size=n_miss, replace=False)
        df.loc[idx, col] = np.nan

    # ── Introduce duplicate rows ───────────────────────────────────────────────
    dup_rows = df.sample(n=NUM_DUPLICATES, random_state=SEED)
    df = pd.concat([df, dup_rows], ignore_index=True)

    # ── Shuffle so duplicates aren't obviously at the end ─────────────────────
    df = df.sample(frac=1, random_state=SEED).reset_index(drop=True)

    return df


def save_dataset(df: pd.DataFrame) -> None:
    """Save the DataFrame to an Excel file."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Raw_Data")

        # Basic formatting so the file opens readably
        ws = writer.sheets["Raw_Data"]
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col) + 4
            ws.column_dimensions[col[0].column_letter].width = min(max_len, 30)

    print(f"✔  Dataset saved  → {OUTPUT_FILE}")
    print(f"   Total records  : {len(df)}")
    print(f"   Columns        : {list(df.columns)}")


# ── Entry point ────────────────────────────────────────────────────────────────

def main() -> pd.DataFrame:
    print("=" * 60)
    print("  Generating Synthetic Student Performance Dataset")
    print("=" * 60)
    df = generate_dataset()
    save_dataset(df)
    print("=" * 60)
    return df


if __name__ == "__main__":
    main()
