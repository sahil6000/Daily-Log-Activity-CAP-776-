# My Data, My Story — Personal Activity Intelligence Report

> **CAP776: Programming in Python — Minor Project #1**

A Python-based personal activity analysis project that reads daily activity records from an Excel workbook, validates the data, calculates activity indices, analyzes relationships between activities and experience, and presents the results in a structured report.

---

## 📌 Project Overview

**My Data, My Story** uses my recorded daily activities to understand how I spend my time across sleep, academics, coding, fitness, other activities, and free/unaccounted time.

The project covers the prescribed activity-recording period:

**13 August 2026 → 21 September 2026**

The analysis is performed only on valid records within this project period.

---

## 🎯 Project Objectives

- Read and process activity data from an Excel workbook.
- Validate the recorded data and check data continuity.
- Calculate the required Personal Activity Intelligence indices.
- Convert Feeling, Satisfaction, and Energy text values into their prescribed scores.
- Analyze relationships between selected activities and experience measures.
- Generate evidence-based findings about my daily routine.
- Document the results in a final Personal Activity Intelligence Report.

---

## 🛠️ Technologies Used

- **Python 3**
- **openpyxl** — Excel file handling
- **datetime** — date and project-period handling
- Python fundamentals:
  - Variables
  - Conditions
  - Loops
  - Lists
  - Dictionaries
  - Functions
  - Exception handling

> **Note:** The project does not use Pandas or NumPy, following the project requirements.

---

## 📁 Project Files

| File | Description |
|---|---|
| `12604444.py` | Main Python program for validation, calculations, indices, and relationship analysis |
| `12604444.xlsx` | Daily activity data and project instructions |
| `CAP776-MinorProject1Report.docx` | Completed Minor Project #1 report |

---

## 📊 Activity Data

The Excel workbook contains daily records including:

- Sleep
- Fitness
- Study
- Coding
- Class
- Classes Attended
- Other Activities
- Total Tracked Time
- Free / Unaccounted Time
- Day's Feeling
- Satisfaction Level
- Energy Level
- Notes

### Experience Scales

| Category | Score Scale |
|---|---|
| Day's Feeling | Excellent = 5, Good = 4, Neutral = 3, Low = 2, Stressed = 1 |
| Satisfaction Level | Very Satisfied = 5, Satisfied = 4, Neutral = 3, Unsatisfied = 2, Very Unsatisfied = 1 |
| Energy Level | High = 3, Medium = 2, Low = 1 |

---

## 🔎 Data Validation & Continuity

For the project period:

| Check | Result |
|---|---:|
| Expected days | **40** |
| Records found | **40** |
| Valid records | **40** |
| Invalid records | **0** |
| Missing dates | **0** |
| Duplicate dates | **0** |
| Data Continuity Index (DCI) | **100.00%** |

---

## 📈 Activity Summary

Average values across valid recorded days:

| Activity | Average per Day |
|---|---:|
| Sleep | **444.75 min** |
| Fitness | **30.12 min** |
| Study | **112.50 min** |
| Coding | **133.50 min** |
| Class | **187.50 min** |
| Other Activities | **114.25 min** |
| Free / Unaccounted Time | **417.38 min** |

---

## 📐 Personal Activity Indices

| Index | Value |
|---|---:|
| **TPI — Tech Productivity** | **133.50 min/day** |
| **AAI — Academic Activity** | **300.00 min/day** |
| **PhAI — Physical Activity** | **30.12 min/day** |
| **SRI — Sleep & Recovery** | **444.75 min/day** |
| **ABI — Activity Balance** | **417.38 min/day** |
| **TUI — Time Utilization** | **1022.62 min/day** |
| **EI — Experience Index** | **3.48 / 5** |
| **DCI — Data Continuity Index** | **100.00%** |
| **PAI — Personal Activity Index** | **332.24** |

---

## 🔗 Relationship Analysis

The project calculates the required relationships using correlation analysis.

| Relationship | Correlation |
|---|---:|
| Sleep ↔ Energy | **0.20** |
| Study ↔ Satisfaction | **0.16** |
| Coding ↔ Energy | **0.00** |
| Feeling ↔ Satisfaction | **0.59** |
| Feeling ↔ Energy | **0.72** |

These values describe relationships present in the recorded dataset; they are not intended to establish causation.

---

## ▶️ How to Run

Make sure Python and `openpyxl` are installed.

From the project directory:

```bash
python 12604444.py
```

The Python program reads:

```text
12604444.xlsx
```

and prints the validation results, activity summary, index values, and relationship analysis in the terminal.

---

## 🧩 Program Structure

The Python program is organized into reusable sections/functions for:

1. Excel file loading
2. Date-range filtering
3. Data validation
4. Data continuity checking
5. Activity summary calculations
6. TPI calculation
7. AAI calculation
8. PhAI calculation
9. SRI calculation
10. ABI calculation
11. TUI calculation
12. EI calculation
13. DCI calculation
14. PAI calculation
15. Correlation analysis

The program is intentionally implemented as a single Python file as required for the project.

---

## 📄 Final Report

The final report contains:

- Student details
- Activity Data Summary
- Personal Activity Index values
- Key Findings from the data
- Findings About Myself
- Areas and Ways to Improve My Lifestyle
- Student Declaration

---

## 👨‍🎓 Student

**Sahil Kumar**  
**Registration / Roll No.: 12604444**  
**Program / Section: CAP776 / D1P2634**

---

## 📚 Academic Project

This repository contains the implementation and final report for:

**CAP776 — Programming in Python**  
**Minor Project #1 — My Data, My Story**
