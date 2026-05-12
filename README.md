# BudgetWatch — Automated Budget Variance Monitoring System

A portfolio project simulating a real-world budget monitoring workflow for Contoso Manufacturing (fictional). Identifies where and why budget variances are happening across 5 departments, and automates alerts so the right people are notified immediately — not at month-end.

---

## Demo

> 📹 **[Watch the demo video](#)** ← *(add your video link here after recording)*

**Executive Overview — Page 1**
*(add screenshot here)*

**Department Drill-Down — Page 2**
*(add screenshot here)*

**Alerts Summary — Page 3**
*(add screenshot here)*

---

## What It Does

- **Tracks budget vs. actuals** across 5 departments, 7 spending categories, and 12 months (420 rows of data)
- **Visualizes variances** in an interactive Power BI dashboard with 3 pages: executive overview, department drill-down, and alerts summary
- **Flags overspending automatically** — any category more than 10% over budget is surfaced on the Alerts page
- **Generates alert reports** via a Python script that can be scheduled to run on a recurring basis

---

## Tech Stack

| Tool | Purpose |
|------|---------|
| Python (pandas, openpyxl) | Data generation + alert automation |
| Power BI Desktop | Interactive dashboard (3 pages) |
| Excel | Data source (`contoso_budget_clean.xlsx`) |
| DAX | Calculated measures (Total Budget, Total Actuals, Variance $, Variance %) |

---

## Project Structure

```
budgetwatch/
├── README.md
├── data/
│   ├── generate_data.py         # generates both raw and clean Excel files
│   ├── contoso_budget_clean.xlsx  # clean data — Power BI data source
│   └── contoso_budget_raw.xlsx    # messy "before" version for demo context
├── automation/
│   ├── budget_alert.py          # reads Excel, flags overruns, saves report
│   └── budget_alert_output.txt  # sample output from budget_alert.py
└── dashboard/
    └── Budget Watch.pbix        # Power BI file (open in Power BI Desktop)
```

---

## Key Findings

- **20 out of 420 rows** are flagged as >10% over budget
- **Biggest overspend:** IT Software & Licenses in January (+50%) — annual software renewals spike
- **Marketing** Contractor Services consistently over budget in Q3–Q4 (+11.3% annual)
- **Operations** Travel & Entertainment overspends Q3–Q4 (+11.7% annual)
- **HR Training** is significantly under budget in Q1–Q2 (-14.5%) — budget going unused

---

## How to Run

**Generate data:**
```bash
pip install pandas openpyxl
python data/generate_data.py
```

**Run budget alert:**
```bash
python automation/budget_alert.py
```
Output is saved to `automation/budget_alert_output.txt`.

**View dashboard:**
Open `dashboard/Budget Watch.pbix` in Power BI Desktop (free download at microsoft.com/power-bi).

---

## Business Value

> *"The dashboard identified where and why budget variances were happening across all departments. The automated alert script ensures the right people are notified immediately instead of waiting until month-end review — so teams can course-correct while there's still time to act."*
