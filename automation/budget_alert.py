"""
BudgetWatch — Budget Alert Script
Reads budget vs. actuals data, identifies categories more than 10% over budget,
and saves a formatted alert report to a text file.

Usage:
  python budget_alert.py

Output: budget_alert_output.txt in the same folder as this script.
In a production environment this report would be emailed to department leads
via SMTP or Power Automate on a scheduled basis.
"""

import pandas as pd
from datetime import date
from pathlib import Path

# --- Configuration ---
DATA_FILE = "C:/Users/chloe/OneDrive/Desktop/Portfolio/data/contoso_budget_clean.xlsx"
OUTPUT_FILE = Path(__file__).parent / "budget_alert_output.txt"
ALERT_THRESHOLD = -0.10  # flag anything more than 10% over budget


def load_data(filepath):
    return pd.read_excel(filepath, sheet_name="Budget_Data")


def get_alerts(df):
    alerts = df[df["Variance_Pct"] < ALERT_THRESHOLD].copy()
    alerts = alerts.sort_values("Variance_Pct")  # worst first
    return alerts


def format_report(alerts):
    count = len(alerts)
    threshold_pct = int(abs(ALERT_THRESHOLD) * 100)

    lines = [
        f"BudgetWatch Alert — {count} items are over {threshold_pct}% budget threshold",
        f"Report date: {date.today().strftime('%B %d, %Y')}",
        "",
        f"{'Department':<22} {'Category':<28} {'Month':<12} {'Budget':>10} {'Actual':>10} {'Variance%':>10}",
        "-" * 95,
    ]

    for _, row in alerts.iterrows():
        lines.append(
            f"{row['Department']:<22} {row['Category']:<28} {row['Month']:<12}"
            f" ${row['Budget_Amount']:>9,.0f} ${row['Actual_Amount']:>9,.0f}"
            f" {row['Variance_Pct'] * 100:>+9.1f}%"
        )

    lines += [
        "",
        f"Source: {Path(DATA_FILE).name}",
        "Action required: Review flagged items with department leads.",
    ]

    return "\n".join(lines)


def save_report(report, output_path):
    with open(output_path, "w") as f:
        f.write(report)


if __name__ == "__main__":
    df = load_data(DATA_FILE)
    alerts = get_alerts(df)

    print(f"Alerts found: {len(alerts)} items over {int(abs(ALERT_THRESHOLD) * 100)}% budget threshold")

    if alerts.empty:
        print("No alerts to report.")
    else:
        report = format_report(alerts)
        save_report(report, OUTPUT_FILE)

        print(report)
        print(f"\nReport saved to: {OUTPUT_FILE}")
