"""Generate synthetic budget and actuals data in a Business Central-like layout.

All data is fictional. The fiscal year runs November -> October.

Output:
    data/budget.csv   - monthly budget by G/L account and cost center
    data/actuals.csv  - posted G/L entries (several per month)
"""
from __future__ import annotations

import random
from datetime import date
from pathlib import Path

import pandas as pd

SEED = 42
FISCAL_YEAR_START = date(2025, 11, 1)  # FY 2026 = Nov 2025 - Oct 2026
DATA_DIR = Path(__file__).parent / "data"

# (G/L account no., name, type, monthly base amount)
ACCOUNTS = [
    ("40100", "Product sales", "Revenue", 180_000),
    ("40200", "Service revenue", "Revenue", 35_000),
    ("50100", "Raw materials", "Cost", 72_000),
    ("50200", "Subcontracting", "Cost", 18_000),
    ("60100", "Personnel costs", "Cost", 48_000),
    ("60200", "Travel expenses", "Cost", 6_000),
    ("60300", "Marketing", "Cost", 9_000),
    ("60400", "IT & software", "Cost", 5_500),
]

COST_CENTERS = {
    "SALES": 0.45,
    "PROD": 0.40,
    "ADMIN": 0.15,
}

# Seasonality index by fiscal month (Nov=1 ... Oct=12)
SEASONALITY = [0.85, 0.80, 0.90, 1.00, 1.10, 1.15, 1.10, 1.05, 1.00, 0.85, 1.05, 1.15]


def month_start(offset: int) -> date:
    y = FISCAL_YEAR_START.year + (FISCAL_YEAR_START.month - 1 + offset) // 12
    m = (FISCAL_YEAR_START.month - 1 + offset) % 12 + 1
    return date(y, m, 1)


def build_budget() -> pd.DataFrame:
    rows = []
    for fm in range(12):
        period = month_start(fm)
        for acc_no, acc_name, acc_type, base in ACCOUNTS:
            # Revenue sits entirely on the SALES cost center; costs are split
            centers = {"SALES": 1.0} if acc_type == "Revenue" else COST_CENTERS
            for cc, share in centers.items():
                amount = round(base * share * SEASONALITY[fm], 2)
                rows.append(
                    {
                        "Period": period.isoformat(),
                        "G/L Account No.": acc_no,
                        "G/L Account Name": acc_name,
                        "Account Type": acc_type,
                        "Cost Center Code": cc,
                        "Budget Amount": amount,
                    }
                )
    return pd.DataFrame(rows)


def build_actuals(budget: pd.DataFrame, rng: random.Random) -> pd.DataFrame:
    rows = []
    entry_no = 100_000
    for _, b in budget.iterrows():
        period = date.fromisoformat(b["Period"])
        # Actual deviates from budget by -15% .. +20%, a few outliers more
        drift = rng.uniform(-0.15, 0.20)
        if rng.random() < 0.05:
            drift += rng.choice([-0.35, 0.40])
        total = b["Budget Amount"] * (1 + drift)
        n_entries = rng.randint(2, 5)
        weights = [rng.random() for _ in range(n_entries)]
        for w in weights:
            entry_no += 1
            amount = round(total * w / sum(weights), 2)
            rows.append(
                {
                    "Entry No.": entry_no,
                    "Posting Date": date(period.year, period.month, rng.randint(1, 28)).isoformat(),
                    "Document No.": f"GJ{entry_no}",
                    "G/L Account No.": b["G/L Account No."],
                    "Description": b["G/L Account Name"],
                    "Cost Center Code": b["Cost Center Code"],
                    # BC convention: revenue is posted as credit (negative)
                    "Amount": -amount if b["Account Type"] == "Revenue" else amount,
                }
            )
    return pd.DataFrame(rows).sort_values("Posting Date").reset_index(drop=True)


def main() -> None:
    rng = random.Random(SEED)
    DATA_DIR.mkdir(exist_ok=True)
    budget = build_budget()
    actuals = build_actuals(budget, rng)
    budget.to_csv(DATA_DIR / "budget.csv", index=False)
    actuals.to_csv(DATA_DIR / "actuals.csv", index=False)
    print(f"Budget rows: {len(budget)} | G/L entries: {len(actuals)} -> {DATA_DIR}")


if __name__ == "__main__":
    main()
