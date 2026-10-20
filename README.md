# Budget vs Actual – automated variance report

Turns a monthly budget and the posted G/L entries from an ERP (Microsoft Dynamics 365 Business Central layout) into a formatted Excel variance report, in seconds instead of an afternoon of copy-paste.

> All data in this repository is **synthetic**. The fiscal year runs November → October.

## What it does

1. Reads the **budget** (period × G/L account × cost center) and the **G/L entries** export.
2. Normalises signs (Business Central posts revenue as negative amounts).
3. Computes **month** and **year-to-date** variances, absolute and %.
4. Classifies each variance as **favourable / unfavourable** (more revenue or less cost than budget = favourable).
5. Writes an Excel report with:
   - **Summary** – operating result (month and YTD) plus the account-level detail
   - **By cost center** – YTD variances by cost center and account
   - **Monthly trend** – revenue and costs, month by month
   - **Top deviations** – the 10 largest gaps to look at first
   
   Rows above the threshold are highlighted in green (favourable) or red (unfavourable).

## Quick start

```bash
pip install -r requirements.txt
python generate_sample_data.py          # creates data/budget.csv and data/actuals.csv
python variance_report.py               # latest month
python variance_report.py --month 2026-06 --threshold 0.15
```

The report is saved in `output/variance_report_<month>.xlsx`.

## Runs by itself

The GitHub Action in `.github/workflows/monthly-report.yml` rebuilds the report on the **5th of every month** (after month-end close) and attaches the Excel file to the run. It can also be started manually from the **Actions** tab, choosing the month.

## Using it with real data

Replace the two CSV files in `data/` with your own exports:

| File | Required columns |
|---|---|
| `budget.csv` | `Period`, `G/L Account No.`, `G/L Account Name`, `Account Type` (Revenue/Cost), `Cost Center Code`, `Budget Amount` |
| `actuals.csv` | `Posting Date`, `G/L Account No.`, `Cost Center Code`, `Amount` |

In Business Central these come from the *G/L Budget Entries* and *General Ledger Entries* pages (export to Excel/CSV), or via the standard API / an OData web service for a fully automated feed.

## Possible extensions

- Read directly from the Business Central API instead of CSV files
- Email the report to cost-center owners
- Add a short AI-written commentary on the main variances
- Publish the same tables to a Power BI dataset

---

🇮🇹 **In breve:** script Python che confronta budget e consuntivo (movimenti contabili in formato Business Central), calcola gli scostamenti del mese e progressivi per conto e centro di costo e produce un report Excel formattato. Con GitHub Actions il report si rigenera da solo ogni mese. Dati di esempio fittizi.

**Author:** [Daniele Marsico](https://github.com/daniele-marsico) · [ControllingHub](https://controllinghub.netlify.app)
