"""
Mock Data Generator for CFO Dashboard
Creates sample Excel files for testing the Finance Dashboard
"""

import pandas as pd
from pathlib import Path

DATA_FOLDER = Path("data")
DATA_FOLDER.mkdir(exist_ok=True)

print("Generating mock data files for CFO Dashboard...")

# 1. Initial Cash Flow
print("  Creating initial_cash_flow.xlsx...")
initial_cf = pd.DataFrame([
    ["Initial Cash Flow", "", "", ""],
    ["", "Amount", "", ""],
    ["Starting Cash", 450000, "", ""],
    ["Operating Receipts", 850000, "", ""],
    ["Operating Payments", -620000, "", ""],
    ["Tax Payments", -75000, "", ""],
    ["Asset Purchases", -50000, "", ""],
    ["Final Cash", 555000, "", ""],
])
initial_cf.to_excel(DATA_FOLDER / "initial_cash_flow.xlsx", index=False, header=False)

# 2. Results and Balance Statements
print("  Creating results_and_balance_statements.xlsx...")
balance = pd.DataFrame([
    ["Results and Balance Statements", "", ""],
    ["", "", ""],
    ["INCOME STATEMENT", "", ""],
    ["Net Sales", 2500000, ""],
    ["Cost of Goods Sold", -1500000, ""],
    ["Gross Income", 1000000, ""],
    ["S&A Expenses", -350000, ""],
    ["Depreciation", -120000, ""],
    ["Operating Income", 530000, ""],
    ["Interest Expense", -45000, ""],
    ["Net Profit Before Tax", 485000, ""],
    ["Taxes", -145500, ""],
    ["Net Profit", 339500, ""],
    ["", "", ""],
    ["BALANCE SHEET", "", ""],
    ["Total Assets", 4200000, ""],
    ["Total Liabilities", 1680000, ""],
    ["Total Equity", 2520000, ""],
    ["Retained Earnings", 850000, ""],
])
balance.to_excel(DATA_FOLDER / "results_and_balance_statements.xlsx", index=False, header=False)

# 3. Sales Admin Expenses
print("  Creating sales_admin_expenses.xlsx...")
sa_expenses = pd.DataFrame([
    ["Sales & Admin Expenses", "", "", "", "", "", "", "", "", "Total"],
    ["", "FN1", "FN2", "FN3", "FN4", "FN5", "FN6", "FN7", "FN8", ""],
    ["Salaries", 25000, 25000, 25000, 25000, 25000, 25000, 25000, 25000, 200000],
    ["Marketing", 8000, 10000, 8000, 12000, 8000, 10000, 8000, 12000, 76000],
    ["Admin", 5000, 5000, 5000, 5000, 5000, 5000, 5000, 5000, 40000],
    ["Other", 4000, 4500, 4000, 5000, 4000, 4500, 4000, 5000, 35000],
    ["Total Sales & Admin Expenses", 42000, 44500, 42000, 47000, 42000, 44500, 42000, 47000, 351000],
])
sa_expenses.to_excel(DATA_FOLDER / "sales_admin_expenses.xlsx", index=False, header=False)

# 4. Accounts Receivable/Payable
print("  Creating accounts_receivable_payable.xlsx...")
ar_ap = pd.DataFrame([
    ["Accounts Receivable & Payable", "", "", "", "", "", "", "", ""],
    ["", "FN1", "FN2", "FN3", "FN4", "FN5", "FN6", "FN7", "FN8"],
    ["Receivables Due", 85000, 92000, 78000, 110000, 95000, 88000, 102000, 75000],
    ["Payables Due", 45000, 52000, 48000, 65000, 55000, 50000, 58000, 42000],
])
ar_ap.to_excel(DATA_FOLDER / "accounts_receivable_payable.xlsx", index=False, header=False)

# 5. Subperiod Cash Flow
print("  Creating subperiod_cash_flow.xlsx...")
subperiod = pd.DataFrame([
    ["Subperiod Cash Flow", "", "", "", "", "", "", "", ""],
    ["", "FN1", "FN2", "FN3", "FN4", "FN5", "FN6", "FN7", "FN8"],
    ["Opening Cash", 555000, 520000, 485000, 510000, 475000, 495000, 460000, 490000],
    ["Sales Receipts", 280000, 310000, 295000, 340000, 320000, 290000, 335000, 300000],
    ["Procurement", -180000, -195000, -170000, -220000, -185000, -175000, -200000, -165000],
    ["Overhead", -42000, -44500, -42000, -47000, -42000, -44500, -42000, -47000],
    ["Other", -93000, -105500, -58000, -108000, -73000, -105500, -63000, -88000],
    ["Ending Cash", 520000, 485000, 510000, 475000, 495000, 460000, 490000, 490000],
])
subperiod.to_excel(DATA_FOLDER / "subperiod_cash_flow.xlsx", index=False, header=False)

# 6. Production (for COGS reference)
print("  Creating production.xlsx...")
production = pd.DataFrame([
    ["Production Report", "", "", "", "", "", "", "", ""],
    ["", "FN1", "FN2", "FN3", "FN4", "FN5", "FN6", "FN7", "FN8"],
    ["Units Produced", 1200, 1300, 1250, 1400, 1350, 1250, 1400, 1300],
    ["Value of Produced Units", 180000, 195000, 187500, 210000, 202500, 187500, 210000, 195000],
    ["Unit Cost", 150, 150, 150, 150, 150, 150, 150, 150],
])
production.to_excel(DATA_FOLDER / "production.xlsx", index=False, header=False)

# 7. Finance Decisions Template
print("  Creating Finance Decisions.xlsx...")
with pd.ExcelWriter(DATA_FOLDER / "Finance Decisions.xlsx", engine='openpyxl') as writer:
    finance = pd.DataFrame([
        ["Finance Decisions", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", ""],
        ["Credit Lines", "FN1", "FN2", "FN3", "FN4", "FN5", "FN6", "FN7", "FN8"],
        ["Amount", 0, 0, 0, 0, 0, 0, 0, 0],
        ["", "", "", "", "", "", "", "", ""],
        ["Investments", "FN1", "FN2", "FN3", "FN4", "FN5", "FN6", "FN7", "FN8"],
        ["Amount", 0, 0, 0, 0, 0, 0, 0, 0],
        ["", "", "", "", "", "", "", "", ""],
        ["Mortgages", "Amount", "Rate", "Payment1", "Payment2", "", "", "", ""],
        ["Loan 1", 0, 0.08, 0, 0, "", "", "", ""],
        ["Loan 2", 0, 0.08, 0, 0, "", "", "", ""],
        ["", "", "", "", "", "", "", "", ""],
        ["Dividends", 0, "", "", "", "", "", "", ""],
    ])
    finance.to_excel(writer, sheet_name='Finance', index=False, header=False)

print("\n[SUCCESS] All mock data files created in 'data/' folder:")
print("  - initial_cash_flow.xlsx")
print("  - results_and_balance_statements.xlsx")
print("  - sales_admin_expenses.xlsx")
print("  - accounts_receivable_payable.xlsx")
print("  - subperiod_cash_flow.xlsx")
print("  - production.xlsx")
print("  - Finance Decisions.xlsx")
print("\nNow run: python generate_finance_dashboard_final.py")