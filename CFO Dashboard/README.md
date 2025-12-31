# ExSim Finance Dashboard - Audit

## Overview

This Financial Control & Liquidity Dashboard helps Finance Managers verify forecast accuracy against historical Balance Sheet and Income Statement data. It integrates cash flow, profit projection, and debt control.

## Setup

1. **Place data files in the `data/` folder:**
    - `initial_cash_flow.xlsx`
    - `results_and_balance_statements.xlsx`
    - `sales_admin_expenses.xlsx`
    - `accounts_receivable_payable.xlsx`
    - `Finance Decisions.xlsx` (optional template)

2. **Run:**

   ```bash
   python generate_finance_dashboard_final.py
   ```

3. **Open:** `Finance_Dashboard_Final.xlsx`

## The Five Tabs

### Tab 1: LIQUIDITY_MONITOR (Cash Flow Engine)

#### Section A: Initialization

- Bridges last period's cash to this period's starting point.
- Deducts taxes, dividends, and asset purchases.

#### Section B: Operational Cash Flow

- **Sales Receipts:** INPUT from Marketing forecasts.
- **Procurement Spend:** INPUT from Purchasing dashboard.
- **Fixed Overhead:** Pre-filled from S&A expenses.
- **Receivables/Payables:** Hard-coded scheduled amounts.

#### Section C: Financing Decisions

- Credit Line changes (+/-).
- Investment changes (+/-).
- New Mortgage inflows.

#### Section D: Cash Balance

- **Net Cash Flow** = Inflows - Outflows.
- **Ending Cash** = Opening + Net.
- **Solvency Check:** "INSOLVENT!" if < $0.

**Color Codes:**

- 🔴 **RED:** Ending cash < 0 (Bankruptcy risk!)
- 🟢 **GREEN:** Ending cash > $200k (Inefficient - invest!)

### Tab 2: PROFIT_CONTROL (Income Statement Projection)

**How to Calibrate Your Forecast Against Reality:**

1. **Review "Last Round Actuals" column:**
    - These values come from `results_and_balance_statements.xlsx`.
    - They represent what ACTUALLY happened last period.

2. **Enter "This Round Projected" column:**
    - **Revenue:** Your sales forecast.
    - **COGS:** Auto-calculated using historical gross margin %.
    - **S&A:** From your budget.
    - **Depreciation:** Usually flat (same as last period).
    - **Interest:** Based on expected loan balances.

3. **Check "Variance %" column:**
    - Shows % difference from actuals.
    - Large variances (>20%) require justification.

4. **Watch "Profit Realism Flag":**
    - If projected net margin > historical + 5%.
    - Flag shows "WARNING: Unrealistic profit jump!".
    - This means your projection may be too optimistic.

**Historical Margins (Reference):**

- **Gross Margin %:** Shows COGS efficiency.
- **Net Margin %:** Shows overall profitability.

**Color Codes:**

- 🟢 **GREEN:** Positive net income.
- 🔴 **RED:** Negative net income (loss projected).

### Tab 3: BALANCE_SHEET_HEALTH

**Why Debt Ratio Matters for Interest Rates:**
> Current Debt Ratio = Total Liabilities / Total Assets

**Thresholds:**

| Ratio | Status | Interest Rate Impact |
| :--- | :--- | :--- |
| < 40% | Healthy | Best available rates |
| 40-60% | Moderate | Standard rates |
| > 60% | Critical | Premium rates, may refuse |

**Projected New Debt Analysis:**

- Enter planned credit lines and mortgages.
- See "Est. Post-Decision Debt Ratio".
- If > 60%, reconsider expansion.

**Warning Flags:**

- "CRITICAL: Debt too high" = Credit rating risk.
- "CRITICAL: Equity Erosion" = Retained earnings negative.

### Tab 4: DEBT_MANAGER

Mortgage calculator for multi-period loans.

**Inputs:**

- **Loan Amount:** Principal borrowed.
- **Interest Rate:** Annual rate.
- **Payment Period 1/2:** Scheduled repayments.

**Outputs:**

- **Total Payments:** Sum of all payments.
- Links to `UPLOAD_READY` for ExSim entry.

### Tab 5: UPLOAD_READY_FINANCE

ExSim format with all financial decisions.

**Sections:**

- **Credit Lines:** Links to LIQUIDITY_MONITOR.
- **Investments:** Links to LIQUIDITY_MONITOR.
- **Mortgages:** Links to DEBT_MANAGER.
- **Dividends:** Links to LIQUIDITY_MONITOR.

## Reviewing Initial vs Subperiod Cash Flow

The dashboard distinguishes between:

**Initial Cash Flow (Section A in Tab 1):**

- Represents the "waterfall" from P0 end to F1 start.
- Includes: Taxes, Dividends, Large Asset Purchases.
- These are one-time deductions at period start.

**Subperiod Cash Flow (Section B in Tab 1):**

- Represents flows during each fortnight.
- Includes: Sales, Procurement, Overhead, AR/AP.
- These recur each fortnight.

**How to Review Both Simultaneously:**

1. Open **Tab 1: LIQUIDITY_MONITOR**.
2. **Section A** shows your STARTING position.
3. **Section B/C/D** shows DURING-period flows.
4. The "Opening Cash" row links them:
    - FN1 Opening = Section A result.
    - FN2+ Opening = Previous FN Ending.

## Strategic Workflow

1. **START with Tab 2: PROFIT_CONTROL**
    - Enter revenue forecast.
    - Verify COGS/margins are realistic.
    - Check "Variance %" vs last round.

2. **GO TO Tab 3: BALANCE_SHEET_HEALTH**
    - Review current debt ratio.
    - Plan any new debt carefully.
    - Avoid crossing 60% threshold.

3. **THEN Tab 1: LIQUIDITY_MONITOR**
    - Enter Section A deductions.
    - Fill Section B estimates.
    - Adjust Section C financing to avoid insolvency.

4. **FINALIZE Tab 4: DEBT_MANAGER**
    - Enter any new mortgages.
    - Schedule repayments.

5. **VERIFY Tab 5 is complete for upload**

## Key Financial Formulas

- **Gross Margin %** = Gross Income / Net Sales
- **Net Margin %** = Net Profit / Net Sales
- **Debt Ratio** = Total Liabilities / Total Assets
- **Ending Cash** = Opening Cash + Net Cash Flow

---
*Maintain financial control and forecast accuracy!*
