# ExSim CPO Dashboard: Workforce Planning & Compensation

## Overview

This dashboard helps the Chief People Officer (CPO) manage:

- Workforce headcount planning by zone.
- Salary decisions to avoid strikes.
- Benefits policies.
- Total labor cost forecasting for Finance.

## How to Avoid Strikes (Inflation Logic)

ExSim workers demand salary increases that **AT MINIMUM** keep pace with inflation. Failure to meet this expectation triggers **STRIKE RISK**.

> [!IMPORTANT]
> **The Formula:**
> Min Salary to Avoid Strike = Previous Salary × (1 + Inflation Rate)

**Example:**

- Previous Salary: $750/fortnight
- Inflation Rate: 3% (from Case Guide)
- Minimum Safe Salary: $750 × 1.03 = $772.50

### Critical Step

1. Get the **Inflation Rate** from your Case Guide (usually 2-5%).
2. Enter it in `COMPENSATION_STRATEGY` → Cell B6.
3. Set Proposed Salaries **ABOVE** the "Min Salary" column.
4. Watch for **"STRIKE RISK!"** flags - if you see any, **INCREASE** that salary.

**Purchasing Power Principle:**

- Workers want their REAL purchasing power maintained.
- Salary increase = Inflation = 0% real change (neutral).
- Salary increase > Inflation = Positive morale boost.
- Salary increase < Inflation = STRIKE RISK.

> [!TIP]
> Aim for 1-2% **ABOVE** inflation for motivation. The extra cost prevents production disruptions from strikes, which are far more expensive.

## The Hidden Cost of Hiring (Turnover)

New hires are EXPENSIVE. The dashboard captures these hidden costs:

**Direct Costs:**

- **Hiring Fee:** ~$3,000 per new worker (recruiting, onboarding).
- **Severance:** ~$5,000 per fired worker (legal, payouts).

**Indirect Costs (not modeled but real):**

- **Training time:** New workers are less productive initially.
- **Quality issues:** Inexperience leads to defects.
- **Team disruption:** Learning curves affect everyone.

**Turnover Reality:**

- Even stable companies lose 5-10% of workers annually.
- **Formula:** Projected Loss = Current Staff × Turnover Rate.
- You must REPLACE these workers PLUS add any new capacity.

> [!NOTE]
> **Example Scenario:**
>
> - Center Zone: 219 workers, Turnover Rate: 5%.
> - Projected Loss: 219 × 0.05 = 11 workers will leave.
> - If you need 220 workers next period:
>   - Net Staff after turnover: 219 - 11 = 208
>   - Hiring Needed: 220 - 208 = 12 workers
>   - Hiring Cost: 12 × $3,000 = $36,000

**Strategy Tips:**

1. Minimize turnover with good benefits (Training, Health Insurance).
2. Plan ahead - hiring costs hit your cash flow.
3. Avoid over-hiring - firing is MORE expensive than hiring.
4. Use realistic turnover estimates (5% is common, 10% if morale is low).

## Benefits and Motivation

**Training Budget (% of Payroll):**

- **CRITICAL** for product quality.
- Low training → More defective products.
- **Recommended:** 2-5% of payroll.

**Health Insurance (% of Payroll):**

- Reduces absenteeism.
- Low health benefits → Higher sick days.
- **Recommended:** 3-5% of payroll.

**Profit Sharing (% of Net Profit):**

- Boosts motivation and retention.
- Aligns worker interests with company.
- **Typical:** 5-10% of net profit.

**Personal Days & Union Reps:**

- Required by labor agreements.
- Check Case Guide for minimum requirements.

## Using the Dashboard

### Step 1: WORKFORCE_PLANNING

- Review current headcount per zone.
- Enter **Required Workers** (from Production Manager).
- Set realistic **Turnover Rate** (5% default).
- Review Hiring/Firing costs.

### Step 2: COMPENSATION_STRATEGY

- **FIRST: Enter Inflation Rate from Case Guide!**
- Set Proposed Salaries above minimum.
- Check for Strike Risk flags.
- Configure Benefits.

### Step 3: LABOR_COST_ANALYSIS

- Enter Estimated Net Profit (for Profit Sharing calc).
- Review **Total People Expense**.
- Share this number with CFO.

### Step 4: UPLOAD_READY_PEOPLE

- Copy values to ExSim People upload.
- Verify all zones have salaries set.

## Files Required

**Input files (in `data/` folder):**

- `workers_balance_overtime.xlsx` (current headcount)
- `sales_admin_expenses.xlsx` (salesforce data)
- `production.xlsx` (previous labor costs)

**Output file:**

- `CPO_Dashboard.xlsx`

## Quick Reference

| Metric | Rule of Thumb |
| :--- | :--- |
| **Avoid Strikes** | New Salary ≥ Old Salary × (1 + Inflation) |
| **Turnover Impact** | Plan for 5-10% annual worker loss |
| **Hiring Cost** | ~$3,000 per new hire |
| **Firing Cost** | ~$5,000 per termination |
| **Training Budget** | 2-5% of payroll prevents defects |
| **Health Insurance** | 3-5% of payroll reduces absenteeism |
