# ExSim ESG Dashboard: CO2 Abatement Strategy Guide

## Overview

This dashboard helps the Sustainability Officer compare green investment options and determine the most cost-effective way to reduce CO2 emissions and tax burden.

## Decision Framework

### Rule 1: IF PAYBACK < 3 YEARS → BUY SOLAR

Solar PV panels are a CAPITAL investment with long-term returns.

**Calculation:**
> Payback Period = Investment / Annual Tax Savings
> Annual Tax Savings = Panels × CO2_Reduction_Per_Panel × Tax_Rate

**Example:**

- 10 Solar Panels × $15,000 = $150,000 investment
- CO2 Reduction: 10 × 0.5 tons = 5 tons/year
- Tax Savings: 5 tons × $30/ton = $150/year
- **Payback:** $150,000 / $150 = 1,000 years (BAD!)

> [!NOTE]
> **Key Insight:** Solar only makes sense when:
>
> 1. CO2 Tax Rate is HIGH (>$50/ton).
> 2. Reduction per panel is HIGH.
> 3. Panel cost is LOW.
>
> **Check:** `STRATEGY_SELECTOR!G13`. If < 3 years, invest.

### Rule 2: IF SHORT-TERM CASH IS LOW → BUY CREDITS

CO2 Credits are an OPERATING expense - pay as you go.

**Why Credits?**

- No upfront CAPEX required.
- Immediate CO2 offset (1 credit = 1 ton).
- Flexibility: Buy only what you need.

**When to Use:**

- Cash-strapped situations.
- Temporary spikes in emissions.
- Meet compliance quickly.

**Cost Comparison:**
> Credit Cost = $25/ton (typical)
> Tax Cost = $30/ton (if you don't abate)
>
> If Credit Cost < Tax Cost → Buy credits to avoid higher tax.

> [!WARNING]
> Credits provide **NO** long-term benefit. They're a quick fix.

### Rule 3: TREES FOR PR & LONG-TERM

Trees are LOW COST but SLOW to impact CO2.

**Typical Values:**

- Cost: $50/tree
- CO2 Reduction: 0.02 tons/tree/year

**Trees are great for:**

- Corporate image / ESG reporting.
- Long-term sustainability goals.
- Community engagement.

**NOT great for:**

- Meeting immediate emission targets.
- Quick ROI.

### Rule 4: GREEN ELECTRICITY FOR OPERATING BUDGET

Switching to green electricity is an OPERATING expense.

**Input:** Percentage of consumption to convert (0-100%).

**Calculation:**
> Cost = Energy_Consumption × % × Price_Premium
> CO2_Reduced = Energy_Consumption × % × Reduction_Rate

**Example:**

- Energy: 500,000 kWh, Convert: 50%, Premium: $0.03/kWh
- Cost: $7,500/year
- CO2 Reduced: 125 tons

**Good option when:**

- Need predictable annual costs.
- Can't afford upfront solar investment.
- Want immediate, measurable impact.

## How to Use the Dashboard

### Step 1: IMPACT_CONFIG

- Enter CO2 Tax Rate from Case Guide.
- Verify initiative cost/reduction rates.

### Step 2: STRATEGY_SELECTOR

- Enter current emissions in B6.
- Enter energy consumption in B8.
- Adjust quantities in yellow cells.
- Review payback periods and cost per ton.
- Check "Cheapest $/Ton Option" indicator.

### Step 3: INTERPRET RESULTS

- Best option highlighted in **GREEN**.
- Compare payback periods for CAPEX.
- Ensure remaining emissions acceptable.

### Step 4: UPLOAD_READY_ESG

- Copy final quantities to ExSim.

## Quick Reference

| Initiative | Type | Best For |
| :--- | :--- | :--- |
| **Solar PV** | CAPEX | Long-term, high tax rates |
| **Trees** | CAPEX | PR, low budget, long-term |
| **Green Electricity** | OpEx | Predictable annual cost |
| **CO2 Credits** | OpEx | Quick fix, low cash |

**Key Metrics:**

- **Payback Period:** CAPEX / Annual Savings (years)
- **Cost per Ton:** Total Cost / Tons Reduced
- **Net Benefit:** Tax Savings - Annual Cost (for OpEx)

> [!TIP]
> **GOAL:** Minimize cost per ton while meeting emission targets.

## Common Pitfalls

1. **Ignoring Tax Rate Changes**
   - If tax rates increase, solar becomes more attractive.
   - Re-run analysis each period.

2. **Over-investing in Credits**
   - Credits don't build long-term value.
   - Balance with solar/trees for sustainability.

3. **Forgetting Energy Consumption**
   - Green electricity sizing depends on kWh.
   - Update B8 each period.

4. **Not Checking Payback**
   - Column G shows years to recover investment.
   - Anything > 5 years is risky.
