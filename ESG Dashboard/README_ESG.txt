================================================================================
                    ESG DASHBOARD - README
               CO2 Abatement Strategy Guide
================================================================================

OVERVIEW
--------
This dashboard helps the Sustainability Officer compare green investment options
and determine the most cost-effective way to reduce CO2 emissions and tax burden.

================================================================================
                    DECISION FRAMEWORK
================================================================================

RULE 1: IF PAYBACK < 3 YEARS → BUY SOLAR
-----------------------------------------
Solar PV panels are a CAPITAL investment with long-term returns.

CALCULATION:
  Payback Period = Investment / Annual Tax Savings
  Annual Tax Savings = Panels × CO2_Reduction_Per_Panel × Tax_Rate

EXAMPLE:
  - 10 Solar Panels × $15,000 = $150,000 investment
  - CO2 Reduction: 10 × 0.5 tons = 5 tons/year
  - Tax Savings: 5 tons × $30/ton = $150/year
  - Payback: $150,000 / $150 = 1,000 years (BAD!)

The key insight: Solar only makes sense when:
  1. CO2 Tax Rate is HIGH (>$50/ton)
  2. Reduction per panel is HIGH
  3. Panel cost is LOW

Check the payback in STRATEGY_SELECTOR!G13. If < 3 years, invest.

================================================================================

RULE 2: IF SHORT-TERM CASH IS LOW → BUY CREDITS
-----------------------------------------------
CO2 Credits are an OPERATING expense - pay as you go.

WHY CREDITS?
  - No upfront CAPEX required
  - Immediate CO2 offset (1 credit = 1 ton)
  - Flexibility: Buy only what you need

WHEN TO USE:
  - Cash-strapped situations
  - Temporary spikes in emissions
  - Meet compliance quickly

COST COMPARISON:
  Credit Cost = $25/ton (typical)
  Tax Cost = $30/ton (if you don't abate)
  
  If Credit Cost < Tax Cost → Buy credits to avoid higher tax

WARNING: Credits provide NO long-term benefit. They're a quick fix.

================================================================================

RULE 3: TREES FOR PR & LONG-TERM
--------------------------------
Trees are LOW COST but SLOW to impact CO2.

TYPICAL VALUES:
  - Cost: $50/tree
  - CO2 Reduction: 0.02 tons/tree/year

Trees are great for:
  - Corporate image / ESG reporting
  - Long-term sustainability goals
  - Community engagement

NOT great for:
  - Meeting immediate emission targets
  - Quick ROI

================================================================================

RULE 4: GREEN ELECTRICITY FOR OPERATING BUDGET
----------------------------------------------
Switching to green electricity is an OPERATING expense.

INPUT: Percentage of consumption to convert (0-100%)

CALCULATION:
  Cost = Energy_Consumption × % × Price_Premium
  CO2_Reduced = Energy_Consumption × % × Reduction_Rate

EXAMPLE:
  - Energy: 500,000 kWh
  - Convert: 50%
  - Premium: $0.03/kWh
  - Cost: 500,000 × 0.5 × 0.03 = $7,500/year
  - CO2 Reduced: 500,000 × 0.5 × 0.0005 = 125 tons

Good option when:
  - Need predictable annual costs
  - Can't afford upfront solar investment
  - Want immediate, measurable impact

================================================================================
                    HOW TO USE THE DASHBOARD
================================================================================

STEP 1: IMPACT_CONFIG
  - Enter CO2 Tax Rate from Case Guide
  - Verify initiative cost/reduction rates

STEP 2: STRATEGY_SELECTOR
  - Enter current emissions in B6
  - Enter energy consumption in B8
  - Adjust quantities in yellow cells
  - Review payback periods and cost per ton
  - Check "Cheapest $/Ton Option" indicator

STEP 3: INTERPRET RESULTS
  - Best option highlighted in GREEN
  - Compare payback periods for CAPEX
  - Ensure remaining emissions acceptable

STEP 4: UPLOAD_READY_ESG
  - Copy final quantities to ExSim

================================================================================
                    QUICK REFERENCE
================================================================================

| Initiative        | Type   | Best For                    |
|-------------------|--------|----------------------------|
| Solar PV          | CAPEX  | Long-term, high tax rates  |
| Trees             | CAPEX  | PR, low budget, long-term  |
| Green Electricity | OpEx   | Predictable annual cost    |
| CO2 Credits       | OpEx   | Quick fix, low cash        |

KEY METRICS:
  - Payback Period: CAPEX / Annual Savings (years)
  - Cost per Ton: Total Cost / Tons Reduced
  - Net Benefit: Tax Savings - Annual Cost (for OpEx)

GOAL: Minimize cost per ton while meeting emission targets.

================================================================================
                    COMMON PITFALLS
================================================================================

1. IGNORING TAX RATE CHANGES
   - If tax rates increase, solar becomes more attractive
   - Re-run analysis each period

2. OVER-INVESTING IN CREDITS
   - Credits don't build long-term value
   - Balance with solar/trees for sustainability

3. FORGETTING ENERGY CONSUMPTION
   - Green electricity sizing depends on kWh
   - Update B8 each period

4. NOT CHECKING PAYBACK
   - Column G shows years to recover investment
   - Anything > 5 years is risky

================================================================================
