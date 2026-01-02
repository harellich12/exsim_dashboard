# ExSim Dashboard Suite - User Manual

**Version 1.0 | January 2026**

Welcome to the ExSim Dashboard Suite! This manual provides step-by-step instructions for using each of the seven integrated decision-support dashboards. These tools are designed to help you optimize every aspect of your ExSim business simulation experience.

---

## Table of Contents

1. [Getting Started](#getting-started)
2. [CMO Dashboard (Marketing)](#1-cmo-dashboard-marketing--innovation)
3. [Production Dashboard](#2-production-dashboard-zone-based-manufacturing)
4. [Purchasing Dashboard](#3-purchasing-dashboard-mrp--sourcing)
5. [Logistics Dashboard](#4-logistics-dashboard-inventory--shipments)
6. [CPO Dashboard (HR)](#5-cpo-dashboard-workforce--compensation)
7. [ESG Dashboard](#6-esg-dashboard-sustainability--co2)
8. [CFO Dashboard (Finance)](#7-cfo-dashboard-finance--liquidity)
9. [Recommended Workflow](#recommended-workflow)
10. [Troubleshooting](#troubleshooting)

---

## Getting Started

### Prerequisites

Before generating dashboards, ensure you have:

- Python 3.8+ installed
- Required libraries: `pip install pandas openpyxl`

### Preparing Your Data

1. Export your ExSim reports to a `Reports` folder (or the local `data/` folder in each dashboard directory).
2. Each dashboard will automatically load the required files.

### Generating a Dashboard

Navigate to any dashboard folder and run:

```bash
python generate_<dashboard_name>.py
```

The generated `.xlsx` file will appear in the same folder.

---

## 1. CMO Dashboard (Marketing & Innovation)

**Purpose:** Optimize your marketing mix, manage product innovation, and ensure your products are allocated by the simulation's market engine.

### Tabs Overview

| Tab | Purpose |
| :--- | :--- |
| **SEGMENT_PULSE** | Analyze your position in High/Low customer segments |
| **INNOVATION_LAB** | Select product features to invest in |
| **STRATEGY_COCKPIT** | Set budgets, prices, and demand forecasts |
| **UPLOAD_READY_MARKETING** | Copy-paste values to ExSim |
| **UPLOAD_READY_INNOVATION** | Copy-paste innovation decisions |

### Step-by-Step Instructions

1. **Open `SEGMENT_PULSE`**
   - Review your **Market Share** and **Awareness Gap** for each segment.
   - Check the **Allocation Flag** column:
     - 🟢 **OK:** No action needed.
     - 🔴 **CRITICAL: Boost TV for Allocation:** Your High segment awareness is below 30%. Increase TV budget.
     - 🟠 **RISK: Losing Volume to Price:** You are priced 5%+ higher than competitors in the Low segment. Consider a price cut.

2. **Go to `INNOVATION_LAB` (if needed)**
   - Set the **Decision** column to `1` for features you want to invest in.
   - The **Total Innovation Cost** updates automatically.
   - Investing in 2-3 features can significantly boost your High segment appeal.

3. **Set Decisions in `STRATEGY_COCKPIT`**
   - **Global Section:**
     - Set your **TV Budget** (drives national awareness for High segment).
     - Set **Brand Focus** (0-30 = awareness focus, 70-100 = attribute focus).
   - **Zonal Section (per zone):**
     - Enter **Target Demand** (your sales forecast). If the "Stockout?" column shows "TRUE DEMAND HIGHER", set this value *above* your last sales.
     - Set **Radio Budget**, **Salespeople**, **Price**, and **Payment Terms**.
   - Review the calculated **Est. Revenue** and **Contribution Margin**.

4. **Copy to ExSim**
   - Go to `UPLOAD_READY_MARKETING` and `UPLOAD_READY_INNOVATION`.
   - Copy the values directly into ExSim's Marketing and Innovation forms.

---

## 2. Production Dashboard (Zone-Based Manufacturing)

**Purpose:** Plan production output per zone, assign machines and workers, and identify capacity bottlenecks.

### Tabs Overview

| Tab | Purpose |
| :--- | :--- |
| **ZONE_CALCULATORS** | Set production targets for each of the 5 zones |
| **RESOURCE_MGR** | Manage machine assignments and expansion plans |
| **UPLOAD_READY_PRODUCTION** | Copy-paste values to ExSim |

### Step-by-Step Instructions

1. **Open `ZONE_CALCULATORS`**
   - Each zone (Center, West, North, East, South) has its own block.
   - For each zone:
     - Enter **Target Production** for FN1-FN8.
     - Set **Overtime?** to `Y` or `N`.
   - Review the calculated **Real Output**. If it's lower than your target, you have a capacity constraint.
   - Check for **"SHIPMENT NEEDED!"** alerts. This means local materials are insufficient—you need to order materials to be delivered to that zone.

2. **Go to `RESOURCE_MGR`**
   - **Section A (Assignments):** Assign machines and workers to specific sections within each zone.
   - **Section B (Expansion):** Review the **Capacity Gap** and **Recommendation** (e.g., "Buy 3 M1 machines for West").
   - **Section C (Modules):** If a zone shows "Buy module in [Zone]", you must purchase a factory module in ExSim before you can add machines there.

3. **Copy to ExSim**
   - Go to `UPLOAD_READY_PRODUCTION` and copy the values into ExSim's Production form.

> [!IMPORTANT]
> **Zone Independence:** Resources in one zone (machines, materials, workers) cannot be used in another zone. If you want to produce in a new zone, you must first buy a module, then machines, then hire/transfer workers, and finally order materials.

---

## 3. Purchasing Dashboard (MRP & Sourcing)

**Purpose:** Plan material procurement using Material Requirements Planning (MRP), manage lead times, and optimize batch sizes.

### Tabs Overview

| Tab | Purpose |
| :--- | :--- |
| **SUPPLIER_CONFIG** | Set up supplier data (costs, lead times, batch sizes) |
| **COST_ANALYSIS** | Evaluate your ordering vs. holding cost efficiency |
| **MRP_ENGINE** | Calculate material requirements and place orders |
| **CASH_FLOW_PREVIEW** | Track procurement spending over time |
| **UPLOAD_READY_PROCUREMENT** | Copy-paste values to ExSim |

### Step-by-Step Instructions

1. **Configure `SUPPLIER_CONFIG`**
   - Enter the data from your Case Guide: **Lead Time**, **Cost/Unit**, **Payment Terms**, and **Batch Size** for each supplier.

2. **Review `COST_ANALYSIS`**
   - Check your **Ordering Cost Ratio**:
     - **> 70%:** You are ordering too frequently. Increase batch sizes.
     - **< 30%:** You are holding too much inventory. Decrease batch sizes or use Just-In-Time.
     - **30-70%:** Balanced. Maintain your current policy.

3. **Plan in `MRP_ENGINE`**
   - **Section A:** Enter **Target Production** (from the Production Dashboard) for FN1-FN8.
   - **Section B:** Review **Projected Inventory**. A **RED** cell means you will have a stockout in that fortnight.
   - **Section C:** Enter order quantities by supplier.

> [!WARNING]
> **"Time Travel" Rule:** When you order in Fortnight X, goods arrive in Fortnight X + Lead Time.
>
> - Example: If you need stock in FN4 and your supplier has a 2 FN lead time, you must order in **FN2**.

1. **Verify `CASH_FLOW_PREVIEW`**
   - Check that your **Cumulative Spend** is within your budget.

2. **Copy to ExSim**
   - Go to `UPLOAD_READY_PROCUREMENT` and copy the order matrix into ExSim.

---

## 4. Logistics Dashboard (Inventory & Shipments)

**Purpose:** Balance finished goods inventory across zones by planning inter-zone shipments and managing warehouse capacity.

### Tabs Overview

| Tab | Purpose |
| :--- | :--- |
| **ROUTE_CONFIG** | View transport modes and warehouse costs |
| **INVENTORY_TETRIS** | Balance inventory across 5 zones |
| **SHIPMENT_BUILDER** | Plan inter-zone transfers |
| **UPLOAD_READY_LOGISTICS** | Copy-paste values to ExSim |

### Step-by-Step Instructions

1. **Review `ROUTE_CONFIG`**
   - Understand the three transport modes:
     - **Train:** 2 FN lead time, $5/unit (cheapest, requires planning ahead).
     - **Truck:** 1 FN lead time, $10/unit (balanced).
     - **Plane:** 0 FN lead time, $25/unit (expensive, emergency only).

2. **Open `INVENTORY_TETRIS`**
   - For each zone, enter:
     - **Production** (from Production Dashboard).
     - **Sales** (your expected demand from the CMO Dashboard).
   - Review the **Flag** column:
     - 🔴 **STOCKOUT!:** Inventory went negative. You need to ship goods *to* this zone.
     - 🟣 **OVERFLOW!:** Inventory exceeds warehouse capacity. Rent a module or ship goods *out*.
     - 🟡 **WARNING: >90%:** Inventory is nearing capacity.

3. **Plan Shipments in `SHIPMENT_BUILDER`**
   - Add a row for each transfer:
     - **Fortnight:** When you *order* the shipment.
     - **Origin/Destination:** Zones.
     - **Mode:** Train/Truck/Plane.
     - **Quantity:** Units to ship.
   - The **Arrival FN** is calculated automatically based on the mode.

4. **Update `INVENTORY_TETRIS` Manually**
   - After adding a shipment:
     - In the **Origin** zone, add a **negative** value to the "Outgoing" column for the *order* fortnight.
     - In the **Destination** zone, add a **positive** value to the "Incoming" column for the *arrival* fortnight.
   - Verify that all flags are cleared (show "✓ OK").

5. **Copy to ExSim**
   - Go to `UPLOAD_READY_LOGISTICS` and copy the warehouse and shipment decisions.

---

## 5. CPO Dashboard (Workforce & Compensation)

**Purpose:** Manage workforce planning, set salaries to avoid strikes, and configure employee benefits.

### Tabs Overview

| Tab | Purpose |
| :--- | :--- |
| **WORKFORCE_PLANNING** | Calculate hiring/firing needs by zone |
| **COMPENSATION_STRATEGY** | Set salaries and benefits to prevent strikes |
| **LABOR_COST_ANALYSIS** | Calculate total labor expense for Finance |
| **UPLOAD_READY_PEOPLE** | Copy-paste values to ExSim |

### Step-by-Step Instructions

1. **Open `WORKFORCE_PLANNING`**
   - Review **Current Staff** per zone.
   - Enter **Required Workers** (from the Production Dashboard).
   - Set a realistic **Turnover Rate** (5% is typical; 10% if morale is low).
   - Review the calculated **Hiring Needed** and **Hiring Cost**.

2. **Set Salaries in `COMPENSATION_STRATEGY`**
   - **First:** Enter the **Inflation Rate** from your Case Guide (e.g., 3%).
   - For each zone, set a **Proposed Salary**.
   - Check the **Strike Risk** column:
     - If it shows **"STRIKE RISK!"**, your proposed salary is below the minimum required. Increase it.

> [!IMPORTANT]
> **The Formula:** Minimum Safe Salary = Previous Salary × (1 + Inflation Rate).
> Example: $750 × 1.03 = $772.50 minimum to avoid a strike.

1. **Configure Benefits**
   - **Training Budget:** 2-5% of payroll (reduces defects).
   - **Health Insurance:** 3-5% of payroll (reduces absenteeism).
   - **Profit Sharing:** 5-10% of net profit (boosts morale).

2. **Review `LABOR_COST_ANALYSIS`**
   - Enter your **Estimated Net Profit** to calculate the profit-sharing amount.
   - Share the **Total People Expense** with the CFO.

3. **Copy to ExSim**
   - Go to `UPLOAD_READY_PEOPLE` and copy the values into ExSim's People form.

---

## 6. ESG Dashboard (Sustainability & CO2)

**Purpose:** Compare green investment options (Solar, Trees, Credits) and determine the most cost-effective way to reduce your CO2 tax burden.

### Tabs Overview

| Tab | Purpose |
| :--- | :--- |
| **IMPACT_CONFIG** | Set CO2 tax rates and initiative parameters |
| **STRATEGY_SELECTOR** | Compare ROI of different abatement strategies |
| **UPLOAD_READY_ESG** | Copy-paste values to ExSim |

### Step-by-Step Instructions

1. **Configure `IMPACT_CONFIG`**
   - Enter the **CO2 Tax Rate** from your Case Guide.
   - Verify the cost and reduction rates for each initiative (Solar, Trees, Green Electricity, Credits).

2. **Make Decisions in `STRATEGY_SELECTOR`**
   - Enter your **Current Emissions** (tons).
   - Enter your **Energy Consumption** (kWh).
   - Adjust the quantities for each initiative in the yellow cells.
   - Review:
     - **Payback Period:** For CAPEX investments (Solar, Trees). Anything under 3 years is good.
     - **Cost per Ton:** Lower is better.
     - **Net Benefit:** For OpEx options (Credits, Green Electricity). Positive is good.

3. **Decision Rules**
   - **Buy Solar PV:** If the Payback Period is < 3 years and you have the upfront capital.
   - **Buy CO2 Credits:** If you are cash-strapped and need to meet a short-term target quickly.
   - **Plant Trees:** For long-term sustainability and PR benefits. Low ROI but good for image.
   - **Switch to Green Electricity:** If you want predictable annual costs without large upfront investment.

4. **Copy to ExSim**
   - Go to `UPLOAD_READY_ESG` and copy the investment quantities.

---

## 7. CFO Dashboard (Finance & Liquidity)

**Purpose:** Forecast cash flow, audit profitability, and manage debt to ensure solvency and maintain credit ratings.

### Tabs Overview

| Tab | Purpose |
| :--- | :--- |
| **LIQUIDITY_MONITOR** | Track cash flow across fortnights |
| **PROFIT_CONTROL** | Project income statement and compare to actuals |
| **BALANCE_SHEET_HEALTH** | Monitor debt ratio and credit risk |
| **DEBT_MANAGER** | Calculate mortgage payments |
| **UPLOAD_READY_FINANCE** | Copy-paste values to ExSim |

### Step-by-Step Instructions

1. **Start with `PROFIT_CONTROL`**
   - Review the **Last Round Actuals** column.
   - Enter your **This Round Projected** values (Revenue, S&A, Depreciation, Interest).
   - Check the **Variance %** column. Large variances (>20%) require justification.
   - Watch the **Profit Realism Flag**. If it says "WARNING: Unrealistic profit jump!", your forecast may be too optimistic.

2. **Check `BALANCE_SHEET_HEALTH`**
   - Review your **Current Debt Ratio** (Liabilities / Assets).
   - Plan any new debt carefully. If your ratio exceeds 60%, lenders may refuse credit or charge premium rates.
   - Look for warning flags:
     - **"CRITICAL: Debt too high":** Credit rating risk.
     - **"CRITICAL: Equity Erosion":** Retained earnings are negative.

3. **Manage `LIQUIDITY_MONITOR`**
   - **Section A (Initialization):** Enter one-time deductions (taxes, dividends, asset purchases).
   - **Section B (Operational):** Enter estimates for Sales Receipts, Procurement Spend, etc. (from other dashboards).
   - **Section C (Financing):** Adjust Credit Line, Investment, and Mortgage values.
   - **Section D (Cash Balance):** Review the **Ending Cash** row.
     - 🔴 **RED (< $0):** Insolvency risk! Adjust financing or reduce spending.
     - 🟢 **GREEN (> $200k):** Inefficient cash holdings. Consider investing.

4. **Configure `DEBT_MANAGER`**
   - Enter any new **Loan Amount**, **Interest Rate**, and **Payment Schedule**.
   - Review the **Total Payments**.

5. **Copy to ExSim**
   - Go to `UPLOAD_READY_FINANCE` and copy the values into ExSim's Finance form.

---

## Recommended Workflow

For the best results, complete the dashboards in the following order:

```
┌─────────────────────────────────────────────────────────────────────────┐
│                        RECOMMENDED WORKFLOW                              │
├─────────────────────────────────────────────────────────────────────────┤
│                                                                          │
│   ┌─────────┐                                                            │
│   │ 1. CMO  │ ──────────────────┐                                        │
│   │  (Set   │                   │                                        │
│   │ Demand) │                   ▼                                        │
│   └─────────┘            ┌──────────────┐                                │
│                          │ 2. Production│                                │
│   ┌─────────┐            │ (Plan Output)│                                │
│   │ 6. ESG  │            └──────────────┘                                │
│   │ (Green  │                   │                                        │
│   │ Invest) │       ┌──────────┬┴──────────┬──────────┐                  │
│   └────┬────┘       │          │           │          │                  │
│        │            ▼          ▼           ▼          │                  │
│        │     ┌──────────┐ ┌─────────┐ ┌─────────┐     │                  │
│        │     │    3.    │ │   4.    │ │   5.    │     │                  │
│        │     │Purchasing│ │Logistics│ │   CPO   │     │                  │
│        │     │ (Order   │ │(Balance │ │ (Hire   │     │                  │
│        │     │Materials)│ │Inventory│ │Workers) │     │                  │
│        │     └────┬─────┘ └────┬────┘ └────┬────┘     │                  │
│        │          │            │           │          │                  │
│        │          └────────────┴───────────┘          │                  │
│        │                       │                      │                  │
│        │                       ▼                      │                  │
│        │               ┌──────────────┐               │                  │
│        └──────────────►│   7. CFO     │◄──────────────┘                  │
│                        │(Verify Cash) │                                  │
│                        └──────────────┘                                  │
│                                                                          │
└─────────────────────────────────────────────────────────────────────────┘
```

1. **CMO:** Set your demand forecast and marketing strategy.
2. **Production:** Convert demand into a feasible production plan.
3. **Purchasing:** Order the materials needed for production.
4. **Logistics:** Ship finished goods to balance inventory across zones.
5. **CPO:** Ensure you have enough workers and set fair salaries.
6. **ESG:** Plan any green investments.
7. **CFO:** Aggregate all spending and verify solvency.

---

## Troubleshooting

| Problem | Solution |
| :--- | :--- |
| Dashboard shows all zeros | Ensure data files are in the `data/` folder and re-run the script. |
| "STOCKOUT!" flag won't clear | Add a shipment in `SHIPMENT_BUILDER` and update the "Incoming" column in `INVENTORY_TETRIS`. |
| "STRIKE RISK!" for all zones | Enter the correct **Inflation Rate** in `COMPENSATION_STRATEGY` and set salaries above the minimum. |
| Cash goes negative | Increase Credit Line, reduce Procurement Spend, or delay Asset Purchases. |
| Payback period is infinite | CO2 tax rate may be too low to justify the solar investment. Use Credits instead. |

---

**Need more help?** Refer to the individual `README.md` file in each dashboard's folder for detailed technical documentation.
