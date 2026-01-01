# ExSim Purchasing Dashboard v2

## Overview

This MRP & Cost-Optimized Sourcing Dashboard helps Purchasing Managers plan material procurement, handle lead time shifts, and optimize batch sizes based on ordering vs holding cost trade-offs.

## Setup

1. **Place these files in the `data/` folder:**
    - `raw_materials.xlsx` (for opening inventory)
    - `production.xlsx` (for cost analysis)
    - `Procurement Decisions.xlsx` (optional, for template)

2. **Run:**

   ```bash
   python generate_purchasing_dashboard_v2.py
   ```

3. **Open:** `Purchasing_Dashboard.xlsx`

## The Five Tabs Explained

### Tab 1: SUPPLIER_CONFIG

Configure your supplier data from the case study.

**Parts Suppliers:**

| Field | Description |
| :--- | :--- |
| **Supplier** | Supplier name (A1, A2, B1, B2) |
| **Lead Time (FN)** | Fortnights until delivery |
| **Cost/Unit** | Purchase price per unit |
| **Payment Terms (FN)** | Fortnights until payment is due |
| **Batch Size** | Minimum order quantity |

**Pieces Config:**

| Field | Description |
| :--- | :--- |
| **Cost/Unit** | Purchase price per piece |
| **Batch Size** | Minimum order quantity |

### Tab 2: COST_ANALYSIS

Evaluates if your previous ordering was efficient.

**Key Metric: Ordering Cost Ratio**
> Formula: Ordering Cost / (Ordering Cost + Holding Cost)

**Interpretation:**

| Ratio | Meaning | Action |
| :--- | :--- | :--- |
| **> 70%** | Ordering too frequently | **INCREASE** batch sizes |
| **< 30%** | Holding too much inventory | **DECREASE** batch sizes/JIT |
| **30-70%** | Balanced approach | Maintain current policy |

> [!NOTE]
> **Why it matters:**
>
> - High ordering costs = many small orders = high transaction costs.
> - High holding costs = too much stock = capital tied up, storage costs.
> - The goal is to find the Economic Order Quantity (EOQ) balance.

### Tab 3: MRP_ENGINE

Material Requirements Planning calculator.

**Section A - Production Demand:**
Enter Target Production from Production Manager for FN1-FN8.

**Section B - Net Requirements:**

| Row | Description |
| :--- | :--- |
| **Gross Requirement** | = Target Production (1:1 ratio assumed) |
| **Scheduled Arrivals** | Orders you've already placed (input) |
| **Projected Inventory** | = Previous + Arrivals - Gross |
| **Net Deficit** | = Shortfall if Inventory goes negative |

> [!WARNING]
> **RED HIGHLIGHT** = Projected Inventory is NEGATIVE = you need to order!

**Section C - Sourcing Strategy:**
Enter order quantities by supplier. Remember the "Time Travel" rule!

### Tab 4: CASH_FLOW_PREVIEW

Tracks procurement spending and cash outflows.

**Key Rows:**

- **Part A/B Orders:** Cost per fortnight.
- **Total Spend:** Sum of all procurement.
- **Cumulative Spend:** Running total.
- **Budget Tracking:** Compare against your budget.

### Tab 5: UPLOAD_READY_PROCUREMENT

Formatted for ExSim Procurement upload.

- **PARTS:** Matrix format with FN1-FN8 columns.
- **PIECES:** Single order quantity column.

All values link to `MRP_ENGINE` - just copy to ExSim!

## The "Time Travel" Logic (Lead Times)

This is the most important concept in purchasing planning.

> [!IMPORTANT]
> **RULE:** When you ORDER in Fortnight X, the goods ARRIVE in Fortnight X + Lead Time.

**Example:**

| Supplier | Lead Time | Order in FN1 | Arrives in |
| :--- | :--- | :--- | :--- |
| Supplier A | 1 FN | 500 units | FN2 |
| Supplier B | 2 FN | 300 units | FN3 |
| Supplier C | 3 FN | 200 units | FN4 |

**How to Use:**

1. Look at `MRP_ENGINE` Section B - find which FN has negative inventory.
2. Count **BACKWARDS** by the lead time to determine when to order.
3. If you need stock in FN4 and Lead Time is 2: **ORDER in FN4 - 2 = FN2**.

### Example Scenario

- FN3 shows Projected Inventory = -200 (deficit).
- You have Supplier A (Lead 1) and Supplier B (Lead 2).

**Solution Options:**

- Order from Supplier A in FN2 → arrives FN3 ✓
- Order from Supplier B in FN1 → arrives FN3 ✓

**WRONG:** Order from Supplier A in FN3 → arrives FN4 (TOO LATE!)

## Batch Size Compliance

Orders should be multiples of the supplier's batch size.

**Example:**

| Supplier | Batch Size | Valid Orders | Invalid |
| :--- | :--- | :--- | :--- |
| A1 | 500 | 500, 1000, 1500, 2000 | 300, 750 |
| B1 | 300 | 300, 600, 900, 1200 | 200, 500 |

> The `MRP_ENGINE` shows a "Batch Compliance Check" row to flag violations.

## Strategic Workflow

1. **START with COST_ANALYSIS:**
    - Check your Ordering Cost Ratio.
    - Adjust batch strategy if needed.

2. **CONFIGURE in SUPPLIER_CONFIG:**
    - Enter your case study supplier data.
    - Note lead times carefully.

3. **PLAN in MRP_ENGINE:**
    - Enter Target Production from Production Manager.
    - Review Projected Inventory (watch for RED = deficit).
    - Enter orders, remembering lead times.
    - Ensure batch compliance.

4. **VERIFY in CASH_FLOW_PREVIEW:**
    - Check cumulative spend vs budget.
    - Adjust orders if over budget.

5. **UPLOAD from UPLOAD_READY_PROCUREMENT:**
    - Copy values to ExSim.

## Common Mistakes to Avoid

1. Ordering in the **WRONG** fortnight (forgetting lead time).
2. Ordering non-batch quantities.
3. Over-ordering (creates holding costs).
4. Under-ordering (creates stockouts → production stops).

---
*Optimize your supply chain!*
