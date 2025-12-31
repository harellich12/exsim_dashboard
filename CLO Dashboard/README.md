# ExSim Logistics Dashboard

## Overview

This Supply Network Optimization Dashboard helps Logistics Managers balance inventory across zones using shipments. It prevents stockouts and manages warehouse overflow situations.

## Setup

1. **Place data files in the `data/` folder:**
    - `finished_goods_inventory.xlsx`
    - `Logistics Decisions.xlsx` (optional template)
    - `logistics.xlsx` (optional costs)

2. **Run:**

   ```bash
   python generate_logistics_dashboard.py
   ```

3. **Open:** `Logistics_Dashboard.xlsx`

## The Four Tabs

### Tab 1: ROUTE_CONFIG

Configure the "physics" of your logistics network.

**Table 1: Transport Modes**

| Mode | Lead Time | Cost/Unit | Best For |
| :--- | :--- | :--- | :--- |
| Train | 2 FN | $5 | Cheap bulk, plan ahead |
| Truck | 1 FN | $10 | Balanced option |
| Plane | 0 FN | $25 | Expensive, emergencies |

**Table 2: Warehouse Configuration**

- Current capacity per zone.
- Cost to add modules (buy/rent).
- Capacity added per module.

### Tab 2: INVENTORY_TETRIS

Balance inventory across 5 zones.

**Each Zone Block Contains:**

| Column | Description |
| :--- | :--- |
| **Production** | INPUT: From Production Dashboard |
| **Sales** | INPUT: From Marketing Dashboard |
| **Outgoing** | INPUT: Negative for shipments OUT |
| **Incoming** | INPUT: Positive for shipments IN |
| **Projected Inv** | CALCULATED: Running balance |
| **Flag** | STOCKOUT or OVERFLOW warning |

**Flags:**

- 🔴 **RED "STOCKOUT: SHIP HERE!"** = Inventory went negative.
- 🟣 **PURPLE "OVERFLOW: RENT!"** = Inventory exceeds capacity.

### Tab 3: SHIPMENT_BUILDER

Plan your inter-zone transfers.

**Columns:**

- **Fortnight:** When you ORDER the shipment.
- **Origin:** Source zone.
- **Destination:** Target zone.
- **Material:** Default is "Electroclean".
- **Mode:** Train/Truck/Plane.
- **Quantity:** Units to ship.
- **Lead Time:** Auto-calculated from mode.
- **Arrival FN:** When it arrives.

> [!IMPORTANT]
> **Manual Step:** After adding shipments here, you MUST update Tab 2:
>
> 1. Add **NEGATIVE** qty to Origin's "Outgoing" in ORDER FN.
> 2. Add **POSITIVE** qty to Destination's "Incoming" in ARRIVAL FN.

### Tab 4: UPLOAD_READY_LOGISTICS

ExSim format with two sections side-by-side.

- **WAREHOUSES:** Zone module decisions.
- **SHIPMENTS:** Links to SHIPMENT_BUILDER table.

## Transport Modes Explained

**TRAIN: Cheap but slow**

- **Lead Time:** 2 fortnights
- **Cost:** $5 per unit
- **Use when:** You can plan 2 FN ahead, bulk transfers

**TRUCK: Balanced option**

- **Lead Time:** 1 fortnight
- **Cost:** $10 per unit
- **Use when:** Normal operations, moderate urgency

**PLANE: Fast but expensive**

- **Lead Time:** 0 fortnights (same period arrival)
- **Cost:** $25 per unit
- **Use when:** Emergency stockout prevention ONLY

## Handling Overflow

When inventory exceeds warehouse capacity, you have two options:

**Option 1: RENT MODULES (Tab 2, "Rent Modules?" cell)**

- Temporarily adds capacity.
- **Cost:** ~$50,000 per module per period.
- **Capacity:** 400 units per module.
- **Use when:** Temporary peak, one-time cost acceptable.

**Option 2: SHIP OUT (Tab 3)**

- Transfer excess to another zone.
- **Cost:** Transport cost per unit.
- **Use when:** Another zone needs inventory anyway.

## Preventing Stockouts

When you see **"STOCKOUT: SHIP HERE!"** flag:

1. **CHECK** other zones for excess inventory.
2. **PLAN** a shipment in `SHIPMENT_BUILDER` (Tab 3).
3. **REMEMBER** the lead time:
    - If stockout is in FN5 and using Truck (lead=1).
    - You must ORDER in FN4 for arrival in FN5.
4. **UPDATE** Tab 2 manually:
    - Origin zone: Add -500 to FN4 Outgoing.
    - Destination zone: Add +500 to FN5 Incoming.

## Lead Time Logic
>
> Arrival Fortnight = Order Fortnight + Lead Time

**Examples:**

| Mode | Order FN | Lead | Arrives FN |
| :--- | :--- | :--- | :--- |
| Train | 1 | 2 | 3 |
| Truck | 2 | 1 | 3 |
| Plane | 3 | 0 | 3 |

## Color Coding

- 🟡 **YELLOW:** User input cells
- 🟢 **GREEN:** Calculated cells
- 🔘 **GRAY:** Reference data
- 🔴 **RED:** Stockout warning
- 🟣 **PURPLE:** Overflow warning

## Strategic Workflow

1. **GET DATA from other dashboards:**
    - Production targets from CPO Dashboard.
    - Sales forecasts from CMO Dashboard.

2. **ENTER in INVENTORY_TETRIS (Tab 2):**
    - Production in each zone.
    - Expected sales in each zone.

3. **REVIEW FLAGS:**
    - Any STOCKOUT? Plan incoming shipments.
    - Any OVERFLOW? Rent modules or ship out.

4. **BUILD SHIPMENTS (Tab 3):**
    - Add transfer records.
    - Note arrival fortnights.

5. **UPDATE Tab 2 Manually:**
    - Outgoing in origin zone.
    - Incoming in destination zone (shifted by lead).

6. **VERIFY flags are cleared.**

7. **COPY to ExSim from Tab 4.**

---
*Balance your supply network across all zones!*
