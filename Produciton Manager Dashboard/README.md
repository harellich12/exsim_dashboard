# ExSim Zone-Specific Production Dashboard

## Overview

This Zone-Specific Dashboard handles the physical separation of resources across geographic zones. Resources in CENTER do NOT count towards WEST capacity.

## Setup

1. **Place data files in the `data/` folder.**
2. **Run:**

   ```bash
   python generate_production_dashboard_zones.py
   ```

3. **Open:** `Production_Dashboard_Zones.xlsx`

## The Three Tabs

### Tab 1: ZONE_CALCULATORS

5 separate production blocks - one for each zone.

**Each Zone Block Contains:**

- **Zone Parameters:** local machines, materials, workers.
- **Production Schedule:** FN1-8.
- **Inputs:** Target Production, Overtime (Y/N).
- **Checks:** Local Capacity, Material Cap, REAL OUTPUT.
- **Alert:** "SHIPMENT NEEDED!" if target exceeds local materials.

### Tab 2: RESOURCE_MGR

**Section A: Assignments by Zone**

- Machine/worker allocation per section per zone.
- Grouped: Center Sec 1, Center Sec 2… West Sec 1…

**Section B: Expansion by Zone**

- Shows capacity gap per zone.
- **Recommendation:** "Buy X machines" per zone.

**Section C: Modules by Zone**

- Module slots vs machines per zone.
- **Recommendation:** "Buy module in [Zone]" if slots < 5.

### Tab 3: UPLOAD_READY_PRODUCTION

ExSim format with values linked from zone calculators.

## Expanding to a New Zone

To start producing in a new zone (e.g., East), you must:

1. **BUY A MODULE (Tab 2, Section C)**
   - A zone with 0 module slots cannot house machines.
   - Go to ExSim and purchase a module for East zone.

2. **BUY MACHINES (Tab 2, Section B)**
   - A zone with 0 machines has 0 production capacity.
   - Go to ExSim and purchase machines for East zone.

3. **HIRE WORKERS (External)**
   - Workers must be hired/transferred to the new zone.

4. **TRANSFER MATERIALS (Tab 1, Shipment Alert)**
   - If local materials = 0, you'll see "SHIPMENT NEEDED!".
   - Order materials to be delivered to the new zone.

> [!NOTE]
> **Step-by-Step Expansion Example (East Zone):**
>
> 1. In Tab 2, Section C: Note "Buy module in East".
> 2. In ExSim: Purchase 1 module for East.
> 3. In Tab 2, Section B: Set East Target Capacity.
> 4. In ExSim: Purchase recommended machines for East.
> 5. In Tab 2, Section A: Assign machines to East sections.
> 6. In Tab 1, East Block: Set production targets.

## Zone Independence Rule

Resources are PHYSICALLY SEPARATED:

- Center machines cannot produce for West
- West materials cannot be used in Center
- Workers in North cannot work in South

**Exception:** You can TRANSFER resources between zones, but:

- Machine transfers take 1 period.
- Material shipments must be ordered in advance.

## Zone Status at Game Start (Typical)

| Zone | Machines | Modules | Workers | Status |
| :--- | :--- | :--- | :--- | :--- |
| Center | 57 | 72 | 219 | **ACTIVE** |
| West | 0 | 0 | 71 | **READY** |
| North | 0 | 0 | 0 | EMPTY |
| East | 0 | 0 | 0 | EMPTY |
| South | 0 | 0 | 0 | EMPTY |

- **ACTIVE:** Producing immediately.
- **READY:** Has workers, can expand quickly.
- **EMPTY:** Must buy module + machines + hire workers.

## Color Coding

Each zone has a distinct color:

- 🔵 **Center:** Blue
- 🟠 **West:** Orange
- 🟢 **North:** Green
- 🟡 **East:** Yellow
- 🟤 **South:** Brown

---
*Manage your geographic production footprint!*
