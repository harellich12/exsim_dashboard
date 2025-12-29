================================================================================
       EXSIM CMO COMPLETE DASHBOARD - README
================================================================================

OVERVIEW
--------
This comprehensive dashboard integrates Marketing, Innovation, Inventory,
and Segment Analysis into a single decision-support tool for ExSim Marketing
Managers. It provides real-time ROI calculations and ExSim-ready upload formats.


SETUP
-----
1. Place these files in the data/ folder:
   - market-report.xlsx
   - Marketing Decisions.xlsx
   - Marketing Innovation Decisions.xlsx
   - sales_admin_expenses.xlsx
   - finished_goods_inventory.xlsx

2. Run: python generate_cmo_dashboard_complete.py

3. Open: CMO_Dashboard_Complete.xlsx


THE FIVE TABS EXPLAINED
-----------------------

>>> TAB 1: SEGMENT_PULSE
    Market allocation analysis split by HIGH and LOW customer segments.

    COLUMNS:
    - My Market Share: Your current share (with data bars visualization)
    - Awareness Gap: My Awareness minus Competitor Average
    - Price Gap: How much more/less you charge vs competitors (%)
    - Attractiveness: Product appeal score

    ALLOCATION FLAGS:
    | Flag                              | Meaning                           |
    |-----------------------------------|-----------------------------------|
    | OK (Green)                        | No immediate action needed        |
    | CRITICAL: Boost TV for Allocation | High segment awareness <30%       |
    | RISK: Losing Volume to Price      | Low segment, you're 5%+ pricier   |

    WHY IT MATTERS: The simulation's Market Allocation Engine considers these
    factors when distributing demand. Fix red flags before setting forecasts.


>>> TAB 2: INNOVATION_LAB
    Select product improvements to boost Attractiveness.

    HOW TO USE:
    1. Set Decision column to 1 (Yes) or 0 (No) for each feature
    2. Enter estimated cost per feature
    3. Total Innovation Cost auto-calculates

    FEATURES (Loaded from your template):
    - Stainless Material, Recyclable Materials, Energy Efficiency
    - Lighter/Compact, Impact Resistance, Noise Reduction
    - Battery Capacity, Self-Cleaning, Speed Settings
    - Digital Controls, Voice Assistance, Automation
    - Multifunctional Accessories, Mapping Technology

    TIP: Innovations are essential for High Segment allocation. If your
    High Segment share is low, consider investing in 2-3 features.


>>> TAB 3: STRATEGY_COCKPIT
    The main decision engine with ROI projections.

    SECTION A - GLOBAL ALLOCATIONS:
    | Input          | Description                                |
    |----------------|-------------------------------------------|
    | TV Budget ($)  | National TV spend - drives High Awareness |
    | Brand Focus    | 0=Awareness focus, 100=Attributes focus   |

    SECTION B - ZONAL ALLOCATIONS (per zone):
    | Column         | Type   | Description                         |
    |---------------|--------|-------------------------------------|
    | Last Sales    | Gray   | Previous round actual sales         |
    | Stockout?     | Gray   | "TRUE DEMAND HIGHER" if inventory=0 |
    | Target Demand | Yellow | YOUR INPUT - set your forecast      |
    | Radio Budget  | Yellow | Zone-specific radio spend           |
    | Salespeople   | Yellow | Number of salespeople               |
    | Price         | Yellow | Selling price for this zone         |
    | Payment       | Yellow | A=strict, B=normal, C=lenient       |
    | Est. Revenue  | Green  | Calculated: Demand × Price          |
    | Mkt Cost      | Green  | TV/5 + Radio + Salaries + Innov/5   |
    | Contribution  | Green  | Revenue - MarketingCost - COGS      |

    ROI CALCULATION:
    Contribution Margin = Est. Revenue - Marketing Cost - (Demand × $40 COGS)

    IF STOCKOUT OCCURRED:
    The "Stockout?" column shows "TRUE DEMAND HIGHER" - this means actual
    demand exceeded your supply. Set Target Demand HIGHER than Last Sales
    to capture the full market opportunity.


>>> TAB 4: UPLOAD_READY_MARKETING
    Formatted exactly like ExSim's Marketing Decisions upload template.
    All values link to STRATEGY_COCKPIT - no re-entry needed!

    SIDE-BY-SIDE LAYOUT:
    | Campaigns (TV/Radio) | Demand by Zone | Pricing | Channels |

    HOW TO USE:
    Simply copy-paste these values into ExSim's Marketing form.


>>> TAB 5: UPLOAD_READY_INNOVATION
    Formatted exactly like ExSim's Innovation Decisions template.
    Links to INNOVATION_LAB - shows 1 or 0 for each feature.


STRATEGIC WORKFLOW
------------------
1. START with SEGMENT_PULSE:
   - Check allocation flags
   - Note if High segment needs awareness boost
   - Note if Low segment is losing to price

2. GO TO INNOVATION_LAB (if needed):
   - Select features to improve Attractiveness
   - This helps High segment allocation

3. SET DECISIONS in STRATEGY_COCKPIT:
   - Adjust TV Budget based on awareness needs
   - Set zonal Radio budgets (helps Low segment)
   - If stockout occurred, increase Target Demand
   - Check Contribution Margin is positive

4. VERIFY in UPLOAD_READY tabs:
   - Values auto-populate from your decisions
   - Copy directly to ExSim


BRAND FOCUS EXPLAINED (0-100)
-----------------------------
| Value  | Effect                                          |
|--------|------------------------------------------------|
| 0-30   | Awareness-focused: Best for low-awareness zones |
| 40-60  | Balanced approach for steady markets            |
| 70-100 | Attributes-focused: Justifies premium pricing   |

RULE OF THUMB: If Awareness < 50%, keep Brand Focus < 50.


TROUBLESHOOTING
---------------
- "0" values: Check data files are in data/ folder
- Innovation cost not updating: Ensure formula links are intact
- Stockout not detected: Verify finished_goods_inventory.xlsx has data


================================================================================
                    Make data-driven marketing decisions!
================================================================================
