"""
ExSim Mock Data Generator

Generates randomized but realistic mock data matching the exact format of /Reports Excel files.
Used by fire_test.py for end-to-end testing of dashboard generators.

Usage:
    python generate_mock_data.py [--seed 42] [--output-dir test_data/mock_reports]
"""

import pandas as pd
import numpy as np
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import argparse

# =============================================================================
# CONFIGURATION
# =============================================================================

ZONES = ["Center", "West", "North", "East", "South"]
FORTNIGHTS = 8
DEFAULT_SEED = 42
DEFAULT_OUTPUT_DIR = Path(__file__).parent / "mock_reports"

# Realistic bounds for generated values
BOUNDS = {
    "workers_per_zone": (50, 300),
    "absenteeism_rate": (0.00, 0.05),
    "overtime_pct": (0.00, 0.15),
    "salespeople": (20, 80),
    "salary_per_person": (500, 1500),
    "tv_spots": (10, 100),
    "radio_spots": (50, 500),
    "tv_cost_per_spot": (2500, 4000),
    "radio_cost_per_spot": (200, 400),
    "inventory_capacity": (1000, 5000),
    "production_units": (500, 2000),
    "unit_price": (50, 150),
    "discount_pct": (5, 15),
    "cash_balance": (20000, 200000),
    "receivables": (50000, 500000),
    "payables": (10000, 200000),
    "depreciation": (20000, 100000),
    "mortgage": (100000, 800000),
    "machines_per_zone": (0, 10),
    "transport_cost_per_unit": (5, 25),
}


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def create_header_rows(title, company="Company X", period=6, course="Test Course"):
    """Create standard 3-row header used in all reports."""
    return [
        [f"Country: {course}", None],
        [title, f"Company:{company}"],
        [f"Period: {period}", f"Course: {course}"],
    ]


def format_currency(value):
    """Format number as currency string."""
    if value < 0:
        return f"$-{abs(value):,.0f}"
    return f"${value:,.0f}"


def format_percent(value):
    """Format number as percentage string."""
    return f"{value:.1%}"


# =============================================================================
# MOCK DATA GENERATORS
# =============================================================================

def generate_workers_balance(rng, output_path):
    """Generate workers_balance_overtime.xlsx matching exact real format."""
    workers = {zone: rng.integers(*BOUNDS["workers_per_zone"]) for zone in ZONES}
    absenteeism = {zone: rng.uniform(*BOUNDS["absenteeism_rate"]) for zone in ZONES}
    overtime = {zone: rng.uniform(*BOUNDS["overtime_pct"]) for zone in ZONES}
    
    data = create_header_rows("Workers Balance")
    # Row 3: Section header
    data.append(["Workers Balance", None, None, None, None, None, None])
    # Row 4: Zone headers
    data.append([None, "Center", "West", "North", "East", "South", "Total"])
    # Row 5: Workers assigned
    data.append(["Workers assigned", workers["Center"], workers["West"], workers["North"], 
                 workers["East"], workers["South"], sum(workers.values())])
    # Row 6: FTE offset reduction
    data.append(["FTE offset of reduction of working hours", 0.00, 0.00, 0.00, 0.00, 0.00, 0.00])
    # Row 7: FTE offset off-days
    data.append(["FTE offset of off-days available for workers", 0.00, 0.00, 0.00, 0.00, 0.00, 0.00])
    # Row 8: Absenteeism
    data.append(["Absenteeism", 0.00, 0.00, 0.00, 0.00, 0.00, 0.00])
    # Row 9: FTE union reps
    data.append(["FTE union representatives", None, None, None, None, None, 0.00])
    # Row 10: Total
    total_workers = [float(workers[z]) for z in ZONES]
    data.append(["Total", *total_workers, sum(total_workers)])
    # Row 11: Empty
    data.append([None, None, None, None, None, None, None])
    # Row 12: Empty with 0.00 in last col (matches real format)
    data.append([None, None, None, None, None, None, 0.00])
    # Row 13: Empty
    data.append([None, None, None, None, None, None, None])
    # Row 14: Overtime header
    data.append(["% of Workforce Working Saturdays (Overtime) [%]", None, None, None, None, None, None])
    # Row 15: Zone headers
    data.append([None, "Center", "West", "North", "East", "South", "Total"])
    # Row 16: Overtime values
    data.append(["Total", "0.0%", "0.0%", "0.0%", "0.0%", "0.0%", "0.0%"])
    
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False, header=False)
    return {"workers": workers, "absenteeism": absenteeism}


def generate_sales_admin(rng, output_path):
    """Generate sales_admin_expenses.xlsx"""
    # Sales by zone
    units = {zone: rng.integers(1000, 10000) for zone in ZONES[:3]}  # Only 3 active zones
    price = {zone: rng.integers(*BOUNDS["unit_price"]) for zone in ZONES[:3]}
    discount = rng.uniform(*BOUNDS["discount_pct"]) / 100
    
    salespeople = rng.integers(*BOUNDS["salespeople"])
    salary_total = salespeople * rng.integers(*BOUNDS["salary_per_person"])
    tv_spots = rng.integers(*BOUNDS["tv_spots"])
    tv_cost = tv_spots * rng.integers(*BOUNDS["tv_cost_per_spot"])
    radio_spots = rng.integers(*BOUNDS["radio_spots"])
    radio_cost = radio_spots * rng.integers(*BOUNDS["radio_cost_per_spot"])
    
    data = create_header_rows("Sales & Administration Expenses")
    data.append(["Sales", None, None, None, None, None, None])
    data.append(["Region", "Brand", "Units", "Local Price", "Gross Sales", "Discount %", "Net Sales"])
    
    for zone in ZONES[:3]:
        gross = units[zone] * price[zone]
        net = gross * (1 - discount)
        data.append([zone, "A", f"{units[zone]:,}", format_currency(price[zone]),
                     format_currency(gross), f"{discount*100:.1f}", format_currency(net)])
    
    total_units = sum(units.values())
    total_gross = sum(units[z] * price[z] for z in ZONES[:3])
    total_net = total_gross * (1 - discount)
    data.append(["Total", None, f"{total_units:,}", None, format_currency(total_gross), None, format_currency(total_net)])
    
    data.append([None] * 7)
    data.append(["Sales & Administration Expenses", None, None, None, None, None, None])
    data.append([None, "Amount", "Expense", None, None, None, None])
    data.append(["Salespeople Salaries", f"{salespeople} Salespeople", format_currency(salary_total)])
    data.append(["Other channels costs", "-", "$0"])
    data.append(["Salespeople Hiring Expenses", "0 Salespeople", "$0"])
    data.append(["TV Advertising Expenses", f"{tv_spots} spots", format_currency(tv_cost)])
    data.append(["Radio Advertising Expenses", f"{radio_spots} spots", format_currency(radio_cost)])
    data.append(["Plant Modules Leasing Expenses", "0 Modules", "$0"])
    data.append(["Plant Module Administrative Expenses", "6 Modules", "$60,000"])
    data.append(["Executive Salaries", None, "$0"])
    data.append(["Total", None, format_currency(salary_total + tv_cost + radio_cost + 60000)])
    
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False, header=False)
    return {"salespeople": salespeople, "salary_total": salary_total}


def generate_finished_goods(rng, output_path):
    """Generate finished_goods_inventory.xlsx"""
    data = create_header_rows("Finished Goods Inventory")
    
    for zone_idx, zone in enumerate(ZONES):
        capacity = rng.integers(*BOUNDS["inventory_capacity"]) if zone_idx < 3 else 0
        
        data.append(["Finished Goods Inventory", None] + [None] * 10)
        data.append([f"Capacity: {capacity}", "Previous"] + list(range(1, 9)) + ["Total", "In transit"])
        
        initial_inv = [rng.integers(500, capacity) if capacity > 0 else 0 for _ in range(FORTNIGHTS)]
        production = [rng.integers(100, 500) if capacity > 0 else 0 for _ in range(FORTNIGHTS)]
        sales = [rng.integers(100, 400) if capacity > 0 else 0 for _ in range(FORTNIGHTS)]
        
        data.append(["Initial inventory", None] + initial_inv + [None, None])
        data.append(["Receptions", None] + [0] * FORTNIGHTS + [0, 0])
        data.append(["Production", None] + production + [sum(production), None])
        data.append(["Shipments", None] + [0] * FORTNIGHTS + [0, None])
        data.append(["Sales", None] + sales + [sum(sales), None])
        
        final_inv = [max(0, initial_inv[i] + production[i] - sales[i]) for i in range(FORTNIGHTS)]
        data.append(["Final inventory", initial_inv[0] if initial_inv else 0] + final_inv + [None, None])
        data.append(["Thrown away", 0] + [0] * FORTNIGHTS + [0, None])
        data.append([None] * 12)
    
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False, header=False)


def generate_initial_cash_flow(rng, output_path):
    """Generate initial_cash_flow.xlsx"""
    cash_start = rng.integers(*BOUNDS["cash_balance"])
    
    data = create_header_rows("Initial Cash Flow")
    data.append(["Initial Cash Flow", None])
    data.append([None, None])
    
    items = [
        ("Tax Payments", -rng.integers(10000, 50000)),
        ("Payments to Suppliers", -rng.integers(5000, 30000)),
        (None, None),
        ("Worker Hiring and Dismissal Expenses", -rng.integers(0, 5000)),
        ("Labor Benefits", 0),
        (None, None),
        ("Machines Rental & Transfer", 0),
        ("Modules leasing", 0),
        ("Warehouse Rental Expenses", -rng.integers(20000, 100000)),
        ("ESG", 0),
        (None, None),
        ("Interest Paid on Mortgages", -rng.integers(10000, 40000)),
        ("Interest Paid on Emergency Loan", 0),
    ]
    
    operations_total = sum(v for _, v in items if v is not None and v != 0)
    items.append(("Cash Flow from Operations", operations_total))
    items.append((None, None))
    items.append(("Product Improvements", 0))
    items.append(("Equipment and Modules Purchases", 0))
    items.append(("Sale of PPE", 0))
    items.append(("Cash Flow from Investments", 0))
    items.append((None, None))
    
    financing = rng.integers(50000, 200000)
    items.append(("Mortgage Increases", financing))
    items.append(("Mortgage Principal Payments", 0))
    items.append(("New Shares Issued", 0))
    items.append(("Dividend Payments", 0))
    items.append(("Emergency Loans", 0))
    items.append(("Cash Flow from Financing", financing))
    items.append((None, None))
    
    variation = operations_total + financing
    items.append(("Cash Variation", variation))
    items.append(("Initial cash (at the end of last period)", cash_start))
    items.append(("Final cash (at the start of the first fortnight)", cash_start + variation))
    
    for label, value in items:
        if value is None:
            data.append([label, None])
        else:
            data.append([label, format_currency(value)])
    
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False, header=False)
    return {"cash_start": cash_start, "cash_end": cash_start + variation}


def generate_production(rng, output_path):
    """Generate production.xlsx matching exact 860-row real format with 5 zones × 5 sections."""
    data = create_header_rows("Production Cost")
    
    # Generate for each zone (Center, West, North = active; East, South = minimal)
    for zone_idx, zone in enumerate(ZONES):
        is_active = zone_idx < 3
        
        # Add zone separator
        if zone_idx > 0:
            data.append([f"--- {zone} ---", None, None, None])
        
        # --- SECTION 1: Part A -> Assembly A ---
        # Block size: ~12-14 rows
        for _ in range(2):  # Header appears twice
            data.append(["Section 1", None, None, None])
            if _ == 0:
                # Cost summary
                part_a_cost = rng.integers(40000, 80000) if is_active else 0
                holding_cost = rng.integers(15000, 30000) if is_active else 0
                direct_cost = rng.integers(2000, 8000) if is_active else 0
                assembly_value = part_a_cost + holding_cost + direct_cost
                
                data.append(["Description", "Costs", None, None])
                data.append(["Cost of Part A Consumed", format_currency(part_a_cost), None, None])
                data.append(["Part A Holding and Ordering Cost", format_currency(holding_cost), None, None])
                data.append(["Direct and Indirect Costs", format_currency(direct_cost), None, None])
                data.append(["Value of Assembly A Produced", format_currency(assembly_value), None, None])
                data.append([None, None, None, None])
            else:
                # Part A inventory
                data.append([None, "Part A", "$", "$ / Part A"])
                initial = rng.integers(3000, 8000) if is_active else 0
                inputs = rng.integers(10000, 20000) if is_active else 0
                consumed = rng.integers(12000, 18000) if is_active else 0
                final = max(0, initial + inputs - consumed)
                unit_cost = rng.uniform(2.5, 4.5)
                
                data.append(["Initial inventory", f"{initial:,}", format_currency(initial * unit_cost), f"${unit_cost:.2f}"])
                data.append(["Inputs", f"{inputs:,}", format_currency(inputs * unit_cost * 1.3), f"${unit_cost * 1.3:.2f}"])
                data.append(["Consumption", f"{consumed:,}", format_currency(consumed * unit_cost * 1.1), f"${unit_cost * 1.1:.2f}"])
                data.append(["Final inventory", f"{final:,}", format_currency(final * unit_cost * 1.2), f"${unit_cost * 1.2:.2f}"])
                
                # Assembly A inventory
                data.append([None, "Assembly A", "$", "$ / Assembly A"])
                init_a = rng.integers(6000, 12000) if is_active else 0
                prod_a = rng.integers(12000, 18000) if is_active else 0
                cons_a = rng.integers(11000, 16000) if is_active else 0
                final_a = max(0, init_a + prod_a - cons_a)
                unit_a = rng.uniform(4.0, 6.0)
                
                data.append(["Initial inventory", f"{init_a:,}", format_currency(init_a * unit_a), f"${unit_a:.2f}"])
                data.append(["Production", f"{prod_a:,}", format_currency(prod_a * unit_a * 1.2), f"${unit_a * 1.2:.2f}"])
                data.append(["Consumption", f"{cons_a:,}", format_currency(cons_a * unit_a * 1.05), f"${unit_a * 1.05:.2f}"])
                data.append(["Final inventory", f"{final_a:,}", format_currency(final_a * unit_a * 1.15), f"${unit_a * 1.15:.2f}"])
                data.append([None, None, None, None])
        
        # --- SECTION 2: Piece 1 + Part B -> Assembly B ---
        # Block size: ~17 rows
        for _ in range(2):
            data.append(["Section 2", None, None, None])
            if _ == 0:
                piece_cost = rng.integers(5000, 15000) if is_active else 0
                part_b_cost = rng.integers(60000, 100000) if is_active else 0
                holding_b = rng.integers(30000, 70000) if is_active else 0
                direct_b = rng.integers(8000, 15000) if is_active else 0
                assembly_b = piece_cost + part_b_cost + holding_b + direct_b
                
                data.append(["Description", "Costs", None, None])
                data.append(["Cost of Piece 1 Consumed", format_currency(piece_cost), None, None])
                data.append(["Cost of Part B Consumed", format_currency(part_b_cost), None, None])
                data.append(["Part B Holding and Ordering Cost", format_currency(holding_b), None, None])
                data.append(["Direct and Indirect Costs", format_currency(direct_b), None, None])
                data.append(["Value of Assembly B Produced", format_currency(assembly_b), None, None])
                data.append([None, None, None, None])
            else:
                # Piece 1 inventory
                data.append([None, "Piece 1", "$", "$ / Piece 1"])
                p1_init = rng.integers(100, 300) if is_active else 0
                p1_in = rng.integers(30, 80) if is_active else 0
                p1_cons = rng.integers(100, 200) if is_active else 0
                p1_final = max(0, p1_init + p1_in - p1_cons)
                p1_cost = rng.uniform(55, 65)
                
                data.append(["Initial inventory", f"{p1_init:,}", format_currency(p1_init * p1_cost), f"${p1_cost:.2f}"])
                data.append(["Inputs", f"{p1_in:,}", format_currency(p1_in * p1_cost), f"${p1_cost:.2f}"])
                data.append(["Consumption", f"{p1_cons:,}", format_currency(p1_cons * p1_cost), f"${p1_cost:.2f}"])
                data.append(["Final inventory", f"{p1_final:,}", format_currency(p1_final * p1_cost), f"${p1_cost:.2f}"])
                
                # Part B inventory
                data.append([None, "Part B", "$", "$ / Part B"])
                pb_init = rng.integers(800, 1500) if is_active else 0
                pb_in = rng.integers(2000, 4000) if is_active else 0
                pb_cons = rng.integers(2500, 3500) if is_active else 0
                pb_final = max(0, pb_init + pb_in - pb_cons)
                pb_cost = rng.uniform(22, 32)
                
                data.append(["Initial inventory", f"{pb_init:,}", format_currency(pb_init * pb_cost), f"${pb_cost:.2f}"])
                data.append(["Inputs", f"{pb_in:,}", format_currency(pb_in * pb_cost * 1.4), f"${pb_cost * 1.4:.2f}"])
                data.append(["Consumption", f"{pb_cons:,}", format_currency(pb_cons * pb_cost * 1.2), f"${pb_cost * 1.2:.2f}"])
                data.append(["Final inventory", f"{pb_final:,}", format_currency(pb_final * pb_cost * 1.3), f"${pb_cost * 1.3:.2f}"])
                
                # Assembly B inventory
                data.append([None, "Assembly B", "$", "$ / Assembly B"])
                ab_init = rng.integers(4000, 8000) if is_active else 0
                ab_prod = rng.integers(12000, 18000) if is_active else 0
                ab_cons = rng.integers(11000, 16000) if is_active else 0
                ab_final = max(0, ab_init + ab_prod - ab_cons)
                ab_cost = rng.uniform(9, 12)
                
                data.append(["Initial inventory", f"{ab_init:,}", format_currency(ab_init * ab_cost), f"${ab_cost:.2f}"])
                data.append(["Production", f"{ab_prod:,}", format_currency(ab_prod * ab_cost * 1.1), f"${ab_cost * 1.1:.2f}"])
                data.append(["Consumption", f"{ab_cons:,}", format_currency(ab_cons * ab_cost * 0.95), f"${ab_cost * 0.95:.2f}"])
                data.append(["Final inventory", f"{ab_final:,}", format_currency(ab_final * ab_cost * 1.05), f"${ab_cost * 1.05:.2f}"])
                data.append([None, None, None, None])
        
        # --- SECTION 3: Assembly A + Assembly B -> Assembly C ---
        # Block size: ~17 rows
        for _ in range(2):
            data.append(["Section 3", None, None, None])
            if _ == 0:
                aa_cost = rng.integers(50000, 80000) if is_active else 0
                aa_hold = rng.integers(10000, 20000) if is_active else 0
                ab_cost2 = rng.integers(100000, 170000) if is_active else 0
                ab_hold = rng.integers(15000, 30000) if is_active else 0
                direct_c = rng.integers(30000, 60000) if is_active else 0
                assembly_c = aa_cost + aa_hold + ab_cost2 + ab_hold + direct_c
                
                data.append(["Description", "Costs", None, None])
                data.append(["Cost of Assembly A Consumed", format_currency(aa_cost), None, None])
                data.append(["Assembly A Holding Cost", format_currency(aa_hold), None, None])
                data.append(["Cost of Assembly B Consumed", format_currency(ab_cost2), None, None])
                data.append(["Assembly B Holding Cost", format_currency(ab_hold), None, None])
                data.append(["Direct and Indirect Costs", format_currency(direct_c), None, None])
                data.append(["Value of Assembly C Produced", format_currency(assembly_c), None, None])
                data.append([None, None, None, None])
            else:
                # Assembly A inventory (subsection)
                data.append([None, "Assembly A", "$", "$ / Assembly A"])
                for row_type, mult in [("Initial inventory", 1), ("Inputs", 1.2), ("Consumption", 1.05), ("Final inventory", 1.15)]:
                    val = rng.integers(5000, 15000) if is_active else 0
                    uc = rng.uniform(4, 6)
                    data.append([row_type, f"{val:,}", format_currency(val * uc * mult), f"${uc * mult:.2f}"])
                
                # Assembly B inventory (subsection)
                data.append([None, "Assembly B", "$", "$ / Assembly B"])
                for row_type, mult in [("Initial inventory", 1), ("Inputs", 1.1), ("Consumption", 0.95), ("Final inventory", 1.08)]:
                    val = rng.integers(4000, 10000) if is_active else 0
                    uc = rng.uniform(9, 12)
                    data.append([row_type, f"{val:,}", format_currency(val * uc * mult), f"${uc * mult:.2f}"])
                
                # Assembly C inventory
                data.append([None, "Assembly C", "$", "$ / Assembly C"])
                for row_type, mult in [("Initial inventory", 1), ("Production", 1.05), ("Consumption", 0.98), ("Final inventory", 1.02)]:
                    val = rng.integers(5000, 12000) if is_active else 0
                    uc = rng.uniform(18, 25)
                    data.append([row_type, f"{val:,}", format_currency(val * uc * mult), f"${uc * mult:.2f}"])
                data.append([None, None, None, None])
        
        # --- SECTION 4: Piece 2 + Assembly C -> Assembly D ---
        # Block size: ~32 rows
        for _ in range(2):
            data.append(["Section 4", None, None, None])
            if _ == 0:
                p2_cost = rng.integers(80000, 120000) if is_active else 0
                p2_hold = rng.integers(20000, 40000) if is_active else 0
                ac_cost = rng.integers(250000, 350000) if is_active else 0
                ac_hold = rng.integers(30000, 50000) if is_active else 0
                direct_d = rng.integers(40000, 80000) if is_active else 0
                assembly_d = p2_cost + p2_hold + ac_cost + ac_hold + direct_d
                
                data.append(["Description", "Costs", None, None])
                data.append(["Cost of Piece 2 Consumed", format_currency(p2_cost), None, None])
                data.append(["Piece 2 Holding Cost", format_currency(p2_hold), None, None])
                data.append(["Cost of Assembly C Consumed", format_currency(ac_cost), None, None])
                data.append(["Assembly C Holding Cost", format_currency(ac_hold), None, None])
                data.append(["Direct and Indirect Costs", format_currency(direct_d), None, None])
                data.append(["Value of Assembly D Produced", format_currency(assembly_d), None, None])
                data.append([None, None, None, None])
            else:
                # Piece 2 inventory
                data.append([None, "Piece 2", "$", "$ / Piece 2"])
                for row_type in ["Initial inventory", "Inputs", "Consumption", "Final inventory"]:
                    val = rng.integers(20000, 100000) if is_active else 0
                    uc = rng.uniform(0.08, 0.15)
                    data.append([row_type, f"{val:,}", format_currency(val * uc), f"${uc:.2f}"])
                
                # Assembly C inventory 
                data.append([None, "Assembly C", "$", "$ / Assembly C"])
                for row_type in ["Initial inventory", "Inputs", "Consumption", "Final inventory"]:
                    val = rng.integers(5000, 12000) if is_active else 0
                    uc = rng.uniform(20, 28)
                    data.append([row_type, f"{val:,}", format_currency(val * uc), f"${uc:.2f}"])
                
                # Assembly D inventory
                data.append([None, "Assembly D", "$", "$ / Assembly D"])
                for row_type in ["Initial inventory", "Production", "Consumption", "Final inventory"]:
                    val = rng.integers(3000, 10000) if is_active else 0
                    uc = rng.uniform(40, 60)
                    data.append([row_type, f"{val:,}", format_currency(val * uc), f"${uc:.2f}"])
                
                # Extra padding rows to match 32 row count
                data.append([None] * 4)
                data.append(["Scrap generated", "0", "$0.00", None])
                data.append(["Defective units", "0", "$0.00", None])
                data.append([None] * 4)
                data.append([None] * 4)

        # --- SECTION 5: Maintenance & Labor ---
        # Block size: ~22 rows (Analysis turned up 5 sections per zone)
        for _ in range(2):
           data.append(["Section 5", None, None, None]) 
           if _ == 0:
               # Cost summary
               maint_cost = rng.integers(10000, 20000) if is_active else 0
               labor_cost = rng.integers(30000, 50000) if is_active else 0
               total_sec5 = maint_cost + labor_cost
               
               data.append(["Description", "Costs", None, None])
               data.append(["Maintenance Costs", format_currency(maint_cost), None, None])
               data.append(["Indirect Labor", format_currency(labor_cost), None, None])
               data.append(["Total Section 5", format_currency(total_sec5), None, None])
               data.append([None, None, None, None])
           else:
               # Detailed metrics
               data.append([None, "Metric", "Value", "Unit"])
               data.append(["Efficiency", "Machine Uptime", "95%" if is_active else "0%", "Percentage"])
               data.append(["Quality", "Defect Rate", "0.5%" if is_active else "0%", "Percentage"])
               data.append(["Safety", "Accidents", "0", "Count"])
               data.append(["OEE", "Overall Equipment Effectiveness", "87%" if is_active else "0%", "Percentage"])
               
               data.append([None] * 4)
               data.append(["Labor Variance", "$0.00", None, None])
               data.append(["Material Variance", "$0.00", None, None])
               data.append(["Overhead Variance", "$0.00", None, None])
               data.append(["Volume Variance", "$0.00", None, None])
               data.append(["Mix Variance", "$0.00", None, None])
               data.append([None] * 4)
               
               # WIP Inventory subsection (~20 additional rows)
               data.append([None, "WIP Inventory", "$", "Units"])
               for stage in ["Stage 1 WIP", "Stage 2 WIP", "Stage 3 WIP", "Stage 4 WIP"]:
                   val = rng.integers(1000, 5000) if is_active else 0
                   cost = rng.uniform(5, 15)
                   data.append([stage, f"{val:,}", format_currency(val * cost), f"{val:,}"])
               data.append([None] * 4)
               
               # Quality Control subsection
               data.append([None, "Quality Control", "Value", "Status"])
               data.append(["Inspections Passed", str(rng.integers(95, 100)) if is_active else "0", None, "OK"])
               data.append(["Inspections Failed", str(rng.integers(0, 5)) if is_active else "0", None, "Review"])
               data.append(["Rework Items", str(rng.integers(0, 20)) if is_active else "0", None, "In Process"])
               data.append(["Scrap Rate", "0.3%" if is_active else "0%", None, "Target"])
               data.append([None] * 4)
               
               # Shift Performance
               data.append([None, "Shift Performance", "Output", "Efficiency"])
               data.append(["Shift 1 (Day)", str(rng.integers(4000, 6000)) if is_active else "0", "94%" if is_active else "0%", None])
               data.append(["Shift 2 (Evening)", str(rng.integers(3500, 5500)) if is_active else "0", "91%" if is_active else "0%", None])
               data.append(["Shift 3 (Night)", str(rng.integers(3000, 5000)) if is_active else "0", "88%" if is_active else "0%", None])
               data.append([None] * 4)
               
               # Downtime Analysis
               data.append([None, "Downtime Analysis", "Hours", "Cause"])
               data.append(["Planned Downtime", str(rng.uniform(2, 8)) if is_active else "0", None, "Maintenance"])
               data.append(["Unplanned Downtime", str(rng.uniform(0, 2)) if is_active else "0", None, "Breakdown"])
               data.append(["Changeover Time", str(rng.uniform(1, 4)) if is_active else "0", None, "Setup"])
               data.append([None] * 4)
               
               # Energy Consumption
               data.append([None, "Energy Consumption", "kWh", "Cost"])
               data.append(["Electricity", str(rng.integers(10000, 30000)) if is_active else "0", format_currency(rng.integers(1000, 3000)) if is_active else "$0", None])
               data.append(["Natural Gas", str(rng.integers(5000, 15000)) if is_active else "0", format_currency(rng.integers(500, 1500)) if is_active else "$0", None])
               data.append(["Water", str(rng.integers(1000, 5000)) if is_active else "0", format_currency(rng.integers(100, 500)) if is_active else "$0", None])
               data.append(["Compressed Air", str(rng.integers(2000, 8000)) if is_active else "0", format_currency(rng.integers(200, 800)) if is_active else "$0", None])
               data.append([None] * 4)
               
               # Equipment Status (~10 rows)
               data.append([None, "Equipment Status", "Capacity", "Utilization"])
               data.append(["Machine Line 1", str(rng.integers(10000, 20000)) if is_active else "0", f"{rng.integers(75, 95)}%" if is_active else "0%", None])
               data.append(["Machine Line 2", str(rng.integers(8000, 18000)) if is_active else "0", f"{rng.integers(70, 92)}%" if is_active else "0%", None])
               data.append(["Machine Line 3", str(rng.integers(6000, 15000)) if is_active else "0", f"{rng.integers(65, 90)}%" if is_active else "0%", None])
               data.append(["Packaging Unit", str(rng.integers(12000, 25000)) if is_active else "0", f"{rng.integers(80, 98)}%" if is_active else "0%", None])
               data.append([None] * 4)
               
               # Personnel Summary (~8 rows)
               data.append([None, "Personnel Summary", "Count", "Hours"])
               data.append(["Operators", str(rng.integers(20, 50)) if is_active else "0", str(rng.integers(300, 500)) if is_active else "0", None])
               data.append(["Technicians", str(rng.integers(5, 15)) if is_active else "0", str(rng.integers(80, 200)) if is_active else "0", None])
               data.append(["Supervisors", str(rng.integers(2, 6)) if is_active else "0", str(rng.integers(40, 80)) if is_active else "0", None])
               data.append(["Temps/Contract", str(rng.integers(0, 10)) if is_active else "0", str(rng.integers(0, 150)) if is_active else "0", None])
               data.append(["Overtime Hours", None, str(rng.integers(50, 200)) if is_active else "0", None])
               data.append([None] * 4)
               
               # Maintenance Schedule (~8 rows)
               data.append([None, "Maintenance Schedule", "Next Due", "Status"])
               data.append(["Preventive Maintenance", "Line 1", "Period 7", "Scheduled"])
               data.append(["Preventive Maintenance", "Line 2", "Period 8", "Scheduled"])
               data.append(["Calibration Check", "All Lines", "Period 6", "Completed"])
               data.append(["Safety Inspection", "All Areas", "Period 7", "Pending"])
               data.append([None] * 4)
               
               # Inventory Summary (~6 rows)
               data.append([None, "Inventory Summary", "Units", "Value"])
               data.append(["Raw Materials", str(rng.integers(10000, 50000)) if is_active else "0", format_currency(rng.integers(50000, 200000)) if is_active else "$0", None])
               data.append(["WIP", str(rng.integers(5000, 20000)) if is_active else "0", format_currency(rng.integers(100000, 300000)) if is_active else "$0", None])
               data.append(["Finished Goods", str(rng.integers(8000, 30000)) if is_active else "0", format_currency(rng.integers(200000, 500000)) if is_active else "$0", None])
               data.append([None] * 4)
               data.append([None] * 4)
               
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False, header=False)



def generate_raw_materials(rng, output_path):
    """Generate raw_materials.xlsx matching exact 422-row real format with 5 zones x 5 sections."""
    data = create_header_rows("Raw Materials")
    
    # Generate for ALL 5 zones (Center, West, North, East, South)
    # Active zones have data; East/South are mostly empty/zeros but structurally present
    for zone_idx, zone in enumerate(ZONES):
        is_active = zone_idx < 3
        
        # --- SECTION 1: Part A -> Assembly A ---
        # Block size: 9 rows (Header + 8 data rows)
        data.append([f"{zone} - Section 1", None] + [None] * 8)
        data.append(["Fortnight", " 1", " 2", " 3", " 4", " 5", " 6", " 7", " 8", "Total"])
        
        # Assembly A Produced
        prod_a = [rng.integers(1200, 1800) if is_active else 0 for _ in range(8)]
        data.append(["Assembly A Produced"] + [f"{p:,}" for p in prod_a] + [f"{sum(prod_a):,}"])
        
        # Part A inventory tracking
        data.append(["Part A"] + [None] * 9)
        pa_init = [rng.integers(2000, 5000) if is_active else 0]
        pa_recv = [rng.integers(0, 6000) if (is_active and fn in [2, 4, 7]) else 0 for fn in range(8)]
        for fn in range(8):
            if fn == 0: continue
            pa_init.append(max(0, pa_init[-1] + pa_recv[fn-1] - prod_a[fn-1]))
        pa_final = [max(0, pa_init[i] + pa_recv[i] - prod_a[i]) for i in range(8)]
        
        data.append(["Initial inventory"] + [f"{pa_init[i]:,}" for i in range(8)] + [None])
        data.append(["Received"] + [f"{pa_recv[i]:,}" for i in range(8)] + [f"{sum(pa_recv):,}"])
        data.append(["Consumed"] + [f"{prod_a[i]:,}" for i in range(8)] + [f"{sum(prod_a):,}"])
        data.append(["Final inventory"] + [str(pa_final[i]) for i in range(8)] + [None])
        # Space to next section is handled by next section header logic in real file (it's tight)
        
        # --- SECTION 2: Piece 1 + Part B -> Assembly B ---
        # Block size: 14 rows
        data.append([f"{zone} - Section 2", None] + [None] * 8)
        data.append(["Fortnight", " 1", " 2", " 3", " 4", " 5", " 6", " 7", " 8", "Total"])
        
        # Assembly B Produced
        prod_b = [rng.integers(1200, 1700) if is_active else 0 for _ in range(8)]
        data.append(["Assembly B Produced"] + [f"{p:,}" for p in prod_b] + [f"{sum(prod_b):,}"])
        
        # Piece 1 inventory
        data.append(["Piece 1"] + [None] * 9)
        p1_init = [rng.integers(140, 220) if is_active else 0]
        p1_recv = [rng.integers(0, 60) if (is_active and fn == 0) else 0 for fn in range(8)]
        p1_cons = [int(prod_b[i] * 0.01) for i in range(8)]
        for fn in range(8):
            if fn == 0: continue
            p1_init.append(max(0, p1_init[-1] + p1_recv[fn-1] - p1_cons[fn-1]))
        p1_final = [max(0, p1_init[i] + p1_recv[i] - p1_cons[i]) for i in range(8)]
        
        data.append(["Initial inventory"] + [f"{p1_init[i]:,}" for i in range(8)] + [None])
        data.append(["Received"] + [f"{p1_recv[i]:,}" for i in range(8)] + [f"{sum(p1_recv):,}"])
        data.append(["Consumed"] + [f"{p1_cons[i]:,}" for i in range(8)] + [f"{sum(p1_cons):,}"])
        data.append(["Final inventory"] + [str(p1_final[i]) for i in range(8)] + [None])
        
        # Part B inventory  
        data.append(["Part B"] + [None] * 9)
        pb_init = [rng.integers(700, 1300) if is_active else 0]
        pb_recv = [rng.integers(0, 2000) if (is_active and fn in [2, 5]) else 0 for fn in range(8)]
        pb_cons = [int(prod_b[i] * 0.2) for i in range(8)]
        for fn in range(8):
            if fn == 0: continue
            pb_init.append(max(0, pb_init[-1] + pb_recv[fn-1] - pb_cons[fn-1]))
        pb_final = [max(0, pb_init[i] + pb_recv[i] - pb_cons[i]) for i in range(8)]
        
        data.append(["Initial inventory"] + [f"{pb_init[i]:,}" for i in range(8)] + [None])
        data.append(["Received"] + [f"{pb_recv[i]:,}" for i in range(8)] + [f"{sum(pb_recv):,}"])
        data.append(["Consumed"] + [f"{pb_cons[i]:,}" for i in range(8)] + [f"{sum(pb_cons):,}"])
        data.append(["Final inventory"] + [str(pb_final[i]) for i in range(8)] + [None])
        
        # --- SECTION 3: Assembly A + Assembly B -> Assembly C ---
        # Block size: 16 rows
        data.append([f"{zone} - Section 3", None] + [None] * 8)
        data.append(["Fortnight", " 1", " 2", " 3", " 4", " 5", " 6", " 7", " 8", "Total"])
        
        prod_c = [rng.integers(1100, 1700) if is_active else 0 for _ in range(8)]
        data.append(["Assembly C Produced"] + [f"{p:,}" for p in prod_c] + [f"{sum(prod_c):,}"])
        
        # Assembly A tracking
        data.append(["Assembly A"] + [None] * 9)
        aa_init = [rng.integers(5000, 8000) if is_active else 0]
        aa_recv = prod_a # Received from Section 1
        for fn in range(8):
            if fn == 0: continue
            aa_init.append(max(0, aa_init[-1] + aa_recv[fn-1] - prod_c[fn-1]))
        aa_final = [max(0, aa_init[i] + aa_recv[i] - prod_c[i]) for i in range(8)]
        
        data.append(["Initial inventory"] + [f"{aa_init[i]:,}" for i in range(8)] + [None])
        data.append(["Received"] + [f"{aa_recv[i]:,}" for i in range(8)] + [f"{sum(aa_recv):,}"])
        data.append(["Consumed"] + [f"{prod_c[i]:,}" for i in range(8)] + [f"{sum(prod_c):,}"])
        data.append(["Final inventory"] + [str(aa_final[i]) for i in range(8)] + [None])
        
        # Assembly B tracking
        data.append(["Assembly B"] + [None] * 9)
        ab_init = [rng.integers(3500, 6000) if is_active else 0]
        ab_recv = prod_b
        for fn in range(8):
            if fn == 0: continue
            ab_init.append(max(0, ab_init[-1] + ab_recv[fn-1] - prod_c[fn-1]))
        ab_final = [max(0, ab_init[i] + ab_recv[i] - prod_c[i]) for i in range(8)]
        
        data.append(["Initial inventory"] + [f"{ab_init[i]:,}" for i in range(8)] + [None])
        data.append(["Received"] + [f"{ab_recv[i]:,}" for i in range(8)] + [f"{sum(ab_recv):,}"])
        data.append(["Consumed"] + [f"{prod_c[i]:,}" for i in range(8)] + [f"{sum(prod_c):,}"])
        data.append(["Final inventory"] + [str(ab_final[i]) for i in range(8)] + [None])
        
        # --- SECTION 4: Piece 2 + Assembly C -> Assembly D ---
        # Block size: 29 rows (Large block!)
        data.append([f"{zone} - Section 4", None] + [None] * 8)
        data.append(["Fortnight", " 1", " 2", " 3", " 4", " 5", " 6", " 7", " 8", "Total"])
        
        prod_d = [rng.integers(1300, 1900) if is_active else 0 for _ in range(8)]
        data.append(["Assembly D Produced"] + [f"{p:,}" for p in prod_d] + [f"{sum(prod_d):,}"])
        
        # Piece 2 inventory
        data.append(["Piece 2"] + [None] * 9)
        p2_init = [rng.integers(30000, 100000) if is_active else 0]
        p2_recv = [rng.integers(0, 70000) if (is_active and fn == 0) else 0 for fn in range(8)]
        p2_cons = [int(prod_d[i] * 8) for i in range(8)]
        for fn in range(8):
            if fn == 0: continue
            p2_init.append(max(0, p2_init[-1] + p2_recv[fn-1] - p2_cons[fn-1]))
        p2_final = [max(0, p2_init[i] + p2_recv[i] - p2_cons[i]) for i in range(8)]
        
        data.append(["Initial inventory"] + [f"{p2_init[i]:,}" for i in range(8)] + [None])
        data.append(["Received"] + [f"{p2_recv[i]:,}" for i in range(8)] + [f"{sum(p2_recv):,}"])
        data.append(["Consumed"] + [f"{p2_cons[i]:,}" for i in range(8)] + [f"{sum(p2_cons):,}"])
        data.append(["Final inventory"] + [str(p2_final[i]) for i in range(8)] + [None])
        
        # Assembly C tracking
        data.append(["Assembly C"] + [None] * 9)
        ac_init = [rng.integers(4000, 8000) if is_active else 0]
        ac_recv = prod_c
        for fn in range(8):
            if fn == 0: continue
            ac_init.append(max(0, ac_init[-1] + ac_recv[fn-1] - prod_d[fn-1]))
        ac_final = [max(0, ac_init[i] + ac_recv[i] - prod_d[i]) for i in range(8)]
        
        data.append(["Initial inventory"] + [f"{ac_init[i]:,}" for i in range(8)] + [None])
        # Extra spacing row in real file logic
        data.append(["Received"] + [f"{ac_recv[i]:,}" for i in range(8)] + [f"{sum(ac_recv):,}"])
        data.append(["Consumed"] + [f"{prod_d[i]:,}" for i in range(8)] + [f"{sum(prod_d):,}"])
        data.append(["Final inventory"] + [str(ac_final[i]) for i in range(8)] + [None])
        data.append(["In-transit from previous section"] + [f"{prod_c[i]:,}" for i in range(8)] + [f"{sum(prod_c):,}"])
        
        # Packaging Material subsection (filler to match row count)
        data.append([None] * 10)
        data.append(["Packaging Material"] + [None] * 9)
        data.append(["Initial inventory"] + ["0"]*8 + [None])
        data.append(["Received"] + ["0"]*8 + [None])
        data.append(["Consumed"] + ["0"]*8 + [None])
        data.append(["Final inventory"] + ["0"]*8 + [None])
        data.append([None] * 10)
        data.append([None] * 10) 
        
        # --- SECTION 5: Scraps / Recycling ---
        # Block size: 16 rows to next zone
        data.append([f"{zone} - Section 5", None] + [None] * 8)
        data.append(["Fortnight", " 1", " 2", " 3", " 4", " 5", " 6", " 7", " 8", "Total"])
        data.append(["Scrap Material Generated"] + ["0"]*8 + ["0"])
        data.append([None] * 10)
        
        # Recycling subsection
        data.append(["Recycling"] + [None] * 9)
        data.append(["Initial inventory"] + ["0"]*8 + [None])
        data.append(["Received"] + ["0"]*8 + [None])
        data.append(["Consumed"] + ["0"]*8 + [None])
        data.append(["Final inventory"] + ["0"]*8 + [None])
        data.append([None] * 10)
        
        # Waste Management subsection (~8 rows)
        data.append(["Hazardous Waste"] + [None] * 9)
        data.append(["Initial inventory"] + ["0"]*8 + [None])
        data.append(["Generated"] + ["0"]*8 + [None])
        data.append(["Disposed"] + ["0"]*8 + [None])
        data.append(["Final inventory"] + ["0"]*8 + [None])
        data.append([None] * 10)
        
        # Quality Metrics subsection (~6 rows)
        data.append(["Quality Metrics"] + [None] * 9)
        data.append(["Defect Rate"] + ["0%"]*8 + ["0%"])
        data.append(["Returns"] + ["0"]*8 + ["0"])
        data.append(["Rework"] + ["0"]*8 + ["0"])
        data.append([None] * 10)
        
        # Supplier Performance subsection (~6 rows)
        data.append(["Supplier Performance"] + [None] * 9)
        data.append(["On-Time Delivery"] + ["95%"]*8 + ["95%"])
        data.append(["Quality Rating"] + ["A"]*8 + ["A"])
        data.append(["Lead Time Variance"] + ["0"]*8 + ["0"])
        data.append([None] * 10)
        data.append([None] * 10)
        data.append([None] * 10)

    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False, header=False)


def generate_machine_spaces(rng, output_path):
    """Generate machine_spaces.xlsx matching exact 125-row real format."""
    data = create_header_rows("Machines and Spaces")
    
    # Generate data for each zone (Center, West, North for 3 active + East/South empty)
    for zone_idx, zone in enumerate(ZONES):
        is_active = zone_idx < 3  # Only Center, West, North are active
        
        # Machine Control section
        data.append(["                 Machine Control            ", None] + [None] * 11)
        data.append(["Machine type", "Acquisition Cost", "Purchase period", "Sale period",
                     "Depreciation period", "Period Amortization / Leasing",
                     "Amortization amount left", "Initial Machines", "Machines Sent",
                     "Machines Received", "Machines Available", None, None])
        
        # Machine data with totals
        machines_data = []
        if is_active:
            machines_data = [
                ("M1", rng.integers(3, 10), rng.integers(8000, 12000)),
                ("M2", rng.integers(8, 25), rng.integers(5000, 8000)),
                ("M3-alpha", rng.integers(1, 5), rng.integers(80000, 100000)),
                ("M3-beta", 0, 0),
                ("M4", rng.integers(8, 30), rng.integers(2000, 3000)),
            ]
        else:
            machines_data = [
                ("M1", 0, 0), ("M2", 0, 0), ("M3-alpha", 0, 0), 
                ("M3-beta", 0, 0), ("M4", 0, 0)
            ]
        
        for m_type, machines, per_machine_cost in machines_data:
            acq_cost = machines * per_machine_cost
            deprec_period = "6/10" if m_type in ["M1", "M2"] else ("6/20" if "M3" in m_type else "6/8")
            period_amort = acq_cost // 10 if m_type in ["M1", "M2"] else (acq_cost // 20 if "M3" in m_type else acq_cost // 8)
            amort_left = period_amort * 4
            
            data.append([m_type, acq_cost, "1" if machines > 0 else None, "-", 
                        deprec_period if machines > 0 else None, period_amort, amort_left,
                        machines, 0, 0, machines, None, None])
            data.append([f"Total {m_type}", acq_cost, None, None, None, period_amort, amort_left,
                        machines, 0, 0, machines, None, None])
        
        data.append([None] * 13)
        
        # Machines Summary section
        data.append(["                 Machines Summary            ", None] + [None] * 11)
        data.append([" ", "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9", "P10", "P11", "P12"])
        
        for m_type, machines, _ in machines_data:
            if is_active and machines > 0:
                data.append([m_type] + [machines] * 6 + [None] * 6)
            else:
                data.append([m_type] + [None] * 12)
        
        data.append([None] * 13)
        
        # Spaces section
        data.append(["                 Spaces            ", None] + [None] * 11)
        data.append([" ", "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9", "P10", "P11", "P12"])
        
        if is_active:
            occupied = rng.integers(50, 80)
            available = rng.integers(60, 90)
            leased = 0
            data.append(["Occupied"] + [occupied] * 6 + [0] * 6)
            data.append(["Available"] + [available] * 6 + [0] * 6)
            data.append(["Leased"] + [leased] * 6 + [0] * 6)
            data.append(["Total"] + [available] * 6 + [0] * 6)
        else:
            data.append(["Occupied"] + [0] * 12)
            data.append(["Available"] + [0] * 12)
            data.append(["Leased"] + [0] * 12)
            data.append(["Total"] + [0] * 12)
        
        data.append([None] * 13)
    
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False, header=False)


def generate_accounts_receivable(rng, output_path):
    """Generate accounts_receivable_payable.xlsx"""
    data = create_header_rows("Accounts Receivable And Payable")
    data.append(["Accounts Receivable And Payable", None, None])
    data.append(["Fortnight", "Receivables", "Payables"])
    
    for fn in range(1, FORTNIGHTS + 1):
        receivable = rng.integers(0, 500000) if fn in [2, 4, 6, 8] else 0
        payable = rng.integers(0, 100000) if fn in [2, 3, 5, 6] else 0
        data.append([fn, format_currency(receivable), format_currency(payable)])
    
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False, header=False)


def generate_logistics(rng, output_path):
    """Generate logistics.xlsx matching exact real format."""
    data = create_header_rows("Logistics")
    # Row 3: Transportation Costs header
    data.append(["Transportation Costs", None, None, None, None, None])
    # Row 4: Column headers
    data.append(["Type", "Units", "Cost", "Total", None, None])
    
    # Row 5-6: Transport type subtotals
    data.append(["Subtotal Airplane", 0, 0.00, 0])
    data.append(["Subtotal Truck", 0, 0.00, 0])
    
    # Train routes
    routes = [
        ("Train (Center - West)", rng.integers(100, 1000), rng.uniform(10, 20)),
        ("Train (Center - North)", rng.integers(100, 1000), rng.uniform(12, 25)),
        ("Train (Center - North)", rng.integers(100, 2000), rng.uniform(10, 18)),  # Second route
    ]
    
    train_total = 0
    train_units = 0
    for route, units, cost in routes:
        total = units * cost
        data.append([route, units, round(cost, 2), f"{total:,.0f}".replace(',', ',')])
        train_total += total
        train_units += units
    
    avg_cost = train_total / train_units if train_units > 0 else 0
    data.append(["Subtotal Train", train_units, round(avg_cost, 2), f"{train_total:,.0f}".replace(',', ',')])
    data.append(["Total", train_units, round(avg_cost, 2), f"{train_total:,.0f}".replace(',', ',')])
    # Row 12: Empty
    data.append([None, None, None, None, None, None])
    
    # Row 13: Zone summary header
    data.append(["Incoming and Outcoming by Zone", None, None, None, None, None])
    # Row 14: Zone column headers
    data.append(["Zone", "Received Units", "In transit", "Sent Units", "Shipping Costs", "Warehouse Costs"])
    
    # Rows 15-19: Zone data
    zone_totals = {"received": 0, "sent": 0, "shipping": 0, "warehouse": 0}
    for zone in ZONES:
        received = rng.integers(0, 4000) if zone != "Center" else 0
        sent = rng.integers(1000, 5000) if zone == "Center" else 0
        shipping = int(sent * rng.uniform(10, 20)) if sent > 0 else 0
        warehouse = rng.integers(10000, 50000)
        data.append([zone, received, 0, sent, shipping, warehouse])
        zone_totals["received"] += received
        zone_totals["sent"] += sent
        zone_totals["shipping"] += shipping
        zone_totals["warehouse"] += warehouse
    
    # Row 20: Total row (matches real format)
    data.append(["Total", zone_totals["received"], 0, zone_totals["sent"], 
                 zone_totals["shipping"], zone_totals["warehouse"]])
    
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False, header=False)


def generate_financial_statements(rng, output_path):
    """Generate results_and_balance_statements.xlsx matching exact 68-row real format."""
    net_sales = rng.integers(800000, 1500000)
    cogs = int(net_sales * rng.uniform(0.35, 0.45))
    gross = net_sales - cogs
    
    # All expense line items matching real format
    warehouse = rng.integers(50000, 100000)
    freight = rng.integers(40000, 80000)
    machine_transfer = 0
    worker_hire = rng.integers(0, 5000)
    machine_rental = 0
    social = 0
    sa_expense = rng.integers(200000, 400000)
    energy = rng.integers(20000, 40000)
    co2 = 0
    assets_disposal = 0
    
    total_expenses = warehouse + freight + machine_transfer + worker_hire + machine_rental + social + sa_expense + energy + co2 + assets_disposal
    ebitda = gross - total_expenses
    depreciation = rng.integers(50000, 80000)
    operating_income = ebitda - depreciation
    
    interest_credit = rng.integers(20000, 35000)
    interest_mortgage = rng.integers(25000, 40000)
    interest_invest = rng.integers(3000, 8000)
    financial_expenses = interest_credit + interest_mortgage - interest_invest
    
    net_profit_before_tax = operating_income - financial_expenses
    taxes = max(0, int(net_profit_before_tax * 0.5))
    net_profit = net_profit_before_tax - taxes
    
    data = create_header_rows("Financial Statements")
    
    # Income Statement section (matching exact format)
    data.append(["                     Income Statement                ", None])
    data.append([None, None])
    data.append(["Net Sales", format_currency(net_sales)])
    data.append(["Cost of Goods Sold", format_currency(cogs)])
    data.append(["Gross Income", format_currency(gross)])
    data.append(["Warehouse Rental Expenses", format_currency(warehouse)])
    data.append(["Freight Expenses", format_currency(freight)])
    data.append(["Machine Transfer Expenses", format_currency(machine_transfer)])
    data.append(["Worker Hiring and Dismissal Expenses", format_currency(worker_hire)])
    data.append(["Machine Rental", format_currency(machine_rental)])
    data.append(["Social Expenses", format_currency(social)])
    data.append(["Sales and Administration Expenses", format_currency(sa_expense)])
    data.append(["Energy costs", format_currency(energy)])
    data.append(["CO2 Abatement Cost", format_currency(co2)])
    data.append(["Assets Disposal", format_currency(assets_disposal)])
    data.append(["EBITDA", format_currency(ebitda)])
    data.append(["Depreciation Of Plant and Equipment", format_currency(depreciation)])
    data.append(["Product Improvements", "$0"])
    data.append(["Depreciation Of ESG Assets", "$0"])
    data.append(["Operating Income", format_currency(operating_income)])
    data.append(["Interest on Lines of Credit", format_currency(interest_credit)])
    data.append(["Interest on Mortgage", format_currency(interest_mortgage)])
    data.append(["Interest on Emergency Loans", "$0"])
    data.append(["Interest on Investments", format_currency(interest_invest)])
    data.append(["Financial Expenses", format_currency(financial_expenses)])
    data.append(["Extraordinary Income and Expenses", "$0"])
    data.append(["Net Profit Before Taxes", format_currency(net_profit_before_tax)])
    data.append(["Taxes Payables (accrued in last period)", format_currency(taxes)])
    data.append(["Net Profit", format_currency(net_profit)])
    data.append(["Dividends", "$0"])
    data.append([None, None])
    
    # Assets section
    cash = rng.integers(150000, 300000)
    short_invest = rng.integers(100000, 250000)
    receivables = rng.integers(200000, 400000)
    raw_mat_inv = rng.integers(50000, 100000)
    wip_inv = rng.integers(200000, 400000)
    finished_inv = rng.integers(100000, 200000)
    inventory = raw_mat_inv + wip_inv + finished_inv
    current_assets = cash + short_invest + receivables + inventory
    
    plant_equip = rng.integers(800000, 1200000)
    esg_assets = 0
    intangible = 0
    accum_deprec = rng.integers(300000, 500000)
    accum_amort = 0
    tax_deferred = 0
    fixed_assets = plant_equip + esg_assets + intangible - accum_deprec - accum_amort
    total_assets = current_assets + fixed_assets
    
    data.append(["                     Assets                ", None])
    data.append([None, None])
    data.append(["Current Assets", format_currency(current_assets)])
    data.append(["Cash", format_currency(cash)])
    data.append(["Short-term Investment", format_currency(short_invest)])
    data.append(["Accounts Receivable", format_currency(receivables)])
    data.append(["Inventory", format_currency(inventory)])
    data.append(["\xa0\xa0Raw Materials", f"\xa0\xa0{format_currency(raw_mat_inv)}"])
    data.append(["\xa0\xa0Work in Progress", f"\xa0\xa0{format_currency(wip_inv)}"])
    data.append(["\xa0\xa0Finished Products", f"\xa0\xa0{format_currency(finished_inv)}"])
    data.append(["Fixed Assets", format_currency(fixed_assets)])
    data.append(["Plant and Equipment", format_currency(plant_equip)])
    data.append(["ESG Assets", format_currency(esg_assets)])
    data.append(["Intangible Assets", format_currency(intangible)])
    data.append(["Accumulated Depreciation of Plant and Equipment", format_currency(accum_deprec)])
    data.append(["Accumulated Amortization of ESG and Intangible Assets", format_currency(accum_amort)])
    data.append(["Tax Deferred Assets", format_currency(tax_deferred)])
    data.append(["Assets", format_currency(total_assets)])
    data.append([None, None])
    
    # Liabilities & Equity section
    payables = rng.integers(100000, 200000)
    line_credit = rng.integers(80000, 150000)
    interest_payable = rng.integers(20000, 40000)
    emergency_loan = 0
    mortgage = rng.integers(400000, 600000)
    dividends_payable = 0
    liabilities = payables + line_credit + interest_payable + emergency_loan + mortgage + taxes + dividends_payable
    
    equity = total_assets - liabilities
    retained = rng.integers(150000, 250000)
    issued_capital = rng.integers(700000, 900000)
    
    data.append(["                     Liabilities & Equity                ", None])
    data.append([None, None])
    data.append(["Liabilities", format_currency(liabilities)])
    data.append(["Accounts Payable", format_currency(payables)])
    data.append(["Line of Credit", format_currency(line_credit)])
    data.append(["Interests Payable", format_currency(interest_payable)])
    data.append(["Emergency Loan", format_currency(emergency_loan)])
    data.append(["Mortgage Loans", format_currency(mortgage)])
    data.append(["Taxes to Pay", format_currency(taxes)])
    data.append(["Dividends Payable", format_currency(dividends_payable)])
    data.append(["Shareholders Equity", format_currency(equity)])
    data.append(["Retained Earnings", format_currency(retained)])
    data.append(["Period profit (Loss)", format_currency(net_profit)])
    data.append(["Issued Capital", format_currency(issued_capital)])
    data.append(["Liabilities & Shareholders Equity", format_currency(total_assets)])
    
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False, header=False)
    return {"depreciation": depreciation, "retained_earnings": retained}


def generate_market_report(rng, output_path):
    """Generate market-report.xlsx matching exact real format with all sections."""
    data = create_header_rows("Market Report")
    
    companies = ["Company 1 A", "Company 2 A", "Company 3 A", "Company 4 A"]
    companies_base = ["Company 1", "Company 2", "Company 3", "Company 4"]
    
    # Section 1: Market Share Per Region (%)
    data.append(["       Market Share Per Region (%)    ", None, None, None, None, None])
    data.append(["Zone"] + companies[:4] + [None])
    for zone in ZONES:
        share = round(rng.uniform(20, 30), 1) if zone in ["Center", "West", "North"] else 0.0
        data.append([zone, str(share), str(share), str(share), str(share), None])
    data.append([None] * 6)
    
    # Section 2: Market Share Per Region Per Segment (%)
    data.append(["       Market Share Per Region Per Segment (%)    ", None, None, None, None, None])
    data.append(["Zone", "Segment"] + companies[:4])
    for zone in ZONES:
        share_high = round(rng.uniform(20, 30), 1) if zone in ["Center", "West", "North"] else 0.0
        share_low = round(rng.uniform(20, 30), 1) if zone in ["Center", "West", "North"] else 0.0
        data.append([zone, "High", str(share_high), str(share_high), str(share_high), str(share_high)])
        data.append([None, "Low", str(share_low), str(share_low), str(share_low), str(share_low)])
    data.append([None] * 6)
    
    # Section 3: Price
    data.append(["       Price    ", None, None, None, None, None])
    data.append(["Zone"] + companies[:4] + [None])
    for zone in ZONES:
        price = round(rng.uniform(60, 100), 2) if zone in ["Center", "West", "North"] else 0.00
        data.append([zone, f"{price:.2f}", f"{price:.2f}", f"{price:.2f}", f"{price:.2f}", None])
    data.append([None] * 6)
    
    # Section 4: Product Improvements
    data.append(["       Product Improvements    ", None, None, None, None, None])
    data.append(["Improvements", "Company 1 - A", "Company 2 - A", "Company 3 - A", "Company 4 - A", None])
    improvements = [
        "STAINLESS MATERIAL", "RECYCLABLE MATERIALS", "ENERGY EFFICIENCY",
        "LIGHTER AND MORE COMPACT", "IMPACT RESISTANCE", "NOISE REDUCTION",
        "IMPROVED BATTERY CAPACITY", "SELF-CLEANING", "SPEED SETTINGS",
        "DIGITAL CONTROLS", "VOICE ASSISTANCE INTEGRATION",
        "AUTOMATION AND PROGRAMMABILITY", "MULTIFUNCTIONAL ACCESSORIES",
        "MAPPING TECHNOLOGY"
    ]
    for imp in improvements:
        # Use zero-width space like real file
        data.append([imp, "\u200b", "\u200b", "\u200b", "\u200b", None])
    data.append([None] * 6)
    
    # Section 5: Product Awareness Percentage Per Segment
    data.append(["       Product Awareness Percentage Per Segment    ", None, None, None, None, None])
    data.append(["Zone", "Segment"] + companies[:4])
    for zone in ZONES:
        aware_high = round(rng.uniform(30, 50), 2) if zone in ["Center", "West", "North"] else 0.00
        aware_low = round(rng.uniform(20, 35), 2) if zone in ["Center", "West", "North"] else 0.00
        data.append([zone, "High", str(aware_high), str(aware_high), str(aware_high), str(aware_high)])
        data.append([None, "Low", str(aware_low), str(aware_low), str(aware_low), str(aware_low)])
    data.append([None] * 6)
    
    # Section 6: Product attractiveness (Perceived)
    data.append(["       Product attractiveness (Perceived)    ", None, None, None, None, None])
    data.append(["Zone", "Segment"] + companies[:4])
    for zone in ZONES:
        attr = round(rng.uniform(15, 25), 2) if zone in ["Center", "West", "North"] else 0.00
        data.append([zone, "High", str(attr), str(attr), str(attr), str(attr)])
        data.append([None, "Low", str(attr), str(attr), str(attr), str(attr)])
    data.append([None] * 6)
    
    # Section 7: Evaluation of the Promotional Impact of Salesforce
    data.append(["       Evaluation of the Promotional Impact of Salesforce    ", None, None, None, None, None])
    data.append(["Zone"] + companies_base + [None])
    for zone in ZONES:
        impact = round(rng.uniform(75, 100), 2) if zone in ["Center", "West", "North"] else 0.00
        data.append([zone, str(impact), str(impact), str(impact), str(impact), None])
    
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False, header=False)


def generate_esg(rng, output_path):
    """Generate ESG.xlsx matching exact real format with full sections."""
    data = create_header_rows("ESG Report")
    
    # Header row with periods P1-P12
    periods = [f"P{i}" for i in range(1, 13)]
    companies = ["Company 1", "Company 2", "Company 3", "Company 4"]
    
    # ESG Actions section
    data.append(["ESG Actions", None] + [None] * 11)
    data.append([None] + periods)
    
    actions = [
        "Trees Planting",
        "Green Electricity",
        "Solar Energy",
        "CO2 Credits Purchased",
    ]
    for action in actions:
        data.append([action] + [0] * 6 + [0] * 6)  # Values for P1-P12
    
    data.append([None] + [None] * 12)
    
    # Costs section
    data.append(["ESG Costs", None] + [None] * 11)
    data.append([None] + periods)
    data.append(["Trees Planting"] + [0] * 12)
    data.append(["Green Electricity"] + [0] * 12)
    data.append(["Solar Energy"] + [0] * 12)
    data.append(["CO2 Credits Purchased"] + [0] * 12)
    data.append([None] + [None] * 12)
    
    # Emissions Breakdown section
    data.append(["Emissions Breakdown [%]", None] + [None] * 11)
    data.append([None] + periods)
    
    emissions = [
        ("Suppliers' purchases Parts A and B", [rng.uniform(15, 35) for _ in range(6)] + [0] * 6),
        ("Electricity emissions", [rng.uniform(20, 40) for _ in range(6)] + [0] * 6),
        ("Product Improvements", [0] * 12),
        ("Shipping Electrocleans Across Regions", [rng.uniform(1, 5) for _ in range(6)] + [0] * 6),
        ("New Plant Modules", [rng.uniform(35, 55) for _ in range(6)] + [0] * 6),
        ("Thrown Electrocleans", [0] * 12),
    ]
    for name, values in emissions:
        data.append([name] + [f"{v:.1f}%" if v > 0 else "0.0%" for v in values])
    
    data.append([None] + [None] * 12)
    
    # Emissions intensity section
    data.append(["Emissions intensity (kg CO2 / Electroclean)", None] + [None] * 11)
    data.append(["Period"] + companies + [None] * 8)
    for period in [6, 7, 8, 9, 10, 11, 12]:
        val = round(rng.uniform(25, 35), 2) if period == 6 else 0.00
        data.append([period, val, val, val, val] + [None] * 8)
    
    data.append([None] + [None] * 12)
    
    # ENV-KPI section
    data.append(["ENV-KPI", None] + [None] * 11)
    data.append(["Period"] + companies + [None] * 8)
    for period in [6, 7, 8, 9, 10, 11, 12]:
        val = rng.integers(75, 100) if period == 6 else 0
        data.append([period, val, val, val, val] + [None] * 8)
    
    data.append([None] + [None] * 12)
    
    # ESG KPI section
    data.append(["ESG KPI", None] + [None] * 11)
    data.append(["Period"] + companies + [None] * 8)
    for period in [6, 7, 8, 9, 10, 11, 12]:
        val = round(rng.uniform(85, 95), 1) if period == 6 else 0.0
        data.append([period, val, val, val, val] + [None] * 8)
    
    data.append([None] + [None] * 12)
    
    # Electricity Offer section
    data.append(["Electricity Offer(kWh)", None] + [None] * 11)
    data.append([None] + periods)
    electricity = [
        ("Solar Energy", [0] * 12),
        ("Green Energy", [0] * 12),
        ("Grid Power", [rng.integers(200000, 450000) for _ in range(6)] + [0] * 6),
    ]
    for name, values in electricity:
        data.append([name] + values)
    
    # Total electricity
    total_elec = [sum(x) for x in zip(electricity[0][1], electricity[1][1], electricity[2][1])]
    data.append(["Total"] + total_elec)
    
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False, header=False)


def generate_subperiod_cash_flow(rng, output_path):
    """Generate subperiod_cash_flow.xlsx matching exact 45-row real format."""
    data = create_header_rows("Fortnight Cash Flow")
    
    # Section header
    data.append(["             Fortnight Cash Flow        ", None, None, None, None, None, None, None, None])
    # Fortnight columns
    data.append([None, " 1", " 2", " 3", " 4", " 5", " 6", " 7", " 8"])
    
    # Generate consistent cash flow
    opening_cash = rng.integers(40000, 80000)
    cash_per_fn = [opening_cash]
    
    # Calculate cash for each fortnight
    for fn in range(8):
        if fn > 0:
            cash_per_fn.append(cash_per_fn[-1] + rng.integers(-50000, 80000))
    
    # Cash at Beginning
    data.append(["Cash at the Beginning of the Fortnight"] + 
                [format_currency(cash_per_fn[i]) for i in range(8)])
    data.append([None] * 9)
    
    # Receipts/Payments
    data.append(["Payments from Customers"] + 
                [format_currency(rng.integers(0, 400000) if i % 2 == 1 else 0) for i in range(8)])
    data.append(["Payments to Suppliers"] + 
                [format_currency(-rng.integers(0, 30000)) for _ in range(8)])
    data.append([None] * 9)
    
    # Labor costs
    salary = rng.integers(5000, 12000)
    data.append(["Salaries"] + [format_currency(-salary) for _ in range(8)])
    data.append(["Overtime (blue-collar workers)"] + 
                [format_currency(-rng.integers(1000, 3000)) for _ in range(8)])
    data.append(["Indirect Labour Expenses"] + 
                [format_currency(-rng.integers(4000, 6000)) for _ in range(8)])
    data.append(["Labor Benefits"] + ["$0.00" for _ in range(8)])
    data.append(["Absenteeism"] + ["$0.00" for _ in range(8)])
    data.append([None] * 9)
    
    # Operations costs
    data.append(["Ordering Cost"] + 
                [format_currency(-rng.integers(0, 15000)) if i % 2 == 0 else "$0.00" for i in range(8)])
    data.append(["Inventory Holding Cost"] + 
                [format_currency(-rng.integers(12000, 20000)) for _ in range(8)])
    data.append(["Additional cost from improvements"] + ["$0" for _ in range(8)])
    data.append(["Energy costs"] + 
                [format_currency(-rng.integers(3000, 5000)) for _ in range(8)])
    data.append(["CO2 Abatement Cost"] + ["$0.00" for _ in range(8)])
    data.append([None] * 9)
    
    # S&A and Freight
    sa_expense = rng.integers(30000, 50000)
    data.append(["Sales and Administration Expenses"] + 
                [format_currency(-sa_expense) for _ in range(8)])
    data.append(["Freight"] + 
                [format_currency(-rng.integers(0, 20000)) if i < 4 else "$0.00" for i in range(8)])
    data.append([None] * 9)
    
    # Interest
    data.append(["Interest Earned on Short-Term Investment"] + 
                ["$0.00", "$0.00"] + [format_currency(rng.integers(500, 1500)) for _ in range(6)])
    data.append(["Interest of Line of Credit"] + 
                [format_currency(-rng.integers(1000, 6000)) for _ in range(8)])
    data.append(["Shareholder Loan Interest"] + ["$0.00" for _ in range(8)])
    data.append(["Cash Flow from Operations"] + 
                [format_currency(rng.integers(-150000, 250000)) for _ in range(8)])
    data.append([None] * 9)
    
    # Investments
    data.append(["Short-Term Investment Increase"] + 
                ["$0.00", "$0.00", format_currency(-rng.integers(100000, 200000))] + ["$0.00"] * 5)
    data.append(["Short-Term Investment Decrease"] + ["$0.00" for _ in range(8)])
    data.append(["Cash Flow from Investments"] + 
                ["$0", "$0", format_currency(-rng.integers(100000, 200000))] + ["$0"] * 5)
    data.append([None] * 9)
    
    # Financing
    data.append(["Line of Credit Increase"] + 
                [format_currency(rng.integers(0, 200000)) if i % 2 == 0 else "$0.00" for i in range(8)])
    data.append(["Repayment of Line of Credit"] + 
                ["$0" if i % 2 == 0 else format_currency(-rng.integers(0, 200000)) for i in range(8)])
    data.append(["Repayment of Emergency Loan"] + ["$0" for _ in range(8)])
    data.append(["Emergency Loan"] + ["$0.00" for _ in range(8)])
    data.append([None] * 9)
    
    data.append(["Cash Flow from Financing"] + 
                [format_currency(rng.integers(-100000, 150000)) for _ in range(8)])
    data.append([None] * 9)
    
    data.append(["Cash Variation"] + 
                [format_currency(rng.integers(-100000, 200000)) for _ in range(8)])
    data.append([None] * 9)
    
    data.append(["Cash at the End of the Fortnight"] + 
                [format_currency(cash_per_fn[i] + rng.integers(-20000, 50000)) for i in range(8)])
    data.append(["Emergency Credit Accrued Interest"] + [None] * 7 + ["$0.00"])
    
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False, header=False)


def generate_financial_statements_summary(rng, output_path):
    """Generate financial_statements_summary.xlsx - Same as results_and_balance but different format"""
    # Reuse the same generator since format is similar
    generate_financial_statements(rng, output_path)


# =============================================================================
# MAIN ORCHESTRATOR
# =============================================================================

def generate_all_mock_data(seed=DEFAULT_SEED, output_dir=DEFAULT_OUTPUT_DIR):
    """Generate all mock data files."""
    rng = np.random.default_rng(seed)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"Generating mock data with seed={seed} to {output_dir}")
    
    generators = [
        ("workers_balance_overtime.xlsx", generate_workers_balance),
        ("sales_admin_expenses.xlsx", generate_sales_admin),
        ("finished_goods_inventory.xlsx", generate_finished_goods),
        ("initial_cash_flow.xlsx", generate_initial_cash_flow),
        ("production.xlsx", generate_production),
        ("raw_materials.xlsx", generate_raw_materials),
        ("machine_spaces.xlsx", generate_machine_spaces),
        ("accounts_receivable_payable.xlsx", generate_accounts_receivable),
        ("logistics.xlsx", generate_logistics),
        ("results_and_balance_statements.xlsx", generate_financial_statements),
        ("financial_statements_summary.xlsx", generate_financial_statements_summary),
        ("market-report.xlsx", generate_market_report),
        ("ESG.xlsx", generate_esg),
        ("subperiod_cash_flow.xlsx", generate_subperiod_cash_flow),
    ]
    
    for filename, generator in generators:
        output_path = output_dir / filename
        try:
            generator(rng, output_path)
            print(f"  [OK] {filename}")
        except Exception as e:
            print(f"  [FAIL] {filename}: {e}")
    
    print(f"\nGenerated {len(generators)} mock data files.")
    return output_dir


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate mock ExSim report data")
    parser.add_argument("--seed", type=int, default=DEFAULT_SEED, help="Random seed for reproducibility")
    parser.add_argument("--output-dir", type=str, default=str(DEFAULT_OUTPUT_DIR), help="Output directory")
    args = parser.parse_args()
    
    generate_all_mock_data(seed=args.seed, output_dir=args.output_dir)
