"""
Mock Data Generator for CMO Dashboard

Creates realistic test data in the CMO Dashboard/data folder
that EXACTLY matches the format exported from the ExSim website.

Run: python mock_data_gen.py
"""

import pandas as pd
import numpy as np
from pathlib import Path
import random

# Configuration
OUTPUT_FOLDER = Path("data")
OUTPUT_FOLDER.mkdir(exist_ok=True)

ZONES = ["Center", "West", "North", "East", "South"]
SEGMENTS = ["High", "Low"]
COMPANIES = ["Company 1", "Company 2", "Company 3", "Company 4"]
MY_COMPANY = "Company 3"

# Seed for reproducibility (change seed for different data)
SEED = 42
np.random.seed(SEED)
random.seed(SEED)


def generate_market_report():
    """
    Generate market-report.xlsx in EXACT format as website export.
    
    Format:
    - Header rows (Country, Market Report, Period)
    - Market Share Per Region (%)
    - Market Share Per Region Per Segment (%)
    - Price
    - Product Improvements
    - Product Awareness Percentage Per Segment
    - Product attractiveness (Perceived)
    - Evaluation of the Promotional Impact of Salesforce
    """
    
    data_rows = []
    
    # ===== Header Section =====
    data_rows.append(["Country: IESE EMBA BCN 26-1", "", "", "", "", ""])
    data_rows.append(["Market Report", f"Company:{MY_COMPANY}", "", "", "", ""])
    data_rows.append(["Period: 6", "Course: IESE EMBA BCN 26-1", "", "", "", ""])
    
    # ===== Market Share Per Region (%) =====
    data_rows.append(["       Market Share Per Region (%)    ", "", "", "", "", ""])
    data_rows.append(["Zone", "Company 1 A", "Company 2 A", "Company 3 A", "Company 4 A", ""])
    
    for zone in ZONES:
        if zone in ["East", "South"]:  # Inactive zones
            shares = [0.0, 0.0, 0.0, 0.0]
        else:
            shares = [random.uniform(15, 35) for _ in range(4)]
            # Normalize to 100%
            total = sum(shares)
            shares = [round(s / total * 100, 1) for s in shares]
        data_rows.append([zone] + shares + [""])
    
    data_rows.append(["", "", "", "", "", ""])
    
    # ===== Market Share Per Region Per Segment (%) =====
    data_rows.append(["       Market Share Per Region Per Segment (%)    ", "", "", "", "", ""])
    data_rows.append(["Zone", "Segment", "Company 1 A", "Company 2 A", "Company 3 A", "Company 4 A"])
    
    for zone in ZONES:
        for i, segment in enumerate(SEGMENTS):
            if zone in ["East", "South"]:
                shares = [0.0, 0.0, 0.0, 0.0]
            else:
                shares = [random.uniform(15, 35) for _ in range(4)]
                total = sum(shares)
                shares = [round(s / total * 100, 1) for s in shares]
            
            if i == 0:  # First segment shows zone name
                data_rows.append([zone, segment] + shares)
            else:  # Second segment, zone is blank
                data_rows.append(["", segment] + shares)
    
    data_rows.append(["", "", "", "", "", ""])
    
    # ===== Price =====
    data_rows.append(["       Price    ", "", "", "", "", ""])
    data_rows.append(["Zone", "Company 1 A", "Company 2 A", "Company 3 A", "Company 4 A", ""])
    
    for zone in ZONES:
        if zone in ["East", "South"]:
            prices = [0.0, 0.0, 0.0, 0.0]
        else:
            # Random prices between 60-120
            prices = [round(random.uniform(60, 120), 2) for _ in range(4)]
        data_rows.append([zone] + prices + [""])
    
    data_rows.append(["", "", "", "", "", ""])
    
    # ===== Product Improvements =====
    data_rows.append(["       Product Improvements    ", "", "", "", "", ""])
    data_rows.append(["Improvements", "Company 1 - A", "Company 2 - A", "Company 3 - A", "Company 4 - A", ""])
    
    improvements = [
        "STAINLESS MATERIAL", "RECYCLABLE MATERIALS", "ENERGY EFFICIENCY",
        "LIGHTER AND MORE COMPACT", "IMPACT RESISTANCE", "NOISE REDUCTION",
        "IMPROVED BATTERY CAPACITY", "SELF-CLEANING", "SPEED SETTINGS",
        "DIGITAL CONTROLS", "VOICE ASSISTANCE INTEGRATION",
        "AUTOMATION AND PROGRAMMABILITY", "MULTIFUNCTIONAL ACCESSORIES",
        "MAPPING TECHNOLOGY"
    ]
    
    for improvement in improvements:
        # Random selection (​ = zero-width space = not selected, X = selected)
        selections = [random.choice(["​", "X"]) for _ in range(4)]
        data_rows.append([improvement] + selections + [""])
    
    data_rows.append(["", "", "", "", "", ""])
    
    # ===== Product Awareness Percentage Per Segment =====
    data_rows.append(["       Product Awareness Percentage Per Segment    ", "", "", "", "", ""])
    data_rows.append(["Zone", "Segment", "Company 1 A", "Company 2 A", "Company 3 A", "Company 4 A"])
    
    for zone in ZONES:
        for i, segment in enumerate(SEGMENTS):
            if zone in ["East", "South"]:
                awareness = [0.0, 0.0, 0.0, 0.0]
            else:
                # Higher awareness for High segment
                if segment == "High":
                    awareness = [round(random.uniform(25, 50), 2) for _ in range(4)]
                else:
                    awareness = [round(random.uniform(15, 35), 2) for _ in range(4)]
            
            if i == 0:
                data_rows.append([zone, segment] + awareness)
            else:
                data_rows.append(["", segment] + awareness)
    
    data_rows.append(["", "", "", "", "", ""])
    
    # ===== Product attractiveness (Perceived) =====
    data_rows.append(["       Product attractiveness (Perceived)    ", "", "", "", "", ""])
    data_rows.append(["Zone", "Segment", "Company 1 A", "Company 2 A", "Company 3 A", "Company 4 A"])
    
    for zone in ZONES:
        for i, segment in enumerate(SEGMENTS):
            if zone in ["East", "South"]:
                attract = [0.0, 0.0, 0.0, 0.0]
            else:
                attract = [round(random.uniform(15, 35), 2) for _ in range(4)]
            
            if i == 0:
                data_rows.append([zone, segment] + attract)
            else:
                data_rows.append(["", segment] + attract)
    
    data_rows.append(["", "", "", "", "", ""])
    
    # ===== Evaluation of the Promotional Impact of Salesforce =====
    data_rows.append(["       Evaluation of the Promotional Impact of Salesforce    ", "", "", "", "", ""])
    data_rows.append(["Zone", "Company 1", "Company 2", "Company 3", "Company 4", ""])
    
    for zone in ZONES:
        if zone in ["East", "South"]:
            impact = [0.0, 0.0, 0.0, 0.0]
        else:
            impact = [round(random.uniform(70, 100), 2) for _ in range(4)]
        data_rows.append([zone] + impact + [""])
    
    data_rows.append(["", "", "", "", "", ""])
    
    # Create DataFrame and save
    df = pd.DataFrame(data_rows)
    df.to_excel(OUTPUT_FOLDER / "market-report.xlsx", index=False, header=False)
    print(f"  Created market-report.xlsx (Website format)")


def generate_marketing_decisions():
    """Generate Marketing Decisions.xlsx template."""
    
    data_rows = []
    
    # Header
    data_rows.append(["MARKETING DECISIONS", "", "", ""])
    data_rows.append(["", "", "", ""])
    
    # TV Budget
    data_rows.append(["TV BUDGET", "", "", ""])
    data_rows.append(["Amount", random.randint(50000, 150000), "", ""])
    data_rows.append(["", "", "", ""])
    
    # Radio by zone
    data_rows.append(["RADIO BUDGET BY ZONE", "", "", ""])
    data_rows.append(["Zone", "Amount", "", ""])
    for zone in ZONES:
        data_rows.append([zone, random.randint(5000, 25000), "", ""])
    data_rows.append(["", "", "", ""])
    
    # Salespeople by zone
    data_rows.append(["SALESPEOPLE BY ZONE", "", "", ""])
    data_rows.append(["Zone", "Count", "Salary", ""])
    for zone in ZONES:
        data_rows.append([zone, random.randint(3, 12), random.randint(4000, 6000), ""])
    data_rows.append(["", "", "", ""])
    
    # Pricing by zone
    data_rows.append(["PRICING BY ZONE", "", "", ""])
    data_rows.append(["Zone", "High Segment Price", "Low Segment Price", ""])
    for zone in ZONES:
        high_price = random.randint(95, 130)
        low_price = random.randint(70, 95)
        data_rows.append([zone, high_price, low_price, ""])
    
    df = pd.DataFrame(data_rows)
    df.to_excel(OUTPUT_FOLDER / "Marketing Decisions.xlsx", index=False, header=False)
    print(f"  Created Marketing Decisions.xlsx")


def generate_innovation_decisions():
    """Generate Marketing Innovation Decisions.xlsx."""
    
    features = [
        ("Bluetooth Connectivity", 5000, 8),
        ("Premium Sound System", 8000, 12),
        ("GPS Navigation", 6000, 10),
        ("Leather Interior", 12000, 15),
        ("Sunroof", 4000, 6),
        ("Advanced Safety Package", 10000, 14),
        ("Sport Suspension", 7000, 11),
        ("Climate Control", 3000, 5)
    ]
    
    data_rows = []
    data_rows.append(["INNOVATION DECISIONS", "", "", ""])
    data_rows.append(["", "", "", ""])
    data_rows.append(["Feature", "Cost ($)", "Attractiveness Boost", "Selected"])
    
    for feature, cost, boost in features:
        selected = random.choice([0, 0, 1])
        data_rows.append([feature, cost, boost, selected])
    
    data_rows.append(["", "", "", ""])
    data_rows.append(["TOTAL SELECTED", "=SUMPRODUCT(B4:B11,D4:D11)", "", ""])
    
    df = pd.DataFrame(data_rows)
    df.to_excel(OUTPUT_FOLDER / "Marketing Innovation Decisions.xlsx", index=False, header=False)
    print(f"  Created Marketing Innovation Decisions.xlsx")


def generate_inventory():
    """
    Generate finished_goods_inventory.xlsx in EXACT website format.
    
    Format:
    - Header: Country, Finished Goods Inventory, Period
    - Per zone section with Capacity and subperiod columns (1-8)
    - Rows: Initial inventory, Receptions, Production, Shipments, Sales, Final inventory, Thrown away
    """
    
    data_rows = []
    
    # Header
    data_rows.append(["Country: IESE EMBA BCN 26-1", "", "", "", "", "", "", "", "", "", "", ""])
    data_rows.append(["Finished Goods Inventory", f"Company:{MY_COMPANY}", "", "", "", "", "", "", "", "", "", ""])
    data_rows.append(["Period: 6", "Course: IESE EMBA BCN 26-1", "", "", "", "", "", "", "", "", "", ""])
    
    zone_capacities = {
        "Center": 4800,
        "West": 2500,
        "North": 2000,
        "East": 0,
        "South": 0
    }
    
    for zone in ZONES:
        capacity = zone_capacities.get(zone, 0)
        
        # Section header
        data_rows.append(["                 Finished Goods Inventory            ", "", "", "", "", "", "", "", "", "", "", ""])
        data_rows.append([f"Capacity: {capacity}", "Previous", "1", "2", "3", "4", "5", "6", "7", "8", "Total", "In transit"])
        
        if capacity > 0:
            # Generate realistic inventory data
            initial_inv = [random.randint(1000, 5000) for _ in range(8)]
            receptions = [random.choice([0, 0, 0, random.randint(500, 3000)]) for _ in range(8)]
            production = [random.randint(300, 1500) for _ in range(8)] if zone in ["Center", "West"] else [0] * 8
            shipments = [random.choice([0, 0, random.randint(500, 2000)]) for _ in range(8)]
            sales = [random.choice([0, random.randint(1000, 2500)]) for _ in range(8)]
            final_inv = [max(0, initial_inv[i] + receptions[i] + production[i] - shipments[i] - sales[i]) for i in range(8)]
            thrown = [0] * 8
            
            prev_inv = random.randint(2000, 5000)
            
            # Format with thousands separator
            def fmt(val):
                return f"{val:,}" if val > 0 else "0"
            
            data_rows.append(["Initial inventory", ""] + [fmt(v) for v in initial_inv] + ["", ""])
            data_rows.append(["Receptions", ""] + [fmt(v) for v in receptions] + [fmt(sum(receptions)), "0"])
            data_rows.append(["Production", ""] + [fmt(v) for v in production] + [fmt(sum(production)), ""])
            data_rows.append(["Shipments", ""] + [fmt(v) for v in shipments] + [fmt(sum(shipments)), ""])
            data_rows.append(["Sales", ""] + [fmt(v) for v in sales] + [fmt(sum(sales)), ""])
            data_rows.append(["Final inventory", fmt(prev_inv)] + [fmt(v) for v in final_inv] + ["", ""])
            data_rows.append(["Thrown away", "0"] + ["0"] * 8 + ["0", ""])
        else:
            # Empty zone
            data_rows.append(["Initial inventory", ""] + ["0"] * 8 + ["", ""])
            data_rows.append(["Receptions", ""] + ["0"] * 8 + ["0", "0"])
            data_rows.append(["Production", ""] + ["0"] * 8 + ["0", ""])
            data_rows.append(["Shipments", ""] + ["0"] * 8 + ["0", ""])
            data_rows.append(["Sales", ""] + ["0"] * 8 + ["0", ""])
            data_rows.append(["Final inventory", "0"] + ["0"] * 8 + ["", ""])
            data_rows.append(["Thrown away", "0"] + ["0"] * 8 + ["0", ""])
        
        data_rows.append([""] * 12)
    
    df = pd.DataFrame(data_rows)
    df.to_excel(OUTPUT_FOLDER / "finished_goods_inventory.xlsx", index=False, header=False)
    print(f"  Created finished_goods_inventory.xlsx (Website format)")


def generate_sales_admin():
    """
    Generate sales_admin_expenses.xlsx in EXACT website format.
    
    Format:
    - Header: Country, Sales & Administration Expenses, Period
    - Sales section: Region, Brand, Units, Local Price, Gross Sales, Discount %, Net Sales
    - Expenses section: Category, Amount, Expense
    """
    
    data_rows = []
    
    # Header
    data_rows.append(["Country: IESE EMBA BCN 26-1", "", "", "", "", "", ""])
    data_rows.append(["Sales & Administration Expenses", f"Company:{MY_COMPANY}", "", "", "", "", ""])
    data_rows.append(["Period: 6", "Course: IESE EMBA BCN 26-1", "", "", "", "", ""])
    
    # Sales section
    data_rows.append(["                 Sales            ", "", "", "", "", "", ""])
    data_rows.append(["Region", "Brand", "Units", "Local Price", "Gross Sales", "Discount %", "Net Sales"])
    
    prices = {"Center": 68, "West": 68, "North": 91, "East": 0, "South": 0}
    total_units = 0
    total_gross = 0
    total_net = 0
    
    for zone in ["Center", "West", "North"]:
        units = random.randint(3000, 9000)
        price = prices[zone]
        gross = units * price
        discount = 7.5
        net = gross * (1 - discount/100)
        
        total_units += units
        total_gross += gross
        total_net += net
        
        data_rows.append([zone, "A", f"{units:,}", f"${price:.2f}", f"${gross:,}", discount, f"${net:,.0f}"])
    
    data_rows.append(["Total", "", f"{total_units:,}", "", f"${total_gross:,}", "", f"${total_net:,.0f}"])
    data_rows.append([""] * 7)
    
    # Expenses section
    data_rows.append(["             Sales & Administration Expenses        ", "", "", "", "", "", ""])
    data_rows.append(["", "Amount", "Expense", "", "", "", ""])
    
    salespeople = random.randint(30, 60)
    tv_spots = random.randint(20, 50)
    radio_spots = random.randint(200, 500)
    modules = random.randint(4, 10)
    
    expenses = [
        ("Salespeople Salaries", f"{salespeople} Salespeople", f"${salespeople * 750:,}"),
        ("Other channels costs", "-", "$0"),
        ("Salespeople Hiring Expenses", "0 Salespeople", "$0"),
        ("TV Advertising Expenses", f"{tv_spots} spots", f"${tv_spots * 3000:,}"),
        ("Radio Advertising Expenses", f"{radio_spots} spots", f"${radio_spots * 300:,}"),
        ("Plant Modules Leasing Expenses", "0 Modules", "$0"),
        ("Plant Module Administrative Expenses", f"{modules} Modules", f"${modules * 10000:,}"),
        ("Executive Salaries", "", "$0"),
    ]
    
    total_expense = salespeople * 750 + tv_spots * 3000 + radio_spots * 300 + modules * 10000
    
    for category, amount, expense in expenses:
        data_rows.append([category, amount, expense, "", "", "", ""])
    
    data_rows.append(["Total", "", f"${total_expense:,}", "", "", "", ""])
    
    df = pd.DataFrame(data_rows)
    df.to_excel(OUTPUT_FOLDER / "sales_admin_expenses.xlsx", index=False, header=False)
    print(f"  Created sales_admin_expenses.xlsx (Website format)")


def main():
    print("CMO Dashboard Mock Data Generator")
    print("=" * 40)
    print("Generating files in EXACT website export format...")
    print(f"\nOutput folder: {OUTPUT_FOLDER.absolute()}")
    print(f"Random seed: {SEED}")
    print()
    
    generate_market_report()
    generate_marketing_decisions()
    generate_innovation_decisions()
    generate_inventory()
    generate_sales_admin()
    
    print("\n[SUCCESS] All mock data files created!")
    print("\nNow run: python generate_cmo_dashboard_complete.py")


if __name__ == "__main__":
    main()
