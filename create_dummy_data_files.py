import pandas as pd
import os
from pathlib import Path

# Config
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

def create_shipping_costs():
    print(f"Creating shipping_costs.xlsx in {DATA_DIR}...")
    # Row with "shipping" and "cost" in first column
    data = [
        ["Report: Logistics Costs", ""],
        ["Period: 6", ""],
        ["", ""],
        ["Total Shipping Costs", 125000],  # This matches the regex requirements
        ["Details", ""]
    ]
    df = pd.DataFrame(data)
    df.to_excel(DATA_DIR / "shipping_costs.xlsx", header=False, index=False)
    print("Done.")

def create_esg_report():
    print(f"Creating esg_report.xlsx in {DATA_DIR}...")
    # Rows for emissions, tax, and energy
    data = [
        ["ESG Performance Report", ""],
        ["", ""],
        ["Total Emissions (kg CO2)", 15000],          # Matches 'emission' and 'total'
        ["Carbon Tax Bill Paid", 4500],               # Matches 'tax' and 'paid'/'bill'
        ["Total Energy Consumption (kWh)", 1200000],  # Matches 'energy' and 'consumption'
        ["", ""]
    ]
    df = pd.DataFrame(data)
    df.to_excel(DATA_DIR / "esg_report.xlsx", header=False, index=False)
    print("Done.")

if __name__ == "__main__":
    create_shipping_costs()
    create_esg_report()
