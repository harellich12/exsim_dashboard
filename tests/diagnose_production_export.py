import sys
import os
import io
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook

# Add project root to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../Production Manager Dashboard')))

from generate_production_dashboard_zones import create_zones_dashboard

def seed_shared_outputs():
    try:
        from shared_outputs import SharedOutputManager
        manager = SharedOutputManager()
        
        # Mock CMO Data (Demand)
        manager.export('CMO', {
            'demand_forecast': {
                'Center': 1200, 'West': 800, 'North': 600, 'East': 900, 'South': 700
            },
            'marketing_spend': 50000,
            'pricing': {z: 100 for z in ['Center', 'West', 'North', 'East', 'South']}
        })
        
        # Mock CPO Data (Workforce)
        manager.export('CPO', {
            'workforce_headcount': {
                'Center': 20, 'West': 10, 'North': 5, 'East': 15, 'South': 8
            },
            'payroll_forecast': 100000,
            'hiring_costs': 5000
        })
        print("Seeded Shared Outputs successfully (CMO & CPO data).")
    except Exception as e:
        print(f"Failed to seed shared outputs: {e}")

def test_production_generation_with_overrides():
    print("Testing Production Dashboard Generation...")
    
    # Mock Input Data
    materials_data = {z: {'part_a': 1000, 'part_b': 500} for z in ['Center', 'West', 'North', 'East', 'South']}
    fg_data = {z: {'inventory': 100, 'capacity': 2000} for z in ['Center', 'West', 'North', 'East', 'South']}
    workers_data = {z: {'workers': 10, 'absenteeism': 0.02} for z in ['Center', 'West', 'North', 'East', 'South']}
    machines_data = {z: {'machines': 5, 'modules': 10, 'modules_used': 5} for z in ['Center', 'West', 'North', 'East', 'South']}
    template_data = {'df': None, 'exists': False}
    
    # Decision Overrides
    overrides = {
        'targets': {'Center': 1500, 'West': 800},
        'overtime': {'Center': 'Y', 'West': 'N'}
    }
    
    output = io.BytesIO()
    create_zones_dashboard(
        materials_data, fg_data, workers_data,
        machines_data, template_data,
        output_buffer=output,
        decision_overrides=overrides
    )
    
    output.seek(0)
    wb = load_workbook(output)
    
    # 1. Verify Zone Calculators
    ws1 = wb["ZONE CALCULATORS"]
    # Header was pushed to Row 10 due to spacer row
    # Data starts at Row 11
    target_val = ws1['B11'].value
    print(f"Center Target (FN1): {target_val} (Expected: 1500)")
    
    ot_val = ws1['C11'].value
    print(f"Center Overtime (FN1): {ot_val} (Expected: Y)")
    
    # 2. Verify Cross Reference (CMO Data)
    print("\nChecking Cross Reference Tab...")
    ws_xref = wb["CROSS REFERENCE"]
    
    # Row 6: Center Demand input from Shared CMO
    dem_val = ws_xref['B6'].value
    print(f"Center Demand Ref: {dem_val} (Expected: 1200 from Seed)")
    
    # 3. Verify Upload Ready
    ws_up = wb["UPLOAD READY PRODUCTION"]
    # Row 6: Center Target
    # Check formula link
    target_link = ws_up['C6'].value
    print(f"Upload Target Link: {target_link}")
    
    print("Test Complete.")

if __name__ == "__main__":
    seed_shared_outputs()
    test_production_generation_with_overrides()
