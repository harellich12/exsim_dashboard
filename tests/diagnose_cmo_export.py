import sys
import os
import io
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook

# Add project root to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../CMO Dashboard')))

from generate_cmo_dashboard_complete import create_complete_dashboard

def test_cmo_generation_with_overrides():
    print("Testing CMO Dashboard Generation with Overrides...")
    
    # Mock Data
    # Mock Data
    market_data = {
        'by_segment': {
            'High': {z: {'my_market_share': 20, 'my_awareness': 50} for z in ['Center', 'West', 'North', 'East', 'South']},
            'Low': {z: {'my_market_share': 10, 'my_awareness': 20} for z in ['Center', 'West', 'North', 'East', 'South']}
        },
        'zones': {
            'Center': {'my_price': 100, 'comp_avg_price': 90},
            'West': {'my_price': 100, 'comp_avg_price': 90},
            'North': {'my_price': 100, 'comp_avg_price': 90},
            'East': {'my_price': 100, 'comp_avg_price': 90},
            'South': {'my_price': 100, 'comp_avg_price': 90},
        }
    }
    
    innovation_features = ["Feature A", "Feature B"]
    
    marketing_template = {
        'tv_budget': 0, 'brand_focus': 0,
        'radio_budgets': {}, 'demand': {}, 'prices': {}, 
        'payment_terms': {}, 'salespeople': {}
    }
    
    sales_data = {
        'by_zone': {z: {'units': 100} for z in ['Center', 'West', 'North', 'East', 'South']},
        'totals': {'units': 500}
    }
    
    inventory_data = {
        'is_stockout': False,
        'by_zone': {z: {'final': 50} for z in ['Center', 'West', 'North', 'East', 'South']}
    }
    
    marketing_intelligence = {
        'economics': {'TV_Cost_Spot': 1000, 'Radio_Cost_Spot': 100},
        'pricing': {z: 95 for z in ['Center', 'West', 'North', 'East', 'South']}
    }
    
    # DECISION OVERRIDES (What ReportBridge sends)
    overrides = {
        'innovation': {'Feature A': 1},
        'tv_spots': 15,
        'brand_focus': 75,
        'zones': {
            'Center': {
                'target_demand': 1234,
                'radio': 5,
                'salespeople': 2,
                'price': 150
            }
        }
    }
    
    # Generate
    output = io.BytesIO()
    create_complete_dashboard(
        market_data, innovation_features, marketing_template,
        sales_data, inventory_data, marketing_intelligence,
        output_buffer=output,
        decision_overrides=overrides
    )
    
    # Verify
    output.seek(0)
    wb = load_workbook(output)
    
    # 1. Check Innovation Lab
    ws_innov = wb["INNOVATION LAB"]
    # Feature A should be 1 (Row 5 usually)
    val_a = ws_innov['B5'].value # Decision col
    print(f"Feature A Decision: {val_a} (Expected: 1)")
    if val_a != 1:
        print("FAIL: Feature A not selected")
    
    # 2. Check Strategy Cockpit
    ws_cockpit = wb["STRATEGY COCKPIT"]
    # TV Spots (Row 9, Col B)
    tv_val = ws_cockpit['B9'].value
    print(f"TV Spots: {tv_val} (Expected: 15)")
    
    # Center Zone (Row 16 usually)
    # Target Demand (Col D)
    dem_val = ws_cockpit['D16'].value
    print(f"Center Demand: {dem_val} (Expected: 1234)")
    
    # Radio (Col E)
    rad_val = ws_cockpit['E16'].value
    print(f"Center Radio: {rad_val} (Expected: 5)")
    
    # Price (Col G)
    price_val = ws_cockpit['G16'].value
    print(f"Center Price: {price_val} (Expected: 150)")
    
    # ... (existing verification code) ...
    # 3. Check Upload Ready Marketing
    ws_upload = wb["UPLOAD READY MARKETING"]
    # TV Row (Row 6) - Amount link
    tv_amount_formula = ws_upload['D6'].value
    print(f"TV Upload Formula: {tv_amount_formula}")
    # Should point to Cockpit C9
    
    # Center Radio (Row 7)
    rad_formula = ws_upload['D7'].value
    print(f"Center Radio Formula: {rad_formula}")
    
    # 4. Check Cross Reference
    print("\nChecking Cross Reference Tab...")
    ws_xref = wb["CROSS REFERENCE"]
    
    # Check Prod Link (Row 5 - Production Plan)
    prod_val = ws_xref['B5'].value
    print(f"Production Plan: {prod_val} (Expected: 1500)")
    
    # Check Prod Link (Row 6 - Utilization)
    util_val = ws_xref['B6'].value
    print(f"Utilization: {util_val} (Expected: 85.0%)")
    
    # Check CFO Link (Row 11 - Liquidity)
    liq_val = ws_xref['B11'].value
    print(f"Liquidity Status: {liq_val} (Expected: HIGH)")
    
    # 5. Check Upload Innovation
    ws_up_innov = wb["UPLOAD READY INNOVATION"]
    innov_val = ws_up_innov['C5'].value
    print(f"Innovation Link: {innov_val} (Expected: =INNOVATION_LAB!B5)")
    
    print("Test Complete.")

def seed_shared_outputs():
    try:
        from shared_outputs import SharedOutputManager
        manager = SharedOutputManager()
        
        # Mock Production Data
        manager.export('Production', {
            'production_plan': {'Zone1': {'Target': 500}, 'Zone2': {'Target': 1000}},
            'capacity_utilization': {'mean': 0.85},
            'overtime_hours': 0,
            'unit_costs': {}
        })
        
        # Mock CFO Data
        manager.export('CFO', {
            'cash_flow_projection': {},
            'debt_levels': 0,
            'liquidity_status': 'HIGH'
        })
        print("Seeded Shared Outputs successfully.")
    except Exception as e:
        print(f"Failed to seed shared outputs: {e}")

if __name__ == "__main__":
    seed_shared_outputs()
    test_cmo_generation_with_overrides()
