
import sys
from pathlib import Path
import pandas as pd

# Add root to path
sys.path.append(str(Path.cwd()))

try:
    from case_parameters import MARKET
    print("Successfully imported MARKET from case_parameters")
except ImportError as e:
    print(f"Failed to import MARKET: {e}")
    sys.exit(1)

costs = MARKET.get("INNOVATION_COSTS", {})
print(f"Loaded {len(costs)} innovation costs from case_parameters")

def get_innovation_cost(feature_name):
    """Mirror logic from generate_cmo_dashboard_complete.py"""
    name = feature_name.upper().strip()
    
    # Try direct match
    if name in costs:
        return costs[name], "Direct"
    
    # Try partial match
    for key, val in costs.items():
        if key in name or name in key:
            return val, f"Partial (Matched {key})"
            
    return {"upfront": 0, "variable": 0}, "None"

# Test against the keys themselves (should be direct matches)
print("\n--- Testing Exact Keys ---")
for key in costs.keys():
    res, method = get_innovation_cost(key)
    print(f"Key: '{key}' -> Upfront: {res.get('upfront')}, Var: {res.get('variable')} [{method}]")

# Test against likely variations handling casing/stripping is already done by function
print("\n--- Testing Variations ---")
variations = [
    "Stainless Material",
    "Recyclable Materials",
    "Energy Efficiency ", 
    "Lighter & More Compact", # Note: Ampersand vs "AND"
    "Impact Resistance",
    "Noise Reduction",
    "Improved Battery",
    "Self Cleaning",
    "Speed Settings",
    "Digital Controls",
    "Voice Assistance",
    "Automation",
    "Accessories", # "Multifunctional Accessories"
    "Mapping"      # "Mapping Technology"
]

for v in variations:
    res, method = get_innovation_cost(v)
    print(f"Input: '{v}' -> Upfront: {res.get('upfront')}, Var: {res.get('variable')} [{method}]")
