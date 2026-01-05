
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

# !!! THIS IS THE FIXED FUNCTION COPIED FROM generate_cmo_dashboard_complete.py to verify logic !!!
def get_innovation_cost(feature_name):
    """Get cost dict for a feature from case parameters."""
    if not costs:
        return {"upfront": 0, "variable": 0}

    def normalize(s):
        """Normalize string for comparison: upper, replace symbols."""
        return s.upper().replace('-', ' ').replace('_', ' ').replace('&', 'AND').replace('  ', ' ').strip()
    
    name_norm = normalize(feature_name)
    
    # 1. Try Direct Key Match
    if name_norm in costs:
        return costs[name_norm]

    # 2. Iterative Match with Normalization
    for key, val in costs.items():
        key_norm = normalize(key)
        
        # Exact normalized match
        if name_norm == key_norm:
            return val, f"Matched {key}"
            
        # Partial match
        if name_norm in key_norm or key_norm in name_norm:
            return val, f"Partial {key}"
            
    return {"upfront": 0, "variable": 0}, "None"

print("\n--- Testing Variations with FIXED Logic ---")
variations = [
    "Stainless Material",
    "Recyclable Materials",
    "Energy Efficiency ", 
    "Lighter & More Compact", 
    "Impact Resistance",
    "Noise Reduction",
    "Improved Battery",
    "Self Cleaning",
    "Speed Settings",
    "Digital Controls",
    "Voice Assistance",
    "Automation",
    "Accessories", 
    "Mapping"      
]

for v in variations:
    res, method = get_innovation_cost(v)
    print(f"Input: '{v}' -> Upfront: {res.get('upfront')}, Var: {res.get('variable')} [{method}]")
