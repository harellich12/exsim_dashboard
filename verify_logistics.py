
import streamlit as st
import pandas as pd
import sys
import os
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent / "war_room"))

# Mock shared_outputs
sys.modules['shared_outputs'] = type('obj', (object,), {'SHARED_OUTPUTS_FILE': 'shared_outputs.json'})

# Safe import
try:
    from war_room.tabs.tab_logistics import render_logistics_tab, init_logistics_state
    print("[OK] tab_logistics imported successfully")
except ImportError as e:
    print(f"[ERROR] ImportError: {e}")
    sys.exit(1)
except Exception as e:
    print(f"[ERROR] Exception during import: {e}")
    sys.exit(1)

# Basic state check
if 'logistics_data' not in st.session_state:
    # Mock loaded data structure
    st.session_state['logistics_data'] = {
        'benchmarks': {'Train (Center - North)': 13.74, 'Train (Center - West)': 11.19},
        'penalties': {'Center': 5000}
    }

try:
    init_logistics_state()
    print("[OK] init_logistics_state ran successfully")
    
    benchmarks = st.session_state.get('logistics_benchmarks', {})
    if 'Train (Center - North)' in benchmarks:
        print(f"[OK] Benchmarks loaded: {benchmarks}")
    else:
        print("[FAIL] Benchmarks NOT loaded correctly")

    penalties = st.session_state.get('logistics_penalties', {})
    if 'Center' in penalties:
        print(f"[OK] Penalties loaded: {penalties}")
    else:
        print("[FAIL] Penalties NOT loaded correctly")
        
except Exception as e:
    print(f"[ERROR] Runtime Error: {e}")
