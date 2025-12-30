"""
ExSim War Room - State Manager
Manages session state for cross-tab synchronization.
"""

import streamlit as st
import pandas as pd
from typing import Any, Optional


def init_session_state():
    """Initialize all session state variables."""
    defaults = {
        # Data from uploads
        'market_data': None,
        'workers_data': None,
        'materials_data': None,
        'finished_goods_data': None,
        'balance_data': None,
        'esg_data': None,
        'production_data': None,
        
        # Cross-tab outputs
        'FORECAST_DEMAND': pd.DataFrame(),
        'PRODUCTION_PLAN': pd.DataFrame(),
        'PROCUREMENT_COST': 0.0,
        'LOGISTICS_COST': 0.0,
        'TOTAL_PAYROLL_CASH': 0.0,
        'ESG_CAPEX': 0.0,
        'ESG_TAX_BILL': 0.0,
        
        # Decision grids (editable)
        'marketing_decisions': None,
        'production_decisions': None,
        'procurement_decisions': None,
        'logistics_decisions': None,
        'people_decisions': None,
        'esg_decisions': None,
        'finance_decisions': None,
    }
    
    for key, default in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = default


def get_state(key: str, default: Any = None) -> Any:
    """Get a value from session state."""
    return st.session_state.get(key, default)


def set_state(key: str, value: Any):
    """Set a value in session state."""
    st.session_state[key] = value


def update_cost(key: str, value: float):
    """Update a cost value and trigger recalculation."""
    st.session_state[key] = value


def get_total_cash_out() -> float:
    """Calculate total cash outflow from all tabs."""
    return (
        get_state('PROCUREMENT_COST', 0) +
        get_state('LOGISTICS_COST', 0) +
        get_state('TOTAL_PAYROLL_CASH', 0) +
        get_state('ESG_CAPEX', 0)
    )


def get_summary_metrics() -> dict:
    """Get summary metrics for the CFO tab."""
    return {
        'Procurement': get_state('PROCUREMENT_COST', 0),
        'Logistics': get_state('LOGISTICS_COST', 0),
        'Payroll': get_state('TOTAL_PAYROLL_CASH', 0),
        'ESG CapEx': get_state('ESG_CAPEX', 0),
        'ESG Tax': get_state('ESG_TAX_BILL', 0),
        'Total Out': get_total_cash_out(),
    }
