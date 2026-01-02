"""
ExSim War Room - Main Application
Unified strategic dashboard integrating all 7 roles.
"""

import streamlit as st
import pandas as pd

# Import utilities
from utils.state_manager import init_session_state, get_state, set_state
from utils.export_engine import create_decisions_zip
from utils.data_loader import (
    load_market_report, load_workers_balance, load_raw_materials,
    load_finished_goods, load_balance_statements, load_esg_report,
    load_production_data, load_sales_data
)

# Import tabs
from tabs.tab_bulk_upload import render_bulk_upload
from tabs.tab_cmo import render_cmo_tab
from tabs.tab_production import render_production_tab
from tabs.tab_purchasing import render_purchasing_tab
from tabs.tab_logistics import render_logistics_tab
from tabs.tab_cpo import render_cpo_tab
from tabs.tab_esg import render_esg_tab
from tabs.tab_cfo import render_cfo_tab

# =============================================================================
# PAGE CONFIG
# =============================================================================
st.set_page_config(
    page_title="ExSim War Room",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize session state
init_session_state()

# =============================================================================
# SIDEBAR - Data Ingestion
# =============================================================================
with st.sidebar:
    st.title("🎯 ExSim War Room")
    
    # User Manual Download
    from pathlib import Path
    # Try multiple possible locations for the PDF
    possible_paths = [
        Path(__file__).parent.parent / "User Manual.pdf",  # Relative to app.py
        Path.cwd() / "User Manual.pdf",  # Current working directory
        Path.cwd().parent / "User Manual.pdf",  # Parent of CWD
    ]
    manual_path = None
    for p in possible_paths:
        if p.exists():
            manual_path = p
            break
    
    if manual_path:
        st.download_button(
            label="📖 Download User Manual",
            data=manual_path.read_bytes(),
            file_name="ExSim_User_Manual.pdf",
            mime="application/pdf",
            use_container_width=True
        )
    
    st.markdown("---")
    
    st.subheader("📁 Data Upload")
    
    # Market Report
    market_file = st.file_uploader("Market Report", type=['xlsx'], key='market_upload')
    if market_file:
        set_state('market_data', load_market_report(market_file))
        st.success("✓ Market Report")
    
    # Workers Balance
    workers_file = st.file_uploader("Workers Balance", type=['xlsx'], key='workers_upload')
    if workers_file:
        set_state('workers_data', load_workers_balance(workers_file))
        st.success("✓ Workers Balance")
    
    # Raw Materials
    materials_file = st.file_uploader("Raw Materials", type=['xlsx'], key='materials_upload')
    if materials_file:
        set_state('materials_data', load_raw_materials(materials_file))
        st.success("✓ Raw Materials")
    
    # Finished Goods
    fg_file = st.file_uploader("Finished Goods", type=['xlsx'], key='fg_upload')
    if fg_file:
        set_state('finished_goods_data', load_finished_goods(fg_file))
        st.success("✓ Finished Goods")
    
    # Balance Statements
    balance_file = st.file_uploader("Balance Statements", type=['xlsx'], key='balance_upload')
    if balance_file:
        set_state('balance_data', load_balance_statements(balance_file))
        st.success("✓ Balance Statements")
    
    # ESG Report
    esg_file = st.file_uploader("ESG Report", type=['xlsx'], key='esg_upload')
    if esg_file:
        set_state('esg_data', load_esg_report(esg_file))
        st.success("✓ ESG Report")
    
    # Production Data
    prod_file = st.file_uploader("Production Data", type=['xlsx'], key='prod_upload')
    if prod_file:
        set_state('production_data', load_production_data(prod_file))
        st.success("✓ Production Data")
    
    # Sales Admin Expenses (for CMO Last Sales)
    sales_file = st.file_uploader("Sales Admin Expenses", type=['xlsx'], key='sales_upload')
    if sales_file:
        set_state('sales_data', load_sales_data(sales_file))
        st.success("✓ Sales Admin Expenses")
    
    st.markdown("---")
    
    # Export Button
    st.subheader("📤 Export Decisions")
    if st.button("Download All Decisions", type="primary", width='stretch'):
        zip_data = create_decisions_zip()
        st.download_button(
            label="💾 Save decisions.zip",
            data=zip_data,
            file_name="decisions.zip",
            mime="application/zip",
            width='stretch'
        )

# =============================================================================
# MAIN CONTENT - Tabs
# =============================================================================
tabs = st.tabs([
    "📦 Bulk Upload",
    "📢 CMO (Marketing)",
    "🏭 Production",
    "🛒 Purchasing",
    "🚚 Logistics",
    "👥 CPO (People)",
    "🌱 ESG",
    "💰 CFO (Finance)"
])

with tabs[0]:
    render_bulk_upload()

with tabs[1]:
    render_cmo_tab()

with tabs[2]:
    render_production_tab()

with tabs[3]:
    render_purchasing_tab()

with tabs[4]:
    render_logistics_tab()

with tabs[5]:
    render_cpo_tab()

with tabs[6]:
    render_esg_tab()

with tabs[7]:
    render_cfo_tab()
