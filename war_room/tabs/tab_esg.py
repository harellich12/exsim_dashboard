"""
ExSim War Room - ESG Tab
Interactive grid for Sustainability Initiatives.
Visualization: MACC Curve.
Uses proper session state caching to prevent data loss.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

from utils.state_manager import get_state, set_state


INITIATIVES = [
    {"name": "Solar PV panels", "key": "solar", "unit_cost": 500, "co2_per_unit": 0.5},
    {"name": "Tree plantation (groups of 80)", "key": "trees", "unit_cost": 2000, "co2_per_unit": 2.0},
    {"name": "Green electricity (%)", "key": "green_electricity", "unit_cost": 100, "co2_per_unit": 0.3},
    {"name": "CO2 credits purchase (1 period)", "key": "co2_1", "unit_cost": 30, "co2_per_unit": 1.0},
    {"name": "CO2 credits purchase (2 periods)", "key": "co2_2", "unit_cost": 28, "co2_per_unit": 1.0},
    {"name": "CO2 credits purchase (3 periods)", "key": "co2_3", "unit_cost": 25, "co2_per_unit": 1.0},
]


def init_esg_state():
    """Initialize ESG state."""
    if 'esg_initialized' not in st.session_state:
        st.session_state.esg_initialized = True
        
        esg_data = get_state('esg_data')
        st.session_state.esg_emissions = esg_data.get('emissions', 150) if esg_data else 150
        st.session_state.esg_tax_rate = esg_data.get('tax_rate', 30) if esg_data else 30
        
        # Initiatives grid
        data = []
        for init in INITIATIVES:
            data.append({
                'Initiative': init['name'],
                'Quantity': 0,
                'Unit_Cost': init['unit_cost'],
                'CO2_Per_Unit': init['co2_per_unit']
            })
        st.session_state.esg_initiatives_df = pd.DataFrame(data)


def render_esg_tab():
    """Render the ESG tab with subtabs."""
    init_esg_state()
    
    st.header("🌱 ESG Dashboard - Sustainability Strategy")
    
    esg_data = get_state('esg_data')
    if esg_data:
        st.success("✅ ESG data loaded from upload")
    else:
        st.info("💡 Upload ESG Report for accurate emissions data")
    
    # SUBTABS
    subtab1, subtab2 = st.tabs(["🌿 Initiatives", "📊 MACC Analysis"])
    
    with subtab1:
        col1, col2, col3 = st.columns(3)
        with col1:
            emissions = st.number_input("Current CO2 (tons)", min_value=0, value=st.session_state.esg_emissions, key='esg_em')
            st.session_state.esg_emissions = emissions
        with col2:
            tax_rate = st.number_input("CO2 Tax ($/ton)", min_value=0, value=st.session_state.esg_tax_rate, key='esg_tax')
            st.session_state.esg_tax_rate = tax_rate
        with col3:
            st.metric("Baseline Tax Bill", f"${emissions * tax_rate:,.0f}")
        
        st.markdown("---")
        st.subheader("🌿 ESG Initiatives")
        
        gb = GridOptionsBuilder.from_dataframe(st.session_state.esg_initiatives_df)
        gb.configure_column('Initiative', editable=False, width=300)
        gb.configure_column('Quantity', editable=True, type=['numericColumn'])
        gb.configure_column('Unit_Cost', editable=False)
        gb.configure_column('CO2_Per_Unit', editable=False)
        gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
        grid_options = gb.build()
        
        grid_response = AgGrid(
            st.session_state.esg_initiatives_df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.MODEL_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            fit_columns_on_grid_load=True,
            height=280,
            key='esg_initiatives_grid'
        )
        
        if grid_response.data is not None:
            st.session_state.esg_initiatives_df = pd.DataFrame(grid_response.data)
            
            # Calculate totals
            total_capex = 0
            total_co2 = 0
            for _, row in st.session_state.esg_initiatives_df.iterrows():
                total_capex += row['Quantity'] * row['Unit_Cost']
                total_co2 += row['Quantity'] * row['CO2_Per_Unit']
            
            remaining = max(0, emissions - total_co2)
            tax_bill = remaining * tax_rate
            
            set_state('ESG_CAPEX', total_capex)
            set_state('ESG_TAX_BILL', tax_bill)
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("ESG CapEx", f"${get_state('ESG_CAPEX', 0):,.0f}")
        with col2:
            st.metric("CO2 Reduced", f"{total_co2:.1f} tons" if 'total_co2' in dir() else "0")
        with col3:
            st.metric("Remaining Tax", f"${get_state('ESG_TAX_BILL', 0):,.0f}")
        with col4:
            net = get_state('ESG_CAPEX', 0) + get_state('ESG_TAX_BILL', 0)
            baseline = emissions * tax_rate
            st.metric("Net Savings", f"${baseline - net:,.0f}")
    
    with subtab2:
        st.subheader("📊 Marginal Abatement Cost Curve")
        
        df = st.session_state.esg_initiatives_df.copy()
        df['Total_Cost'] = df['Quantity'] * df['Unit_Cost']
        df['CO2_Reduced'] = df['Quantity'] * df['CO2_Per_Unit']
        df['Cost_Per_Ton'] = df.apply(lambda r: r['Total_Cost'] / r['CO2_Reduced'] if r['CO2_Reduced'] > 0 else 0, axis=1)
        
        active = df[df['Quantity'] > 0].sort_values('Cost_Per_Ton')
        
        if not active.empty:
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=active['Initiative'],
                y=active['Cost_Per_Ton'],
                marker_color=['green' if c < st.session_state.esg_tax_rate else 'red' for c in active['Cost_Per_Ton']]
            ))
            fig.add_hline(y=st.session_state.esg_tax_rate, line_dash="dash", line_color="red", annotation_text=f"Tax: ${st.session_state.esg_tax_rate}/ton")
            fig.update_layout(title='Cost per Ton Abated (Green = Profitable)', height=400, yaxis_title='$/ton')
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Select initiatives to see the MACC curve")
