"""
ExSim War Room - CMO (Marketing) Tab
Interactive grid for Prices, Promo Budgets, and Target Demand.
Visualizations: Positioning Matrix, Segment Pulse.
Uses proper session state caching to prevent data loss on rerender.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

from utils.state_manager import get_state, set_state

ZONES = ['Center', 'West', 'North', 'East', 'South']
SEGMENTS = ['High', 'Low']


def init_cmo_state():
    """Initialize CMO state with defaults or from uploaded data."""
    if 'cmo_initialized' not in st.session_state:
        st.session_state.cmo_initialized = True
        
        # Initialize marketing decisions
        st.session_state.marketing_decisions = {
            'tv_budget': 0,
            'radio': {z: 0 for z in ZONES},
            'salespeople': {z: {'count': 0, 'salary': 0} for z in ZONES}
        }
        
        # Initialize salespeople DataFrame for AgGrid
        sp_data = []
        for zone in ZONES:
            sp_data.append({'Zone': zone, 'Count': 0, 'Salary': 0})
        st.session_state.cmo_salespeople_df = pd.DataFrame(sp_data)


def render_cmo_tab():
    """Render the CMO (Marketing) tab with subtabs."""
    init_cmo_state()
    
    st.header("📢 CMO Dashboard - Marketing Strategy")
    
    # Load market data if available
    market_data = get_state('market_data')
    
    # Show data status
    if market_data and market_data.get('zones'):
        st.success("✅ Market Report loaded - data populated from upload")
    else:
        st.info("💡 Upload Market Report in sidebar to populate data")
    
    # SUBTABS
    subtab1, subtab2, subtab3 = st.tabs(["📊 Budget & Sales", "🎯 Pricing Strategy", "📈 Analytics"])
    
    # ------------------------------------------------------------------
    # SUBTAB 1: Budget & Salespeople
    # ------------------------------------------------------------------
    with subtab1:
        col1, col2 = st.columns([2, 1])
        
        with col1:
            st.subheader("💰 Budget Allocation")
            
            # TV Budget - use session state value
            tv_budget = st.number_input(
                "TV Budget ($)", 
                min_value=0, 
                value=st.session_state.marketing_decisions.get('tv_budget', 0), 
                step=1000, 
                key='cmo_tv_budget'
            )
            st.session_state.marketing_decisions['tv_budget'] = tv_budget
            
            # Radio by Zone
            st.markdown("**Radio Budget by Zone**")
            radio_cols = st.columns(5)
            for i, zone in enumerate(ZONES):
                with radio_cols[i]:
                    radio_val = st.number_input(
                        zone, 
                        min_value=0, 
                        value=st.session_state.marketing_decisions.get('radio', {}).get(zone, 0), 
                        step=500, 
                        key=f'cmo_radio_{zone}'
                    )
                    st.session_state.marketing_decisions['radio'][zone] = radio_val
        
        with col2:
            st.subheader("📊 Quick Metrics")
            total_marketing = tv_budget + sum(st.session_state.marketing_decisions.get('radio', {}).values())
            st.metric("Total Marketing Spend", f"${total_marketing:,.0f}")
            
            if market_data and market_data.get('zones'):
                avg_share = sum(
                    z.get('High', {}).get('market_share', 0) 
                    for z in market_data['zones'].values()
                ) / max(len(market_data['zones']), 1)
                st.metric("Avg Market Share (High)", f"{avg_share:.1f}%")
        
        st.markdown("---")
        
        # Salespeople Grid - Use cached DataFrame
        st.subheader("👥 Salespeople by Zone")
        st.caption("Edit cells directly. Changes are saved automatically.")
        
        gb = GridOptionsBuilder.from_dataframe(st.session_state.cmo_salespeople_df)
        gb.configure_column('Zone', editable=False, pinned='left')
        gb.configure_column('Count', editable=True, type=['numericColumn'])
        gb.configure_column('Salary', editable=True, type=['numericColumn'])
        gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
        grid_options = gb.build()
        
        grid_response = AgGrid(
            st.session_state.cmo_salespeople_df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.MODEL_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            fit_columns_on_grid_load=True,
            height=200,
            key='cmo_salespeople_grid'
        )
        
        # Update state from grid
        if grid_response.data is not None:
            st.session_state.cmo_salespeople_df = pd.DataFrame(grid_response.data)
            updated_sp = {}
            for _, row in st.session_state.cmo_salespeople_df.iterrows():
                updated_sp[row['Zone']] = {'count': int(row['Count']), 'salary': int(row['Salary'])}
            st.session_state.marketing_decisions['salespeople'] = updated_sp
    
    # ------------------------------------------------------------------
    # SUBTAB 2: Pricing Strategy
    # ------------------------------------------------------------------
    with subtab2:
        st.subheader("💵 Pricing by Zone & Segment")
        
        # Create pricing grid from market data or defaults
        pricing_data = []
        for zone in ZONES:
            for seg in SEGMENTS:
                price = 100  # Default
                if market_data and market_data.get('zones', {}).get(zone, {}).get(seg):
                    price = market_data['zones'][zone][seg].get('price', 100)
                pricing_data.append({
                    'Zone': zone,
                    'Segment': seg,
                    'Current_Price': price,
                    'New_Price': price,
                    'Promo_Budget': 0
                })
        
        if 'cmo_pricing_df' not in st.session_state:
            st.session_state.cmo_pricing_df = pd.DataFrame(pricing_data)
        
        gb = GridOptionsBuilder.from_dataframe(st.session_state.cmo_pricing_df)
        gb.configure_column('Zone', editable=False)
        gb.configure_column('Segment', editable=False)
        gb.configure_column('Current_Price', editable=False)
        gb.configure_column('New_Price', editable=True, type=['numericColumn'])
        gb.configure_column('Promo_Budget', editable=True, type=['numericColumn'])
        gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
        grid_options = gb.build()
        
        price_response = AgGrid(
            st.session_state.cmo_pricing_df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.MODEL_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            fit_columns_on_grid_load=True,
            height=350,
            key='cmo_pricing_grid'
        )
        
        if price_response.data is not None:
            st.session_state.cmo_pricing_df = pd.DataFrame(price_response.data)
    
    # ------------------------------------------------------------------
    # SUBTAB 3: Analytics
    # ------------------------------------------------------------------
    with subtab3:
        viz_col1, viz_col2 = st.columns(2)
        
        with viz_col1:
            st.subheader("🎯 Positioning Matrix")
            
            # Use pricing data if available
            if 'cmo_pricing_df' in st.session_state:
                pos_df = st.session_state.cmo_pricing_df.copy()
                pos_df['Attractiveness'] = 50 + pos_df['Promo_Budget'] / 1000  # Simple calc
                
                fig = px.scatter(
                    pos_df, x='New_Price', y='Attractiveness',
                    color='Zone', symbol='Segment',
                    title='Price vs Attractiveness'
                )
            else:
                fig = go.Figure()
                fig.add_annotation(text="Configure pricing to see matrix", showarrow=False)
            
            fig.update_layout(height=350)
            st.plotly_chart(fig, use_container_width=True)
        
        with viz_col2:
            st.subheader("📊 Segment Pulse")
            
            if market_data and market_data.get('zones'):
                shares = {'Zone': [], 'Segment': [], 'Share': []}
                for zone, data in market_data['zones'].items():
                    for seg in SEGMENTS:
                        shares['Zone'].append(zone)
                        shares['Segment'].append(seg)
                        shares['Share'].append(data.get(seg, {}).get('market_share', 0))
                
                share_df = pd.DataFrame(shares)
                fig = px.bar(
                    share_df, x='Zone', y='Share', color='Segment',
                    barmode='group', title='Market Share by Zone & Segment'
                )
            else:
                fig = go.Figure()
                fig.add_annotation(text="Upload Market Report to see data", showarrow=False)
            
            fig.update_layout(height=350)
            st.plotly_chart(fig, use_container_width=True)
