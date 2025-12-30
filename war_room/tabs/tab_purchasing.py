"""
ExSim War Room - Purchasing Tab
Interactive grid for Supplier Orders.
Visualization: Inventory Sawtooth.
Uses proper session state caching to prevent data loss.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

from utils.state_manager import get_state, set_state

ZONES = ['Center', 'West', 'North', 'East', 'South']
SUPPLIERS = ['Supplier A', 'Supplier B', 'Supplier C']
PARTS = ['Part A', 'Part B']
FORTNIGHTS = list(range(1, 9))


def init_purchasing_state():
    """Initialize Purchasing state."""
    if 'purchasing_initialized' not in st.session_state:
        st.session_state.purchasing_initialized = True
        st.session_state.purchasing_zone = 'Center'
        
        # Initialize order grids per zone
        st.session_state.purchasing_orders = {}
        for zone in ZONES:
            data = []
            for supplier in SUPPLIERS:
                for part in PARTS:
                    row = {
                        'Supplier': supplier,
                        'Part': part,
                        'Lead_Time': 2 if supplier == 'Supplier A' else (3 if supplier == 'Supplier B' else 4),
                        'Unit_Cost': 100
                    }
                    for fn in FORTNIGHTS:
                        row[f'FN{fn}'] = 0
                    data.append(row)
            st.session_state.purchasing_orders[zone] = pd.DataFrame(data)


def render_purchasing_tab():
    """Render the Purchasing tab with subtabs."""
    init_purchasing_state()
    
    st.header("🛒 Purchasing Dashboard - Procurement Planning")
    
    materials_data = get_state('materials_data')
    
    if materials_data:
        st.success("✅ Raw Materials data loaded from upload")
    else:
        st.info("💡 Upload Raw Materials file for accurate inventory data")
    
    # SUBTABS
    subtab1, subtab2 = st.tabs(["📦 Order Entry", "📈 Inventory Projection"])
    
    with subtab1:
        # Zone selector
        selected_zone = st.selectbox(
            "Select Zone", 
            ZONES, 
            index=ZONES.index(st.session_state.purchasing_zone),
            key='purchasing_zone_select'
        )
        st.session_state.purchasing_zone = selected_zone
        
        st.subheader(f"📦 Orders for {selected_zone}")
        st.caption("Enter order quantities per fortnight. Lead times are shown for reference.")
        
        current_df = st.session_state.purchasing_orders[selected_zone]
        
        gb = GridOptionsBuilder.from_dataframe(current_df)
        gb.configure_column('Supplier', editable=False, pinned='left')
        gb.configure_column('Part', editable=False)
        gb.configure_column('Lead_Time', editable=False)
        gb.configure_column('Unit_Cost', editable=False)
        for fn in FORTNIGHTS:
            gb.configure_column(f'FN{fn}', editable=True, type=['numericColumn'], width=80)
        gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
        grid_options = gb.build()
        
        grid_response = AgGrid(
            current_df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.MODEL_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            fit_columns_on_grid_load=True,
            height=280,
            key=f'purchasing_grid_{selected_zone}'
        )
        
        if grid_response.data is not None:
            st.session_state.purchasing_orders[selected_zone] = pd.DataFrame(grid_response.data)
            
            # Calculate total cost
            order_df = st.session_state.purchasing_orders[selected_zone]
            total_cost = 0
            for _, row in order_df.iterrows():
                for fn in FORTNIGHTS:
                    total_cost += row[f'FN{fn}'] * row['Unit_Cost']
            
            set_state('PROCUREMENT_COST', total_cost)
        
        # Metrics
        st.markdown("---")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Zone Procurement Cost", f"${get_state('PROCUREMENT_COST', 0):,.0f}")
        with col2:
            total_units = sum(
                st.session_state.purchasing_orders[selected_zone][f'FN{fn}'].sum()
                for fn in FORTNIGHTS
            )
            st.metric("Total Units Ordered", f"{total_units:,.0f}")
    
    with subtab2:
        st.subheader("📈 Inventory Sawtooth Projection")
        
        # Simulate inventory projection
        starting_inv = materials_data.get('parts', {}).get('Part A', {}).get('stock', 1000) if materials_data else 1000
        consumption_rate = st.number_input("Consumption Rate per FN", min_value=0, value=200, step=50)
        
        order_df = st.session_state.purchasing_orders[st.session_state.purchasing_zone]
        inventory_levels = []
        inv = starting_inv
        
        for fn in FORTNIGHTS:
            arrivals = 0
            for _, row in order_df.iterrows():
                lead = row['Lead_Time']
                order_fn = fn - lead
                if order_fn >= 1:
                    arrivals += row[f'FN{order_fn}']
            
            inv = inv + arrivals - consumption_rate
            inventory_levels.append({
                'Fortnight': fn,
                'Inventory': max(inv, -500),
                'Arrivals': arrivals
            })
        
        inv_df = pd.DataFrame(inventory_levels)
        
        fig = go.Figure()
        
        fig.add_trace(go.Scatter(
            x=inv_df['Fortnight'],
            y=inv_df['Inventory'],
            mode='lines+markers',
            name='Projected Inventory',
            line=dict(color='blue', width=2),
            fill='tozeroy',
            fillcolor='rgba(0, 100, 200, 0.2)'
        ))
        
        fig.add_trace(go.Bar(
            x=inv_df['Fortnight'],
            y=inv_df['Arrivals'],
            name='Arrivals',
            marker_color='green',
            opacity=0.5
        ))
        
        fig.add_hline(y=0, line_dash="dash", line_color="red", annotation_text="Stockout Threshold")
        
        fig.update_layout(
            title=f'Inventory Projection - {st.session_state.purchasing_zone}',
            xaxis_title='Fortnight',
            yaxis_title='Units',
            height=400
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # Stockout warning
        stockouts = sum(1 for l in inventory_levels if l['Inventory'] < 0)
        if stockouts > 0:
            st.error(f"⚠️ {stockouts} fortnight(s) with stockout risk!")
        else:
            st.success("✅ No stockout risk detected")
