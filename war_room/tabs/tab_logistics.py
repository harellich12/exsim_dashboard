"""
ExSim War Room - Logistics Tab
Interactive grids for Shipments and Warehouse Rentals.
Visualization: Warehouse Tetris.
Uses proper session state caching to prevent data loss.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

from utils.state_manager import get_state, set_state

ZONES = ['Center', 'West', 'North', 'East', 'South']
TRANSPORT_MODES = ['Train', 'Truck', 'Airplane']


def init_logistics_state():
    """Initialize Logistics state."""
    if 'logistics_initialized' not in st.session_state:
        st.session_state.logistics_initialized = True
        
        fg_data = get_state('finished_goods_data')
        
        # Warehouse data
        warehouse_data = []
        for zone in ZONES:
            current = fg_data.get('zones', {}).get(zone, {}).get('capacity', 1000) // 100 if fg_data else 10
            warehouse_data.append({
                'Zone': zone,
                'Current_Modules': int(current),
                'Additional_Modules': 0,
                'Module_Cost': 5000
            })
        st.session_state.logistics_warehouses_df = pd.DataFrame(warehouse_data)
        
        # Shipments data
        shipment_data = []
        for i in range(8):
            shipment_data.append({
                'Fortnight': i + 1,
                'Origin': 'Center',
                'Destination': 'West',
                'Material': 'Electroclean',
                'Transport': 'Train',
                'Quantity': 0
            })
        st.session_state.logistics_shipments_df = pd.DataFrame(shipment_data)


def render_logistics_tab():
    """Render the Logistics tab with subtabs."""
    init_logistics_state()
    
    st.header("🚚 Logistics Dashboard - Distribution & Warehousing")
    
    fg_data = get_state('finished_goods_data')
    if fg_data:
        st.success("✅ Finished Goods data loaded from upload")
    else:
        st.info("💡 Upload Finished Goods file for accurate inventory data")
    
    # SUBTABS
    subtab1, subtab2, subtab3 = st.tabs(["🏭 Warehouses", "📦 Shipments", "📊 Inventory Analysis"])
    
    with subtab1:
        st.subheader("🏭 Warehouse Rentals")
        st.caption("Add additional modules to increase capacity (100 units per module)")
        
        gb = GridOptionsBuilder.from_dataframe(st.session_state.logistics_warehouses_df)
        gb.configure_column('Zone', editable=False, pinned='left')
        gb.configure_column('Current_Modules', editable=False)
        gb.configure_column('Additional_Modules', editable=True, type=['numericColumn'])
        gb.configure_column('Module_Cost', editable=False)
        gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
        grid_options = gb.build()
        
        wh_response = AgGrid(
            st.session_state.logistics_warehouses_df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.MODEL_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            fit_columns_on_grid_load=True,
            height=220,
            key='logistics_warehouse_grid'
        )
        
        if wh_response.data is not None:
            st.session_state.logistics_warehouses_df = pd.DataFrame(wh_response.data)
            wh_cost = (st.session_state.logistics_warehouses_df['Additional_Modules'] * 
                       st.session_state.logistics_warehouses_df['Module_Cost']).sum()
            st.session_state.warehouse_cost = wh_cost
        
        st.metric("Total Warehouse Cost", f"${st.session_state.get('warehouse_cost', 0):,.0f}")
    
    with subtab2:
        st.subheader("📦 Shipments")
        st.caption("Plan shipments between zones by fortnight")
        
        gb = GridOptionsBuilder.from_dataframe(st.session_state.logistics_shipments_df)
        gb.configure_column('Fortnight', editable=True, type=['numericColumn'])
        gb.configure_column('Origin', editable=True, cellEditor='agSelectCellEditor', cellEditorParams={'values': ZONES})
        gb.configure_column('Destination', editable=True, cellEditor='agSelectCellEditor', cellEditorParams={'values': ZONES})
        gb.configure_column('Material', editable=False)
        gb.configure_column('Transport', editable=True, cellEditor='agSelectCellEditor', cellEditorParams={'values': TRANSPORT_MODES})
        gb.configure_column('Quantity', editable=True, type=['numericColumn'])
        gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
        grid_options = gb.build()
        
        ship_response = AgGrid(
            st.session_state.logistics_shipments_df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.MODEL_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            fit_columns_on_grid_load=True,
            height=300,
            key='logistics_shipments_grid'
        )
        
        transport_costs = {'Train': 10, 'Truck': 15, 'Airplane': 50}
        
        if ship_response.data is not None:
            st.session_state.logistics_shipments_df = pd.DataFrame(ship_response.data)
            ship_cost = 0
            for _, row in st.session_state.logistics_shipments_df.iterrows():
                ship_cost += row['Quantity'] * transport_costs.get(row['Transport'], 10)
            st.session_state.shipping_cost = ship_cost
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Shipping Cost", f"${st.session_state.get('shipping_cost', 0):,.0f}")
        with col2:
            total = st.session_state.get('warehouse_cost', 0) + st.session_state.get('shipping_cost', 0)
            set_state('LOGISTICS_COST', total)
            st.metric("Total Logistics Cost", f"${total:,.0f}")
    
    with subtab3:
        st.subheader("📊 Warehouse Tetris")
        
        wh_df = st.session_state.logistics_warehouses_df.copy()
        wh_df['Total_Capacity'] = (wh_df['Current_Modules'] + wh_df['Additional_Modules']) * 100
        
        # Current inventory from upload
        fg_data = get_state('finished_goods_data')
        for idx, row in wh_df.iterrows():
            zone = row['Zone']
            inv = 500  # Default
            if fg_data and isinstance(fg_data, dict) and 'zones' in fg_data:
                inv = fg_data.get('zones', {}).get(zone, {}).get('inventory', 500)
            wh_df.at[idx, 'Current_Inventory'] = inv
        
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            name='Current Inventory',
            x=wh_df['Zone'],
            y=wh_df['Current_Inventory'],
            marker_color='steelblue'
        ))
        
        fig.add_trace(go.Scatter(
            name='Capacity',
            x=wh_df['Zone'],
            y=wh_df['Total_Capacity'],
            mode='lines+markers',
            line=dict(color='red', width=2, dash='dash')
        ))
        
        fig.update_layout(
            title='Inventory vs Capacity by Zone',
            height=350,
            yaxis_title='Units'
        )
        
        st.plotly_chart(fig, use_container_width=True)
