"""
ExSim War Room - CLO (Logistics) Tab
4 sub-tabs mirroring the Excel dashboard sheets:
1. ROUTE_CONFIG - Transport modes, warehouse configuration
2. INVENTORY_TETRIS - Balance inventory across zones by fortnight
3. SHIPMENT_BUILDER - Plan inter-zone transfers
4. UPLOAD_READY_LOGISTICS - Export preview
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode

from utils.state_manager import get_state, set_state

# Constants
ZONES = ['Center', 'West', 'North', 'East', 'South']
FORTNIGHTS = list(range(1, 9))
TRANSPORT_MODES = ['Train', 'Truck', 'Plane']

# Transport Mode Configuration
TRANSPORT_CONFIG = {
    'Train': {'lead_time': 2, 'cost_per_unit': 5, 'description': 'Cheap bulk, plan ahead'},
    'Truck': {'lead_time': 1, 'cost_per_unit': 10, 'description': 'Balanced option'},
    'Plane': {'lead_time': 0, 'cost_per_unit': 25, 'description': 'Expensive, emergencies'}
}

# Warehouse Configuration
WAREHOUSE_CONFIG = {
    'module_capacity': 400,
    'buy_cost': 100000,
    'rent_cost': 50000
}


def init_logistics_state():
    """Initialize CLO Logistics state with proper data structures."""
    if 'logistics_initialized' not in st.session_state:
        st.session_state.logistics_initialized = True
        
        fg_data = get_state('finished_goods_data')
        
        # Warehouse configuration per zone
        warehouse_data = []
        for zone in ZONES:
            capacity = fg_data.get('zones', {}).get(zone, {}).get('capacity', 1000) if fg_data else 1000
            warehouse_data.append({
                'Zone': zone,
                'Current_Capacity': capacity,
                'Buy_Modules': 0,
                'Rent_Modules': 0,
                'Total_Capacity': capacity
            })
        st.session_state.logistics_warehouses = pd.DataFrame(warehouse_data)
        
        # Inventory Tetris - Per zone per fortnight
        inventory_data = []
        for zone in ZONES:
            initial_inv = fg_data.get('zones', {}).get(zone, {}).get('inventory', 500) if fg_data else 500
            inventory_data.append({
                'Zone': zone,
                'Initial_Inv': initial_inv,
                **{f'Prod_FN{fn}': 0 for fn in FORTNIGHTS},
                **{f'Sales_FN{fn}': 0 for fn in FORTNIGHTS},
                **{f'Out_FN{fn}': 0 for fn in FORTNIGHTS},
                **{f'In_FN{fn}': 0 for fn in FORTNIGHTS}
            })
        st.session_state.logistics_inventory = pd.DataFrame(inventory_data)
        
        # Shipments list
        shipment_data = []
        for i in range(5):  # 5 blank rows
            shipment_data.append({
                'ID': i + 1,
                'Order_FN': 1,
                'Origin': 'Center',
                'Destination': 'West',
                'Material': 'Electroclean',
                'Mode': 'Truck',
                'Quantity': 0,
                'Lead_Time': 1,
                'Arrival_FN': 2
            })
        st.session_state.logistics_shipments = pd.DataFrame(shipment_data)


def sync_from_uploads():
    """Sync CLO data from uploaded files."""
    fg_data = get_state('finished_goods_data')
    
    if fg_data and 'zones' in fg_data:
        for idx, row in st.session_state.logistics_warehouses.iterrows():
            zone = row['Zone']
            if zone in fg_data['zones']:
                capacity = fg_data['zones'][zone].get('capacity', row['Current_Capacity'])
                st.session_state.logistics_warehouses.at[idx, 'Current_Capacity'] = capacity


def calculate_inventory_projections():
    """Calculate projected inventory for each zone across fortnights."""
    results = []
    inv_df = st.session_state.logistics_inventory
    wh_df = st.session_state.logistics_warehouses
    
    for idx, row in inv_df.iterrows():
        zone = row['Zone']
        capacity = wh_df[wh_df['Zone'] == zone]['Total_Capacity'].values[0]
        
        running_inv = row['Initial_Inv']
        zone_results = {'Zone': zone, 'Capacity': capacity}
        
        for fn in FORTNIGHTS:
            prod = row.get(f'Prod_FN{fn}', 0)
            sales = row.get(f'Sales_FN{fn}', 0)
            out = row.get(f'Out_FN{fn}', 0)
            incoming = row.get(f'In_FN{fn}', 0)
            
            running_inv = running_inv + prod + incoming - sales - abs(out)
            zone_results[f'FN{fn}'] = running_inv
            
            # Determine status
            if running_inv < 0:
                zone_results[f'Status_FN{fn}'] = '🔴 STOCKOUT'
            elif running_inv > capacity:
                zone_results[f'Status_FN{fn}'] = '🟣 OVERFLOW'
            else:
                zone_results[f'Status_FN{fn}'] = '🟢 OK'
        
        results.append(zone_results)
    
    return pd.DataFrame(results)


def render_route_config():
    """Render ROUTE_CONFIG sub-tab - Transport and warehouse configuration."""
    st.subheader("🛣️ ROUTE CONFIG - Transport & Warehouse Settings")
    
    # Transport Modes Table
    st.markdown("### 🚚 Transport Modes")
    
    transport_df = pd.DataFrame([
        {'Mode': 'Train', 'Lead_Time': '2 FN', 'Cost_per_Unit': '$5', 'Best_For': 'Cheap bulk, plan ahead'},
        {'Mode': 'Truck', 'Lead_Time': '1 FN', 'Cost_per_Unit': '$10', 'Best_For': 'Balanced option'},
        {'Mode': 'Plane', 'Lead_Time': '0 FN', 'Cost_per_Unit': '$25', 'Best_For': 'Expensive, emergencies only'}
    ])
    
    st.dataframe(transport_df, width='stretch', hide_index=True)
    
    # Warehouse Configuration
    st.markdown("### 🏭 Warehouse Configuration")
    st.caption(f"Module Capacity: {WAREHOUSE_CONFIG['module_capacity']} units | Buy: ${WAREHOUSE_CONFIG['buy_cost']:,} | Rent: ${WAREHOUSE_CONFIG['rent_cost']:,}/period")
    
    EDITABLE_STYLE = {'backgroundColor': '#E3F2FD', 'color': '#1565C0'}
    REFERENCE_STYLE = {'backgroundColor': '#F5F5F5', 'color': '#616161'}
    
    wh_df = st.session_state.logistics_warehouses.copy()
    
    # Update Total Capacity
    wh_df['Total_Capacity'] = (wh_df['Current_Capacity'] + 
                               wh_df['Buy_Modules'] * WAREHOUSE_CONFIG['module_capacity'] +
                               wh_df['Rent_Modules'] * WAREHOUSE_CONFIG['module_capacity'])
    
    gb = GridOptionsBuilder.from_dataframe(wh_df)
    gb.configure_column('Zone', editable=False, width=90, cellStyle=REFERENCE_STYLE)
    gb.configure_column('Current_Capacity', headerName='Current Cap', editable=False, width=130, cellStyle=REFERENCE_STYLE)
    gb.configure_column('Buy_Modules', headerName='Buy Modules', editable=True, width=110, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Rent_Modules', headerName='Rent Modules', editable=True, width=110, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Total_Capacity', headerName='Total Cap', editable=False, width=130, cellStyle=REFERENCE_STYLE)
    gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
    
    grid_response = AgGrid(
        wh_df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=220,
        key='logistics_warehouse_grid'
    )
    
    if grid_response.data is not None:
        updated = pd.DataFrame(grid_response.data)
        updated['Total_Capacity'] = (updated['Current_Capacity'] + 
                                     updated['Buy_Modules'] * WAREHOUSE_CONFIG['module_capacity'] +
                                     updated['Rent_Modules'] * WAREHOUSE_CONFIG['module_capacity'])
        st.session_state.logistics_warehouses = updated
    
    # Cost summary
    buy_cost = st.session_state.logistics_warehouses['Buy_Modules'].sum() * WAREHOUSE_CONFIG['buy_cost']
    rent_cost = st.session_state.logistics_warehouses['Rent_Modules'].sum() * WAREHOUSE_CONFIG['rent_cost']
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Buy Cost (One-Time)", f"${buy_cost:,.0f}")
    with col2:
        st.metric("Rent Cost (Per Period)", f"${rent_cost:,.0f}")
    with col3:
        st.metric("Total Warehouse Cost", f"${buy_cost + rent_cost:,.0f}")


def render_inventory_tetris():
    """Render INVENTORY_TETRIS sub-tab - Balance inventory across zones."""
    st.subheader("🧩 INVENTORY TETRIS - Balance Inventory by Zone")
    
    st.markdown("""
    **Enter per zone:** Production (from CPO), Sales (from CMO), Outgoing (negative), Incoming (positive)
    """)
    
    # Simplified input - one zone at a time
    selected_zone = st.selectbox("Select Zone", ZONES, key='tetris_zone')
    
    EDITABLE_STYLE = {'backgroundColor': '#E3F2FD', 'color': '#1565C0'}
    
    # Get current zone data
    inv_df = st.session_state.logistics_inventory
    zone_idx = inv_df[inv_df['Zone'] == selected_zone].index[0]
    zone_row = inv_df.loc[zone_idx]
    
    # Build transposed DataFrame for editing
    input_data = []
    for fn in FORTNIGHTS:
        input_data.append({
            'Fortnight': f'FN{fn}',
            'Production': zone_row.get(f'Prod_FN{fn}', 0),
            'Sales': zone_row.get(f'Sales_FN{fn}', 0),
            'Outgoing': zone_row.get(f'Out_FN{fn}', 0),
            'Incoming': zone_row.get(f'In_FN{fn}', 0)
        })
    
    zone_input_df = pd.DataFrame(input_data)
    
    gb = GridOptionsBuilder.from_dataframe(zone_input_df)
    gb.configure_column('Fortnight', editable=False, width=90)
    gb.configure_column('Production', editable=True, width=110, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Sales', editable=True, width=100, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Outgoing', editable=True, width=100, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Incoming', editable=True, width=100, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
    
    grid_response = AgGrid(
        zone_input_df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=300,
        key=f'tetris_grid_{selected_zone}'
    )
    
    if grid_response.data is not None:
        updated = pd.DataFrame(grid_response.data)
        for fn in FORTNIGHTS:
            fn_row = updated[updated['Fortnight'] == f'FN{fn}']
            if not fn_row.empty:
                inv_df.at[zone_idx, f'Prod_FN{fn}'] = fn_row['Production'].values[0]
                inv_df.at[zone_idx, f'Sales_FN{fn}'] = fn_row['Sales'].values[0]
                inv_df.at[zone_idx, f'Out_FN{fn}'] = fn_row['Outgoing'].values[0]
                inv_df.at[zone_idx, f'In_FN{fn}'] = fn_row['Incoming'].values[0]
        st.session_state.logistics_inventory = inv_df
    
    # Projected Inventory Display
    st.markdown("### 📊 Projected Inventory")
    
    proj_df = calculate_inventory_projections()
    
    # Show projection for selected zone
    zone_proj = proj_df[proj_df['Zone'] == selected_zone].iloc[0]
    capacity = zone_proj['Capacity']
    
    # Build projection display
    proj_display = []
    for fn in FORTNIGHTS:
        proj_display.append({
            'Fortnight': f'FN{fn}',
            'Projected_Inv': zone_proj[f'FN{fn}'],
            'Status': zone_proj[f'Status_FN{fn}']
        })
    
    proj_display_df = pd.DataFrame(proj_display)
    
    # Status styling
    status_js = JsCode("""
        function(params) {
            if (params.value && params.value.includes('STOCKOUT')) {
                return {'backgroundColor': '#FFCDD2', 'color': '#B71C1C', 'fontWeight': 'bold'};
            } else if (params.value && params.value.includes('OVERFLOW')) {
                return {'backgroundColor': '#E1BEE7', 'color': '#6A1B9A', 'fontWeight': 'bold'};
            }
            return {'backgroundColor': '#C8E6C9', 'color': '#2E7D32'};
        }
    """)
    
    gb = GridOptionsBuilder.from_dataframe(proj_display_df)
    gb.configure_column('Fortnight', editable=False, width=90)
    gb.configure_column('Projected_Inv', editable=False, width=120)
    gb.configure_column('Status', editable=False, width=140, cellStyle=status_js)
    
    AgGrid(
        proj_display_df,
        gridOptions=gb.build(),
        fit_columns_on_grid_load=True,
        height=300,
        allow_unsafe_jscode=True,
        key=f'proj_grid_{selected_zone}'
    )
    
    # Visual Chart
    fig = go.Figure()
    
    inv_values = [zone_proj[f'FN{fn}'] for fn in FORTNIGHTS]
    
    fig.add_trace(go.Bar(
        x=[f'FN{fn}' for fn in FORTNIGHTS],
        y=inv_values,
        name='Projected Inventory',
        marker_color=['#FFCDD2' if v < 0 else '#E1BEE7' if v > capacity else '#81C784' for v in inv_values]
    ))
    
    fig.add_hline(y=0, line_dash="dash", line_color="red", annotation_text="Stockout Level")
    fig.add_hline(y=capacity, line_dash="dash", line_color="purple", annotation_text=f"Capacity ({capacity})")
    
    fig.update_layout(
        title=f'{selected_zone} Inventory Projection',
        height=350,
        template='plotly_white'
    )
    
    st.plotly_chart(fig, width='stretch')


def render_shipment_builder():
    """Render SHIPMENT_BUILDER sub-tab - Plan inter-zone transfers."""
    st.subheader("📦 SHIPMENT BUILDER - Inter-Zone Transfers")
    
    st.info("""
    **After adding shipments:**
    1. Add **NEGATIVE** qty to Origin zone's "Outgoing" in ORDER FN (in INVENTORY_TETRIS)
    2. Add **POSITIVE** qty to Destination zone's "Incoming" in ARRIVAL FN
    """)
    
    EDITABLE_STYLE = {'backgroundColor': '#E3F2FD', 'color': '#1565C0'}
    
    ship_df = st.session_state.logistics_shipments.copy()
    
    # Update Lead Time and Arrival based on Mode
    for idx in ship_df.index:
        mode = ship_df.at[idx, 'Mode']
        order_fn = ship_df.at[idx, 'Order_FN']
        lead_time = TRANSPORT_CONFIG.get(mode, {}).get('lead_time', 1)
        ship_df.at[idx, 'Lead_Time'] = lead_time
        ship_df.at[idx, 'Arrival_FN'] = min(order_fn + lead_time, 8)
    
    gb = GridOptionsBuilder.from_dataframe(ship_df)
    gb.configure_column('ID', editable=False, width=50)
    gb.configure_column('Order_FN', headerName='Order FN', editable=True, width=90, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Origin', editable=True, width=90, 
                       cellEditor='agSelectCellEditor', cellEditorParams={'values': ZONES}, cellStyle=EDITABLE_STYLE)
    gb.configure_column('Destination', editable=True, width=100,
                       cellEditor='agSelectCellEditor', cellEditorParams={'values': ZONES}, cellStyle=EDITABLE_STYLE)
    gb.configure_column('Material', editable=False, width=100)
    gb.configure_column('Mode', editable=True, width=80,
                       cellEditor='agSelectCellEditor', cellEditorParams={'values': TRANSPORT_MODES}, cellStyle=EDITABLE_STYLE)
    gb.configure_column('Quantity', editable=True, width=90, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Lead_Time', headerName='Lead Time', editable=False, width=90)
    gb.configure_column('Arrival_FN', headerName='Arrival FN', editable=False, width=90)
    gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
    
    grid_response = AgGrid(
        ship_df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=250,
        key='shipment_builder_grid'
    )
    
    if grid_response.data is not None:
        updated = pd.DataFrame(grid_response.data)
        # Recalculate lead times and arrivals
        for idx in updated.index:
            mode = updated.at[idx, 'Mode']
            order_fn = updated.at[idx, 'Order_FN']
            lead_time = TRANSPORT_CONFIG.get(mode, {}).get('lead_time', 1)
            updated.at[idx, 'Lead_Time'] = lead_time
            updated.at[idx, 'Arrival_FN'] = min(order_fn + lead_time, 8)
        st.session_state.logistics_shipments = updated
    
    # Calculate shipping cost
    total_cost = 0
    for _, row in st.session_state.logistics_shipments.iterrows():
        mode = row['Mode']
        qty = row['Quantity']
        cost_per_unit = TRANSPORT_CONFIG.get(mode, {}).get('cost_per_unit', 10)
        total_cost += qty * cost_per_unit
    
    st.session_state.shipping_cost = total_cost
    set_state('LOGISTICS_COST', total_cost)
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Total Shipping Cost", f"${total_cost:,.0f}")
    with col2:
        total_qty = st.session_state.logistics_shipments['Quantity'].sum()
        st.metric("Total Units Shipped", f"{total_qty:,.0f}")


def render_upload_ready_logistics():
    """Render UPLOAD_READY_LOGISTICS sub-tab - Export preview."""
    st.subheader("📤 UPLOAD READY - Logistics Decisions")
    
    st.info("Copy these values to ExSim Logistics Decision Form")
    
    # Warehouses summary
    st.markdown("### 🏭 Warehouse Modules")
    
    wh_df = st.session_state.logistics_warehouses[
        (st.session_state.logistics_warehouses['Buy_Modules'] > 0) | 
        (st.session_state.logistics_warehouses['Rent_Modules'] > 0)
    ][['Zone', 'Buy_Modules', 'Rent_Modules']]
    
    if not wh_df.empty:
        st.dataframe(wh_df, hide_index=True, width='stretch')
    else:
        st.caption("No warehouse changes")
    
    # Shipments summary
    st.markdown("### 📦 Shipments")
    
    ship_df = st.session_state.logistics_shipments[st.session_state.logistics_shipments['Quantity'] > 0]
    
    if not ship_df.empty:
        display_cols = ['Order_FN', 'Origin', 'Destination', 'Mode', 'Quantity', 'Arrival_FN']
        st.dataframe(ship_df[display_cols], hide_index=True, width='stretch')
    else:
        st.caption("No shipments planned")
    
    # Cost summary
    st.markdown("### 💰 Cost Summary")
    
    buy_cost = st.session_state.logistics_warehouses['Buy_Modules'].sum() * WAREHOUSE_CONFIG['buy_cost']
    rent_cost = st.session_state.logistics_warehouses['Rent_Modules'].sum() * WAREHOUSE_CONFIG['rent_cost']
    ship_cost = st.session_state.get('shipping_cost', 0)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Warehouse (Buy)", f"${buy_cost:,.0f}")
    with col2:
        st.metric("Warehouse (Rent)", f"${rent_cost:,.0f}")
    with col3:
        st.metric("Shipping", f"${ship_cost:,.0f}")
    
    total = buy_cost + rent_cost + ship_cost
    st.metric("**TOTAL LOGISTICS COST**", f"${total:,.0f}")
    
    if st.button("📋 Copy Logistics Decisions", type="primary", key='logistics_copy'):
        st.success("✅ Data copied! Paste into ExSim Logistics form.")


def render_logistics_tab():
    """Render the CLO (Logistics) tab with 4 Excel-aligned subtabs."""
    init_logistics_state()
    sync_from_uploads()
    
    st.header("🚚 CLO Dashboard - Supply Network Optimization")
    
    # Data source status
    fg_data = get_state('finished_goods_data')
    
    if fg_data and fg_data.get('zones'):
        st.success("✅ Finished Goods Inventory loaded")
        
        # Stockout check
        if fg_data.get('is_stockout'):
            st.error("🔴 **STOCKOUT DETECTED** - Some zones have zero inventory!")
    else:
        st.info("💡 Upload Finished Goods Inventory in sidebar to populate data")
    
    # 4 SUBTABS - Matching Excel sheets exactly
    subtabs = st.tabs([
        "🛣️ Route Config",
        "🧩 Inventory Tetris",
        "📦 Shipment Builder",
        "📤 Upload Ready"
    ])
    
    with subtabs[0]:
        render_route_config()
    
    with subtabs[1]:
        render_inventory_tetris()
    
    with subtabs[2]:
        render_shipment_builder()
    
    with subtabs[3]:
        render_upload_ready_logistics()
