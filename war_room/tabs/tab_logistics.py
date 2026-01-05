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

# Import centralized constants from case_parameters
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
try:
    from case_parameters import COMMON
    ZONES = COMMON.get('ZONES', ['Center', 'West', 'North', 'East', 'South'])
    FORTNIGHTS = COMMON.get('FORTNIGHTS', list(range(1, 9)))
    TRANSPORT_MODES = COMMON.get('TRANSPORT_MODES', ['Train', 'Truck', 'Plane'])
except ImportError:
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
        log_data = get_state('logistics_data')
        prod_zones = get_state('production_zones') # Get Production data
        
        # Load benchmarks and penalties
        st.session_state.logistics_benchmarks = log_data.get('benchmarks', {}) if log_data else {}
        st.session_state.logistics_penalties = log_data.get('penalties', {}) if log_data else {}
        
        # Warehouse configuration per zone
        warehouse_data = []
        for zone in ZONES:
            capacity = fg_data.get('zones', {}).get(zone, {}).get('capacity', 1000) if fg_data else 1000
            # Check for penalty
            penalty = st.session_state.logistics_penalties.get(zone, 0)
            
            warehouse_data.append({
                'Zone': zone,
                'Current_Capacity': capacity,
                'Buy_Modules': 0,
                'Rent_Modules': 0,
                'Total_Capacity': capacity,
                'Last_Rent_Penalty': penalty
            })
        st.session_state.logistics_warehouses = pd.DataFrame(warehouse_data)
        
        # Inventory Tetris - Per zone per fortnight
        inventory_data = []
        for zone in ZONES:
            initial_inv = fg_data.get('zones', {}).get(zone, {}).get('inventory', 500) if fg_data else 500
            
            # Sync Production Targets if available
            prod_values = {}
            if prod_zones is not None and not prod_zones.empty:
                 # Find zone row
                 z_row = prod_zones[prod_zones['Zone'] == zone]
                 if not z_row.empty:
                     for fn in FORTNIGHTS:
                         prod_values[f'Prod_FN{fn}'] = z_row.iloc[0].get(f'Target_FN{fn}', 0)
            
            inventory_data.append({
                'Zone': zone,
                'Initial_Inv': initial_inv,
                **{f'Prod_FN{fn}': prod_values.get(f'Prod_FN{fn}', 0) for fn in FORTNIGHTS},
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
    """Sync CLO data from uploaded files and upstream."""
    fg_data = get_state('finished_goods_data')
    prod_zones = get_state('production_zones')
    
    # Sync FG Capacity
    if fg_data and 'zones' in fg_data:
        for idx, row in st.session_state.logistics_warehouses.iterrows():
            zone = row['Zone']
            if zone in fg_data['zones']:
                capacity = fg_data['zones'][zone].get('capacity', row['Current_Capacity'])
                st.session_state.logistics_warehouses.at[idx, 'Current_Capacity'] = capacity

    # Sync Production Targets (Always overwrite Prod_FN if Production Dashboard is active)
    if prod_zones is not None and not prod_zones.empty:
        for idx, row in st.session_state.logistics_inventory.iterrows():
            zone = row['Zone']
            z_row = prod_zones[prod_zones['Zone'] == zone]
            if not z_row.empty:
                for fn in FORTNIGHTS:
                    target = z_row.iloc[0].get(f'Target_FN{fn}', 0)
                    st.session_state.logistics_inventory.at[idx, f'Prod_FN{fn}'] = target


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
            
            # Outgoing is usually negative in Tetris, but if user enters positive 'Out', subtract it.
            # Logic: Inventory = Previous + Prod + In - Sales - Out
            # We assume 'Out' column in Tetris is entered as positive magnitude by user or automation.
            # But wait, original code said "Add NEGATIVE qty". Let's standardize on SUBTRACTING the value.
            # So if automation puts positive 500 in Out column, we subtract it.
            
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
        {'Mode': 'Train', 'Lead Time': '2 FN', 'Cost/Unit': '$5', 'Best For': 'Cheap bulk, plan ahead'},
        {'Mode': 'Truck', 'Lead Time': '1 FN', 'Cost/Unit': '$10', 'Best For': 'Balanced option'},
        {'Mode': 'Plane', 'Lead Time': '0 FN', 'Cost/Unit': '$25', 'Best For': 'Expensive, emergencies only'}
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


def render_route_optimizer():
    """Render ROUTE_OPTIMIZER sub-tab - Transport Mode Matrix & Calculator."""
    st.subheader("🚀 ROUTE OPTIMIZER - Transport Physics")
    
    st.markdown("### 📊 Lowest Cost Transport Matrix")
    
    benchmarks = st.session_state.get('logistics_benchmarks', {})
    
    # Build 5x5 Matrix
    matrix_data = []
    zones = ['Center', 'West', 'North', 'East', 'South']
    
    for origin in zones:
        row = {'Origin': origin}
        for dest in zones:
            if origin == dest:
                row[dest] = "-"
                continue
            
            # Find best mode from benchmarks
            # Benchmark keys format: "Train Center-North" or "Truck West-South" etc.
            # We need to scan keys to find matches for this OD pair
            best_cost = 999999
            best_mode = "N/A"
            
            for mode in ['Train', 'Truck', 'Airplane']:
                # Try various key formats as they might appear in Excel
                # Usually: "{Mode} {Origin}-{Destination}"
                key_candidates = [
                    f"{mode} {origin}-{dest}",
                    f"{mode} {origin} - {dest}",
                    f"{mode} {origin} to {dest}",
                    f"{mode} ({origin} - {dest})",  # Matches Excel format: "Train (Center - North)"
                    f"{mode} ({origin}-{dest})"
                ]
                
                cost = 0
                for k in key_candidates:
                     if k in benchmarks:
                         cost = benchmarks[k]
                         break
                
                # Fallback to default if not found (Test Mode / No Data)
                if cost == 0:
                     cost = TRANSPORT_CONFIG.get(mode, {}).get('cost_per_unit', 99)
                
                if cost < best_cost:
                    best_cost = cost
                    best_mode = mode
            
            row[dest] = f"{best_mode} (${best_cost:.2f})"
        
        matrix_data.append(row)
    
    matrix_df = pd.DataFrame(matrix_data)
    st.dataframe(matrix_df, hide_index=True, width='stretch')
    
    st.markdown("### 🧮 Route Calculator")
    
    c1, c2 = st.columns(2)
    with c1:
        origin = st.selectbox("From (Origin)", zones, key='calc_origin')
    with c2:
        dest = st.selectbox("To (Destination)", zones, key='calc_dest')
    
    if origin == dest:
        st.warning("Origin and Destination are the same.")
    else:
        results = []
        for mode in ['Train', 'Truck', 'Airplane']:
            # Look up specific route cost
            # key = f"{mode} {origin}-{dest}" 
            # Fuzzy match attempt with multiple formats
            cost = 0
            
            # Exact match candidates
            candidates = [
                 f"{mode} {origin}-{dest}",
                 f"{mode} {origin} - {dest}",
                 f"{mode} ({origin} - {dest})",
                 f"{mode} ({origin}-{dest})"
            ]
            
            for k in candidates:
                if k in benchmarks:
                    cost = benchmarks[k]
                    break
            
            # Fallback fuzzy match if exact fails
            if cost == 0:
                for k in benchmarks:
                    if mode in k and origin in k and dest in k:
                        cost = benchmarks[k]
                        break
            
            # Fallback to default
            if cost == 0:
                cost = TRANSPORT_CONFIG.get(mode, {}).get('cost_per_unit', 0)
            
            lead_time = TRANSPORT_CONFIG.get(mode, {}).get('lead_time', 0)
            
            results.append({
                'Mode': mode,
                'Cost Per Unit': cost,
                'Lead Time': f"{lead_time} FN",
                'Total (1000 units)': cost * 1000
            })
        
        res_df = pd.DataFrame(results)
        st.table(res_df.style.format({'Cost Per Unit': '${:.2f}', 'Total (1000 units)': '${:,.0f}'}))


def render_inventory_tetris():
    """Render INVENTORY_TETRIS sub-tab - Balance inventory across zones."""
    st.subheader("🧩 INVENTORY TETRIS - Balance Inventory by Zone")
    
    st.markdown("""
    **Production** is auto-synced from Production tab. **Sales** is auto-synced from CMO.
    **Outgoing/Incoming** are auto-filled by the Shipment Builder.
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
    gb.configure_column('Production', editable=False, width=110) # Made Read-Only
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
                # Production is read-only, dont update it back or we lose sync if user somehow edits it
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
    
    col_inst, col_btn = st.columns([3, 1])
    with col_inst:
        st.info("1. Define shipments below. \n2. Click the button to auto-update Inventory Tetris (Outgoing/Incoming).")
    with col_btn:
        if st.button("🚀 Apply Shipments", type="primary"):
            # Apply Logic
            inv_df = st.session_state.logistics_inventory
            ship_df = st.session_state.logistics_shipments
            
            # Reset Out/In columns first to avoid double counting
            for fn in FORTNIGHTS:
                inv_df[f'Out_FN{fn}'] = 0
                inv_df[f'In_FN{fn}'] = 0
            
            # Process each shipment
            count = 0
            for _, row in ship_df.iterrows():
                qty = row['Quantity']
                if qty > 0:
                    origin = row['Origin']
                    dest = row['Destination']
                    order_fn = int(row['Order_FN'])
                    arrival_fn = int(row['Arrival_FN'])
                    
                    # Add to Origin Outgoing (FN = order_fn)
                    if order_fn <= 8:
                        origin_idx = inv_df[inv_df['Zone'] == origin].index[0]
                        inv_df.at[origin_idx, f'Out_FN{order_fn}'] += qty
                    
                    # Add to Destination Incoming (FN = arrival_fn)
                    if arrival_fn <= 8:
                        dest_idx = inv_df[inv_df['Zone'] == dest].index[0]
                        inv_df.at[dest_idx, f'In_FN{arrival_fn}'] += qty
                    
                    count += 1
            
            st.session_state.logistics_inventory = inv_df
            st.success(f"✅ Applied {count} shipments to Inventory Tetris!")
            
    EDITABLE_STYLE = {'backgroundColor': '#E3F2FD', 'color': '#1565C0'}
    
    ship_df = st.session_state.logistics_shipments.copy()
    
    # Update Lead Time and Arrival based on Mode (Python backup)
    for idx in ship_df.index:
        mode = ship_df.at[idx, 'Mode']
        order_fn = ship_df.at[idx, 'Order_FN']
        lead_time = TRANSPORT_CONFIG.get(mode, {}).get('lead_time', 1)
        ship_df.at[idx, 'Lead_Time'] = lead_time
        ship_df.at[idx, 'Arrival_FN'] = min(order_fn + lead_time, 8)
    
    # JS Logic for Lead Time
    # TRANSPORT_CONFIG = {'Train': 2, 'Truck': 1, 'Plane': 0}
    lead_getter = JsCode("""
        function(params) {
            const modes = {'Train': 2, 'Truck': 1, 'Plane': 0};
            const mode = params.data.Mode;
            if (mode && mode in modes) {
                return modes[mode];
            }
            return 1; // Default
        }
    """)
    
    # JS Logic for Arrival FN
    # Arrival = Order_FN + Lead_Time
    arrival_getter = JsCode("""
        function(params) {
            const modes = {'Train': 2, 'Truck': 1, 'Plane': 0};
            const mode = params.data.Mode;
            let lead = 1;
            if (mode && mode in modes) {
                lead = modes[mode];
            }
            let order = Number(params.data.Order_FN) || 1;
            let arrival = order + lead;
            return (arrival > 8) ? 8 : arrival; // Cap at FN8
        }
    """)
    
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
    
    # Use valueGetter for instant updates
    gb.configure_column('Lead_Time', headerName='Lead Time', editable=False, width=90, 
                       valueGetter=lead_getter)
    gb.configure_column('Arrival_FN', headerName='Arrival FN', editable=False, width=90,
                       valueGetter=arrival_getter)
    
    gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
    
    grid_response = AgGrid(
        ship_df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=250,
        allow_unsafe_jscode=True,
        key='shipment_builder_grid'
    )
    
    if grid_response.data is not None:
        updated = pd.DataFrame(grid_response.data)
        # Recalculate lead times and arrivals (Python side for state)
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
    
    # CSV download button
    import io
    output = io.StringIO()
    output.write("=== WAREHOUSE MODULES ===\n")
    wh_df.to_csv(output, index=False)
    output.write("\n=== SHIPMENTS ===\n")
    ship_df[['Order_FN', 'Origin', 'Destination', 'Mode', 'Quantity', 'Arrival_FN']].to_csv(output, index=False)
    csv_data = output.getvalue()
    
    st.download_button(
        label="📥 Download Decisions as CSV",
        data=csv_data,
        file_name="logistics_decisions.csv",
        mime="text/csv",
        type="primary",
        key='logistics_csv_download'
    )


def render_cross_reference():
    """Render CROSS_REFERENCE sub-tab - Upstream data visibility."""
    st.subheader("🔗 CROSS REFERENCE - Upstream Support")
    st.caption("Live visibility into Purchasing arrivals and CMO demand.")
    
    # Load shared data
    try:
        from shared_outputs import import_dashboard_data
        purch_data = import_dashboard_data('Purchasing') or {}
        cmo_data = import_dashboard_data('CMO') or {}
    except ImportError:
        st.error("Could not load shared_outputs module")
        purch_data = {}
        cmo_data = {}
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📦 Purchasing (Incoming Stock)")
        st.info("Goods arriving that need warehousing.")
        
        spend = purch_data.get('supplier_spend', 0)
        
        st.metric("Total Supplier Spend", f"${spend:,.0f}")
        
        if spend > 0:
            st.success("✅ Purchasing is Active")
        else:
            st.warning("⚠️ No Purchases detected")

    with col2:
        st.markdown("### 📢 CMO (Demand Forecast)")
        st.info("Demand driving your shipping requirements.")
        
        mkt_spend = cmo_data.get('marketing_spend', 0)
        
        st.metric("Marketing Spend", f"${mkt_spend:,.0f}")
        
        # Show demand per zone if available
        demand_forecast = cmo_data.get('demand_forecast', {})
        if demand_forecast:
            st.markdown("**Demand by Zone:**")
            df = pd.DataFrame(list(demand_forecast.items()), columns=['Zone', 'Demand'])
            st.dataframe(df, hide_index=True)


def render_logistics_tab():
    """Render the CLO (Logistics) tab with 4 Excel-aligned subtabs."""
    init_logistics_state()
    sync_from_uploads()
    
    # Header with Download Button
    col_header, col_download = st.columns([4, 1])
    with col_header:
        st.header("🚚 CLO Dashboard - Supply Network Optimization")
    with col_download:
        try:
            from utils.report_bridge import ReportBridge
            excel_buffer = ReportBridge.export_logistics_dashboard()
            st.download_button(
                label="📥 Download Live",
                data=excel_buffer,
                file_name="Logistics_Dashboard_Live.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        except Exception as e:
            try:
                # Fallback to simple create_download_button if ReportBridge fails or method changes
                from utils.report_bridge import create_download_button
                create_download_button('CLO', 'Logistics')
            except:
                st.error(f"Export: {e}")
    
    # Data source status
    fg_data = get_state('finished_goods_data')
    
    if fg_data and fg_data.get('zones'):
        st.success("✅ Finished Goods Inventory loaded")
        
        # Stockout check
        if fg_data.get('is_stockout'):
            st.error("🔴 **STOCKOUT DETECTED** - Some zones have zero inventory!")
    else:
        st.info("💡 Upload Finished Goods Inventory in sidebar to populate data")
    
    # 6 SUBTABS (Updated)
    subtabs = st.tabs([
        "🛣️ Route Config",
        "🚀 Route Optimizer",
        "🧩 Inventory Tetris",
        "📦 Shipment Builder",
        "📤 Upload Ready",
        "🔗 Cross Reference"
    ])
    
    with subtabs[0]:
        render_route_config()
        
    with subtabs[1]:
        render_route_optimizer()
    
    with subtabs[2]:
        render_inventory_tetris()
    
    with subtabs[3]:
        render_shipment_builder()
    
    with subtabs[4]:
        render_upload_ready_logistics()
        
    with subtabs[5]:
        render_cross_reference()
    
    # ---------------------------------------------------------
    # EXSIM SHARED OUTPUTS - EXPORT
    # ---------------------------------------------------------
    try:
        from shared_outputs import export_dashboard_data
        
        # Calculate final outputs for export
        # Shipping Schedule: dict by FN? Or just totals
        # "shipping_schedule": {}, "logistics_costs": 0, "inventory_by_zone": {}
        
        # 1. Logistics Costs
        # Usually from state. 'shipping_cost' + 'warehouse_cost'
        shipping_cost = st.session_state.get('shipping_cost', 0)
        # Calculate warehouse cost
        buy_cost = st.session_state.logistics_warehouses['Buy_Modules'].sum() * WAREHOUSE_CONFIG['buy_cost']
        rent_cost = st.session_state.logistics_warehouses['Rent_Modules'].sum() * WAREHOUSE_CONFIG['rent_cost']
        total_logistics_cost = shipping_cost + buy_cost + rent_cost
        
        # 2. Inventory By Zone (Projected for FN1 or something?)
        # Let's use the 'Initial_Inv' + applied shipments for FN1?
        # Or just 'Initial_Inv' as a proxy for current state?
        # The schema likely wants the 'Closing Inventory' of the period.
        # Let's grab FN1 projection from 'calculate_inventory_projections'
        proj_df = calculate_inventory_projections() # returns columns Zone, Capacity, FN1..FN8
        inventory_by_zone = dict(zip(proj_df['Zone'], proj_df['FN1']))
        
        # 3. Shipping Schedule
        # Shipments list?
        # Let's export list of shipments id/origin/dest/qty
        # Format might be flexible unless specific consumer expects something.
        # CFO just imports 'logistics_costs' mainly.
        # CLO -> CFO: logistics_costs
        
        outputs = {
            'shipping_schedule': st.session_state.logistics_shipments.to_dict('records'),
            'logistics_costs': total_logistics_cost,
            'inventory_by_zone': inventory_by_zone
        }
        
        export_dashboard_data('CLO', outputs)
        
    except Exception as e:
        print(f"Shared Output Export Error: {e}")

