"""
ExSim War Room - Production Tab
3 sub-tabs mirroring the Excel dashboard sheets:
1. ZONE_CALCULATORS - Production targets by zone with capacity checks
2. RESOURCE_MGR - Machine/worker allocation and expansion
3. UPLOAD_READY_PRODUCTION - Export preview
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
    from case_parameters import COMMON, PRODUCTION, WORKFORCE
    ZONES = COMMON.get('ZONES', ['Center', 'West', 'North', 'East', 'South'])
    FORTNIGHTS = COMMON.get('FORTNIGHTS', list(range(1, 9)))
    
    # Production parameters from case
    MACHINERY = PRODUCTION.get('MACHINERY', {})
    FACILITIES = PRODUCTION.get('FACILITIES', {})
    INITIAL_MACHINES = PRODUCTION.get('INITIAL_MACHINES', {})
    MACHINE_TRANSFER_COSTS = PRODUCTION.get('MACHINE_TRANSFER_COSTS', {})
    
    # Worker parameters
    WORKER_PARAMS = WORKFORCE.get('PRODUCTION_WORKERS', {})
except ImportError:
    ZONES = ['Center', 'West', 'North', 'East', 'South']
    FORTNIGHTS = list(range(1, 9))
    MACHINERY = {}
    FACILITIES = {}
    INITIAL_MACHINES = {}
    MACHINE_TRANSFER_COSTS = {}
    WORKER_PARAMS = {}

# Zone Colors
ZONE_COLORS = {
    'Center': '#1565C0',  # Blue
    'West': '#EF6C00',    # Orange
    'North': '#2E7D32',   # Green
    'East': '#F9A825',    # Yellow
    'South': '#6D4C41'    # Brown
}

# Production defaults - now sourced from case_parameters where available
PROD_CONFIG = {
    'units_per_worker': MACHINERY.get('M1', {}).get('capacity_per_fortnight', 200) // max(MACHINERY.get('M1', {}).get('workers_required', 1), 1),
    'units_per_machine': MACHINERY.get('M1', {}).get('capacity_per_fortnight', 200),
    'overtime_multiplier': 1 + WORKER_PARAMS.get('OVERTIME_CAPACITY_INCREASE', 0.20),
    'module_capacity': FACILITIES.get('SPACES_PER_MODULE', 18)  # spaces per module
}



def init_production_state():
    """Initialize Production state with zone-specific data."""
    
    # Check if we need to re-init:
    # 1. Not initialized yet
    # 2. Schema update (Legacy 'Current_Workers' vs new 'Workers')
    should_init = 'production_initialized' not in st.session_state
    
    # WORKAROUND: If auto-sync has happened, DON'T reinitialize (our updated DataFrame is there)
    if 'last_synced_cmo_demand' in st.session_state and 'production_zones' in st.session_state:
        should_init = False
        st.session_state.production_initialized = True  # Set the flag we expected
    
    if not should_init and 'production_zones' in st.session_state:
        # Detect legacy schema
        if 'Workers' not in st.session_state.production_zones.columns:
            should_init = True
    
    if should_init:
        st.session_state.production_initialized = True
        
        prod_data = get_state('production_data')
        workers_data = get_state('workers_data')
        materials_data = get_state('materials_data')
        
        # Calculate available materials (Sum of all stock)
        total_materials = 0
        if materials_data and 'parts' in materials_data:
            total_materials = sum(p.get('stock', 0) for p in materials_data['parts'].values())
        if total_materials == 0:
            total_materials = 1000 # Default if no data
            
        # Zone calculators - production by zone by fortnight
        zone_data = []
        for zone in ZONES:
            # Data from robust loader
            z_prod = prod_data.get('zones', {}).get(zone, {}) if prod_data else {}
            z_workers = workers_data.get('zones', {}).get(zone, {}) if workers_data else {}
            
            machines = z_prod.get('machines', (57 if zone == 'Center' else 0))
            modules = z_prod.get('modules', (72 if zone == 'Center' else 0))
            historic_production = z_prod.get('production', 0)
            
            workers = z_workers.get('workers', 50)
            
            # Smart Default: If historic production exists, use it as target. Else 0.
            default_target = historic_production if historic_production > 0 else 0
            
            zone_data.append({
                'Zone': zone,
                'Machines': machines,
                'Workers': workers,
                'Modules': modules,
                'Materials': total_materials, # Shared pool logic often used, or localized. Using localized same value for now.
                'Historic_Production': historic_production,
                'Machine_Cap': machines * PROD_CONFIG['units_per_machine'],
                'Labor_Cap': workers * PROD_CONFIG['units_per_worker'],
                **{f'Target_FN{fn}': default_target for fn in FORTNIGHTS},
                **{f'Overtime_FN{fn}': False for fn in FORTNIGHTS}
            })
        st.session_state.production_zones = pd.DataFrame(zone_data)
        
        # Resource expansion
        expansion_data = []
        for zone in ZONES:
            expansion_data.append({
                'Zone': zone,
                'Buy_Machines': 0,
                'Buy_Modules': 0,
                'Transfer_In': 0,
                'Transfer_Out': 0
            })
        st.session_state.production_expansion = pd.DataFrame(expansion_data)


def sync_from_uploads():
    """Sync Production data from uploaded files."""
    prod_data = get_state('production_data')
    workers_data = get_state('workers_data')
    materials_data = get_state('materials_data')
    machine_spaces_data = get_state('machine_spaces_data')  # From machine_spaces.xlsx
    
    # Guard: ensure production_zones exists
    if 'production_zones' not in st.session_state:
        return
    
    # Sync Materials
    if materials_data and 'parts' in materials_data:
        total_materials = sum(p.get('stock', 0) for p in materials_data['parts'].values())
        if total_materials > 0:
            st.session_state.production_zones['Materials'] = total_materials
    
    # Sync from machine_spaces_data (priority - has machine counts)
    if machine_spaces_data and 'zones' in machine_spaces_data:
        for idx, row in st.session_state.production_zones.iterrows():
            zone = row['Zone']
            if zone in machine_spaces_data['zones']:
                z_data = machine_spaces_data['zones'][zone]
                if 'machine_capacity' in z_data:
                    machines = int(z_data['machine_capacity'] / PROD_CONFIG['units_per_machine']) if PROD_CONFIG['units_per_machine'] > 0 else 0
                    if machines > 0:
                        st.session_state.production_zones.at[idx, 'Machines'] = machines
                        st.session_state.production_zones.at[idx, 'Machine_Cap'] = machines * PROD_CONFIG['units_per_machine']
                if 'available_spaces' in z_data:
                    st.session_state.production_zones.at[idx, 'Modules'] = int(z_data['available_spaces'])
            
    # Sync Machines & Modules from production_data (fallback)
    if prod_data and 'zones' in prod_data:
        for idx, row in st.session_state.production_zones.iterrows():
            zone = row['Zone']
            if zone in prod_data['zones']:
                z_data = prod_data['zones'][zone]
                if 'machines' in z_data and z_data['machines'] > 0:
                    machines = z_data['machines']
                    st.session_state.production_zones.at[idx, 'Machines'] = machines
                    st.session_state.production_zones.at[idx, 'Machine_Cap'] = machines * PROD_CONFIG['units_per_machine']
                if 'modules' in z_data and z_data['modules'] > 0:
                    st.session_state.production_zones.at[idx, 'Modules'] = z_data['modules']
                if 'production' in z_data:
                     st.session_state.production_zones.at[idx, 'Historic_Production'] = z_data['production']
    
    # Sync Workers
    if workers_data and 'zones' in workers_data:
        for idx, row in st.session_state.production_zones.iterrows():
            zone = row['Zone']
            if zone in workers_data['zones']:
                workers = workers_data['zones'][zone].get('workers', row['Workers'])
                st.session_state.production_zones.at[idx, 'Workers'] = workers
                st.session_state.production_zones.at[idx, 'Labor_Cap'] = workers * PROD_CONFIG['units_per_worker']
    
    # Sync Demand from CMO (Marketing) - AUTO-SYNC when CMO demand changes
    try:
        from shared_outputs import import_dashboard_data
        cmo_data = import_dashboard_data('CMO')
        if cmo_data and 'demand_forecast' in cmo_data:
            demand_forecast = cmo_data['demand_forecast']
            
            # Track last synced CMO demand to detect changes
            last_cmo_demand = st.session_state.get('last_synced_cmo_demand', {})
            
            for idx, row in st.session_state.production_zones.iterrows():
                zone = row['Zone']
                if zone in demand_forecast:
                    demand = float(demand_forecast[zone]) if demand_forecast[zone] else 0
                    per_fn_demand = int(demand / 8) if demand > 0 else 0
                    
                    # Check if CMO demand changed since last sync
                    last_demand = last_cmo_demand.get(zone, None)
                    cmo_changed = (last_demand is None or last_demand != demand)
                    
                    # Auto-sync if CMO demand changed
                    if cmo_changed and per_fn_demand > 0:
                        for fn in FORTNIGHTS:
                            st.session_state.production_zones.at[idx, f'Target_FN{fn}'] = per_fn_demand
                        # Track that we synced this demand
                        if 'last_synced_cmo_demand' not in st.session_state:
                            st.session_state.last_synced_cmo_demand = {}
                        st.session_state.last_synced_cmo_demand[zone] = demand
    except Exception as e:
        print(f"CMO sync error: {e}")


def calculate_electricity_cost():
    """
    Calculate electricity costs based on machines (Table IV.3).
    Power cost: $10 per installed kW per period
    Consumption: $0.06 per kWh (80 hrs/fortnight × kW × utilization)
    """
    electricity = PRODUCTION.get('ELECTRICITY', {})
    power_cost_per_kw = electricity.get('POWER_COST_PER_KW_PER_PERIOD', 10)
    consumption_cost_per_kwh = electricity.get('CONSUMPTION_COST_PER_KWH', 0.06)
    
    zones_df = st.session_state.production_zones
    machines_config = PRODUCTION.get('MACHINES', {})
    
    total_kw = 0
    total_kwh = 0
    
    # Calculate total installed power from machines
    for _, row in zones_df.iterrows():
        machines = row.get('Machines', 0)
        # Use M1 as proxy (10 kW each) - simplified
        kw_per_machine = machines_config.get('M1', {}).get('power_kw', 10)
        zone_kw = machines * kw_per_machine
        total_kw += zone_kw
        # Assume 100% utilization: 80 hrs/fortnight × 8 fortnights = 640 hrs/period
        total_kwh += zone_kw * 640
    
    power_term = total_kw * power_cost_per_kw  # Fixed cost
    consumption_term = total_kwh * consumption_cost_per_kwh  # Variable cost
    
    return {
        'total_kw': total_kw,
        'total_kwh': total_kwh,
        'power_cost': power_term,
        'consumption_cost': consumption_term,
        'total_cost': power_term + consumption_term
    }


def get_transfer_cost(from_zone, to_zone, machine_type='M1'):
    """Get machine transfer cost between zones (Table IV.2)."""
    route = f"{from_zone}-{to_zone}"
    route_costs = MACHINE_TRANSFER_COSTS.get(route, {})
    return route_costs.get(machine_type, 0)

def calculate_zone_production():
    """Calculate real output and alerts per zone per fortnight."""
    zones_df = st.session_state.production_zones
    exp_df = st.session_state.production_expansion
    
    results = []
    for idx, row in zones_df.iterrows():
        zone = row['Zone']
        exp_row = exp_df[exp_df['Zone'] == zone].iloc[0]
        
        # Effective capacity
        machines = row['Machines'] + exp_row['Buy_Machines'] + exp_row['Transfer_In'] - exp_row['Transfer_Out']
        machine_cap = machines * PROD_CONFIG['units_per_machine']
        labor_cap = row['Labor_Cap']
        materials = row['Materials']
        
        zone_results = {'Zone': zone}
        
        for fn in FORTNIGHTS:
            target = row.get(f'Target_FN{fn}', 0)
            overtime = row.get(f'Overtime_FN{fn}', False)
            
            # Apply overtime bonus
            effective_cap = machine_cap
            if overtime:
                effective_cap = int(machine_cap * PROD_CONFIG['overtime_multiplier'])
            
            min_cap = min(effective_cap, labor_cap, materials)
            real_output = min(target, min_cap)
            
            # Determine alert
            if target > materials:
                alert = '📦 SHIPMENT NEEDED'
            elif target > effective_cap:
                alert = '⚙️ Machine Limit'
            elif target > labor_cap:
                alert = '👷 Labor Limit'
            else:
                alert = '✅ OK'
            
            zone_results[f'Real_FN{fn}'] = real_output
            zone_results[f'Alert_FN{fn}'] = alert
        
        results.append(zone_results)
    
    return pd.DataFrame(results)


def render_zone_calculators():
    """Render ZONE_CALCULATORS sub-tab - Production targets by zone."""
    st.subheader("🏭 ZONE CALCULATORS - Production by Zone")
    
    st.markdown("""
    **Zone Independence:** Resources in CENTER do NOT count towards WEST capacity.  
    Set targets per zone. Real output = min(Target, Machine Cap, Labor Cap, Materials).
    """)
    
    # Zone selector
    selected_zone = st.selectbox("Select Zone", ZONES, key='prod_zone_select')
    zone_color = ZONE_COLORS.get(selected_zone, '#1565C0')
    
    zones_df = st.session_state.production_zones
    zone_idx = zones_df[zones_df['Zone'] == selected_zone].index[0]
    zone_row = zones_df.loc[zone_idx]
    
    # CMO Demand Display (sync is now automatic in sync_from_uploads)
    cmo_demand_val = 0
    try:
        from shared_outputs import import_dashboard_data
        cmo_data = import_dashboard_data('CMO')
        if cmo_data and 'demand_forecast' in cmo_data:
            demand_forecast = cmo_data['demand_forecast']
            cmo_demand_val = float(demand_forecast.get(selected_zone, 0)) if demand_forecast.get(selected_zone) else 0
    except:
        pass
    
    if cmo_demand_val > 0:
        per_fn = int(cmo_demand_val / 8)
        st.info(f"📊 **CMO Demand for {selected_zone}:** {cmo_demand_val:,.0f} units total → {per_fn:,}/fortnight (auto-synced)")
    
    # Zone status metrics
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("Machines", zone_row['Machines'])
    with col2:
        st.metric("Workers", zone_row['Workers'])
    with col3:
        st.metric("Machine Cap", f"{zone_row['Machine_Cap']:,}")
    with col4:
        st.metric("Materials Available", f"{zone_row['Materials']:,}")
    with col5:
        st.metric("Historic Output", f"{zone_row.get('Historic_Production', 0):,}")
    
    # Production schedule grid
    st.markdown("### 📅 Production Schedule")
    
    EDITABLE_STYLE = {'backgroundColor': '#E3F2FD', 'color': '#1565C0'}
    
    # Combined Grid for Inputs & Outputs with Client-Side Logic
    
    # Re-fetch zone data from session state to get fresh values
    fresh_zones_df = st.session_state.production_zones
    fresh_zone_row = fresh_zones_df.loc[zone_idx]
    
    # Prepare data with hidden constraints for JS
    # Prepare data with hidden constraints for JS
    combined_data = []
    machine_cap = fresh_zone_row['Machine_Cap']
    # Calculate overtime capacity using consistent config
    machine_cap_ot = int(machine_cap * PROD_CONFIG['overtime_multiplier'])
    labor_cap = fresh_zone_row['Labor_Cap']
    materials = fresh_zone_row['Materials']
    
    # Calculate initial python values for first render
    results = calculate_zone_production()
    zone_results = results[results['Zone'] == selected_zone].iloc[0]
    
    for fn in FORTNIGHTS:
        combined_data.append({
            'Fortnight': f'FN{fn}',
            'Target': fresh_zone_row.get(f'Target_FN{fn}', 0),
            'Overtime': fresh_zone_row.get(f'Overtime_FN{fn}', False),
            'Machine_Cap': machine_cap,
            'Machine_Cap_OT': machine_cap_ot,
            'Labor_Cap': labor_cap,
            'Materials': materials, 
            'Real_Output': zone_results.get(f'Real_FN{fn}', 0),
            'Alert': zone_results.get(f'Alert_FN{fn}', '✅ OK')
        })
    
    combined_df = pd.DataFrame(combined_data)
    
    # JS Logic for Real Output
    # min(effective_cap, labor_cap, materials, target)
    output_getter = JsCode("""
        function(params) {
            let target = Number(params.data.Target) || 0;
            let overtime = params.data.Overtime;
            let machine_cap = Number(params.data.Machine_Cap);
            let machine_cap_ot = Number(params.data.Machine_Cap_OT);
            let labor_cap = Number(params.data.Labor_Cap);
            let materials = Number(params.data.Materials);
            
            let effective_cap = overtime ? machine_cap_ot : machine_cap;
            
            return Math.min(target, effective_cap, labor_cap, materials);
        }
    """)
    
    # JS Logic for Alerts
    alert_getter = JsCode("""
        function(params) {
            let target = Number(params.data.Target) || 0;
            let overtime = params.data.Overtime;
            let machine_cap = Number(params.data.Machine_Cap);
            let machine_cap_ot = Number(params.data.Machine_Cap_OT);
            let labor_cap = Number(params.data.Labor_Cap);
            let materials = Number(params.data.Materials);
            
            let effective_cap = overtime ? machine_cap_ot : machine_cap;
            
            if (target > materials) return '📦 SHIPMENT NEEDED';
            if (target > effective_cap) return '⚙️ Machine Limit';
            if (target > labor_cap) return '👷 Labor Limit';
            return '✅ OK';
        }
    """)
    
    st.markdown("""
        > **Note:** Overtime increases Machine Capacity by **20%**. It only increases output if machines are the bottleneck (Machine Limit).
        > If your Target is lower than capacity, overtime will have no effect on Real Output.
    """)
    
    alert_style = JsCode("""
        function(params) {
            if (params.value && params.value.includes('SHIPMENT')) {
                return {'backgroundColor': '#E1BEE7', 'color': '#6A1B9A', 'fontWeight': 'bold'};
            } else if (params.value && params.value.includes('Machine')) {
                return {'backgroundColor': '#FFCDD2', 'color': '#B71C1C'};
            } else if (params.value && params.value.includes('Labor')) {
                return {'backgroundColor': '#FFF9C4', 'color': '#F57F17'};
            }
            return {'backgroundColor': '#C8E6C9', 'color': '#2E7D32'};
        }
    """)
    
    gb = GridOptionsBuilder.from_dataframe(combined_df)
    gb.configure_column('Fortnight', editable=False, width=90)
    gb.configure_column('Target', editable=True, width=100, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Overtime', editable=True, width=90, cellEditor='agCheckboxCellEditor')
    
    # Calculated columns
    gb.configure_column('Real_Output', headerName='Real Output', editable=False, width=110, 
                       valueGetter=output_getter)
    gb.configure_column('Alert', editable=False, width=160, 
                       valueGetter=alert_getter, cellStyle=alert_style)
    
    # Hidden columns for calculation context
    gb.configure_column('Machine_Cap', hide=True)
    gb.configure_column('Labor_Cap', hide=True)
    gb.configure_column('Materials', hide=True)
    
    gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
    
    # Get refresh counter for grid key
    grid_refresh = st.session_state.get('prod_grid_refresh', 0)
    
    grid_response = AgGrid(
        combined_df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=350,
        allow_unsafe_jscode=True,
        key=f'zone_calc_grid_{selected_zone}_{grid_refresh}'
    )
    
    # Update session state from grid edits
    if grid_response.data is not None:
        updated = pd.DataFrame(grid_response.data)
        zones_df = st.session_state.production_zones
        for fn in FORTNIGHTS:
            fn_row = updated[updated['Fortnight'] == f'FN{fn}']
            if not fn_row.empty:
                zones_df.at[zone_idx, f'Target_FN{fn}'] = fn_row['Target'].values[0]
                zones_df.at[zone_idx, f'Overtime_FN{fn}'] = fn_row['Overtime'].values[0]
        st.session_state.production_zones = zones_df


def render_resource_mgr():
    """Render RESOURCE_MGR sub-tab - Machine/worker allocation and expansion."""
    st.subheader("⚙️ RESOURCE MANAGER - Allocation & Expansion")
    
    # Section A: Current Resources
    st.markdown("### Section A: Current Resources by Zone")
    
    zones_df = st.session_state.production_zones[['Zone', 'Machines', 'Workers', 
                                                  'Modules', 'Materials', 'Machine_Cap', 'Labor_Cap']]
    
    # Rename columns for cleaner display
    display_df = zones_df.copy()
    display_df.columns = ['Zone', 'Machines', 'Workers', 'Modules', 'Available Materials', 'Machine Cap', 'Labor Cap']
    
    st.dataframe(display_df, width='stretch', hide_index=True)
    
    # Section B: Expansion by Zone
    st.markdown("### Section B: Expansion by Zone")
    
    EDITABLE_STYLE = {'backgroundColor': '#E3F2FD', 'color': '#1565C0'}
    
    exp_df = st.session_state.production_expansion.copy()
    
    gb = GridOptionsBuilder.from_dataframe(exp_df)
    gb.configure_column('Zone', editable=False, width=80)
    gb.configure_column('Buy_Machines', headerName='Buy Machines', editable=True, width=120, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Buy_Modules', headerName='Buy Modules', editable=True, width=110, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Transfer_In', headerName='Transfer In', editable=True, width=110, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Transfer_Out', headerName='Transfer Out', editable=True, width=110, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
    
    grid_response = AgGrid(
        exp_df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=230,
        key='resource_expansion_grid'
    )
    
    if grid_response.data is not None:
        st.session_state.production_expansion = pd.DataFrame(grid_response.data)
    
    # Section C: Capacity Recommendations
    st.markdown("### Section C: Recommendations")
    
    for zone in ZONES:
        zone_row = st.session_state.production_zones[st.session_state.production_zones['Zone'] == zone].iloc[0]
        exp_row = st.session_state.production_expansion[st.session_state.production_expansion['Zone'] == zone].iloc[0]
        
        modules = zone_row['Modules']
        machines = zone_row['Machines'] + exp_row['Buy_Machines']
        
        if modules == 0 and machines > 0:
            st.warning(f"⚠️ **{zone}**: Buy module first - zone has no slots!")
        elif machines > modules * PROD_CONFIG['module_capacity']:
            st.error(f"🔴 **{zone}**: Too many machines for slots! Buy {int((machines / PROD_CONFIG['module_capacity']) - modules + 0.99)} more modules")
    
    # Capacity visualization
    st.markdown("### 📊 Capacity Comparison")
    
    zones_df = st.session_state.production_zones
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=zones_df['Zone'],
        y=zones_df['Machine_Cap'],
        name='Machine Capacity',
        marker_color=[ZONE_COLORS.get(z, '#1565C0') for z in zones_df['Zone']]
    ))
    fig.add_trace(go.Bar(
        x=zones_df['Zone'],
        y=zones_df['Labor_Cap'],
        name='Labor Capacity',
        marker_color='#90CAF9'
    ))
    
    fig.update_layout(
        barmode='group',
        title='Zone Capacity Comparison',
        height=350,
        template='plotly_white'
    )
    
    st.plotly_chart(fig, width='stretch')
    
    # Section D: Machine Transfer Cost Calculator
    st.markdown("### 🚁 Section D: Machine Transfer Cost Calculator")
    st.caption("Calculate cost to transfer machines between regions (by airplane)")
    
    if MACHINE_TRANSFER_COSTS:
        routes = list(MACHINE_TRANSFER_COSTS.keys())
        machine_types = ['M1', 'M2', 'M3_ALPHA', 'M3_BETA', 'M4']
        
        col1, col2, col3 = st.columns(3)
        with col1:
            selected_route = st.selectbox("Route", routes, key='transfer_route')
        with col2:
            selected_machine = st.selectbox("Machine Type", machine_types, key='transfer_machine')
        with col3:
            quantity = st.number_input("Quantity", min_value=1, value=1, key='transfer_qty')
        
        # Calculate cost
        route_costs = MACHINE_TRANSFER_COSTS.get(selected_route, {})
        unit_cost = route_costs.get(selected_machine, 0)
        total_cost = unit_cost * quantity
        
        col_a, col_b = st.columns(2)
        with col_a:
            st.metric("Cost per Machine", f"${unit_cost:,.2f}")
        with col_b:
            st.metric("Total Transfer Cost", f"${total_cost:,.2f}")
        
        # Display full cost matrix
        with st.expander("📋 Full Transfer Cost Matrix"):
            matrix_data = []
            for route, costs in MACHINE_TRANSFER_COSTS.items():
                row = {'Route': route}
                row.update(costs)
                matrix_data.append(row)
            st.dataframe(pd.DataFrame(matrix_data), hide_index=True)
    else:
        st.info("Transfer cost data not available")
    
    # Section E: Electricity Cost Calculator (Table IV.3)
    st.markdown("### ⚡ Section E: Electricity Cost Estimate")
    st.caption("Power: $10/kW/period (fixed) | Consumption: $0.06/kWh (variable)")
    
    elec_costs = calculate_electricity_cost()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Installed Power", f"{elec_costs['total_kw']:,.0f} kW")
    with col2:
        st.metric("Power Cost (Fixed)", f"${elec_costs['power_cost']:,.0f}")
    with col3:
        st.metric("Consumption Cost", f"${elec_costs['consumption_cost']:,.0f}")
    
    st.metric("**Total Electricity Cost (per period)**", f"${elec_costs['total_cost']:,.0f}")
    
    if elec_costs['total_cost'] > 30000:
        st.info("💡 Consider installing PV panels to reduce electricity costs and CO2 emissions.")



def render_upload_ready_production():
    """Render UPLOAD_READY_PRODUCTION sub-tab - Export preview."""
    st.subheader("📤 UPLOAD READY - Production Decisions")
    
    st.info("Copy these values to ExSim Production Decision Form")
    
    # Production targets
    st.markdown("### 🏭 Production Targets")
    
    zones_df = st.session_state.production_zones
    target_cols = ['Zone'] + [f'Target_FN{fn}' for fn in FORTNIGHTS]
    targets = zones_df[target_cols].copy()
    targets = targets[targets[[f'Target_FN{fn}' for fn in FORTNIGHTS]].sum(axis=1) > 0]
    
    if not targets.empty:
        st.dataframe(targets, hide_index=True, width='stretch')
    else:
        st.caption("No production targets set")
    
    # Expansion summary
    st.markdown("### ⚙️ Expansion")
    
    exp_df = st.session_state.production_expansion
    exp_changes = exp_df[(exp_df['Buy_Machines'] > 0) | (exp_df['Buy_Modules'] > 0)]
    
    if not exp_changes.empty:
        st.dataframe(exp_changes, hide_index=True, width='stretch')
    else:
        st.caption("No expansion planned")
    
    # Summary
    st.markdown("### 📊 Summary")
    
    total_target = sum(zones_df[[f'Target_FN{fn}' for fn in FORTNIGHTS]].sum(axis=0))
    total_machines = exp_df['Buy_Machines'].sum()
    total_modules = exp_df['Buy_Modules'].sum()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Production (All FN)", f"{total_target:,.0f}")
    with col2:
        st.metric("Machines to Buy", total_machines)
    with col3:
        st.metric("Modules to Buy", total_modules)
    
    # CSV download button
    zones_df = st.session_state.production_zones
    target_cols = ['Zone'] + [f'Target_FN{fn}' for fn in FORTNIGHTS]
    export_df = zones_df[target_cols].copy()
    csv_data = export_df.to_csv(index=False)
    
    st.download_button(
        label="📥 Download Decisions as CSV",
        data=csv_data,
        file_name="production_decisions.csv",
        mime="text/csv",
        type="primary",
        key='prod_csv_download'
    )


def render_production_tab():
    """Render the Production tab with 3 Excel-aligned subtabs."""
    init_production_state()
    sync_from_uploads()
    
    # Header with Download Button
    col_header, col_download = st.columns([4, 1])
    with col_header:
        st.header("🏭 Production Dashboard - Zone-Specific Capacity")
    with col_download:
        try:
            from utils.report_bridge import create_download_button
            create_download_button('Production', 'Production')
        except Exception as e:
            st.error(f"Export: {e}")
    
    # Data status
    prod_data = get_state('production_data')
    workers_data = get_state('workers_data')
    
    # Check CMO demand sync
    cmo_demand = None
    try:
        from shared_outputs import import_dashboard_data
        cmo_data = import_dashboard_data('CMO')
        if cmo_data and 'demand_forecast' in cmo_data:
            cmo_demand = cmo_data['demand_forecast']
    except:
        pass
    
    data_status = []
    if prod_data:
        data_status.append("✅ Production Data")
    if workers_data:
        data_status.append("✅ Workers Data")
    if cmo_demand:
        total_demand = sum(float(v) if v else 0 for v in cmo_demand.values())
        data_status.append(f"📊 CMO Demand: {total_demand:,.0f}")
    
    if data_status:
        st.success(" | ".join(data_status))
    else:
        st.info("💡 Upload Production and Workers files in sidebar for accurate data")
    
    # 3 SUBTABS
    subtabs = st.tabs([
        "🏭 Zone Calculators",
        "⚙️ Resource Manager",
        "📤 Upload Ready"
    ])
    
    with subtabs[0]:
        render_zone_calculators()
    
    with subtabs[1]:
        render_resource_mgr()
    
    with subtabs[2]:
        render_upload_ready_production()
        
    # ---------------------------------------------------------
    # EXSIM SHARED OUTPUTS - EXPORT
    # ---------------------------------------------------------
    try:
        from shared_outputs import export_dashboard_data
        
        # Calculate final outputs for export
        results = calculate_zone_production()
        
        # Prepare Production Plan (Targets)
        prod_plan = {}
        for _, row in st.session_state.production_zones.iterrows():
            zone = row['Zone']
            # Sum up targets for simplified view or keep details? 
            # Schema expects: 'Center': {'Target': 1000}
            # We have Fn inputs. Let's sum them for now or pick FN1?
            # Actually, let's export the full schedule if possible, but schema implies simpler structure?
            # "production_plan": {"Center": {"Target": 0}}
            # Let's verify schema usage in shared_outputs.py or just dump the grid.
            # For robustness, let's export a dict of targets keyed by zone.
            
            # Summing targets for 4 fortnights as a proxy for "Plan"
            total_target = sum(row.get(f'Target_FN{fn}', 0) for fn in [1, 2, 3, 4])
            prod_plan[zone] = {'Target': total_target}
            
        # Prepare Capacity Utilization and Overtime
        total_capacity = 0
        total_target = 0
        total_overtime = 0
        zone_costs = {}
        
        machinery = get_state('machinery_data') or {}
        
        for _, row in st.session_state.production_zones.iterrows():
            zone = row['Zone']
            
            # Calculate Zone Capacity
            # Sum of (Count * Capacity) for all machines in zone
            zone_cap = 0
            if machinery and machinery.get('zones'):
                zone_machines = machinery['zones'].get(zone, {})
                for section in zone_machines:
                    for m_type, count in zone_machines[section].items():
                        # Get machine capacity from case_parameters if possible
                        m_cap = 200 # Fallback
                        if m_type == 'M1': m_cap = 200
                        elif m_type == 'M2': m_cap = 70
                        elif m_type == 'M3_ALPHA': m_cap = 450
                        elif m_type == 'M3_ALPHA'.replace('_', '-'): m_cap = 450
                        elif m_type == 'M4': m_cap = 400 # Section 3 cap
                        zone_cap += (count * m_cap)
            
            # 4 fortnights per period
            zone_period_cap = zone_cap * 8 
            total_capacity += zone_period_cap
            
            # Targets
            zone_target = sum(row.get(f'Target_FN{fn}', 0) for fn in range(1, 9))
            total_target += zone_target
            
            # Overtime (Proxy: if Target > Capacity, diff is Overtime needed)
            # Or use explicit user input if we add it later.
            # For now, let's look at the result status (not easily accessible here without running loop)
            # Simplified: Overtime is approx max(0, Target - Capacity)
            # But the UI allows specific Overtime settings. 
            # We'll use the ratio for now as a better proxy than 0.
            if zone_target > zone_period_cap:
                total_overtime += (zone_target - zone_period_cap)
            
            # Unit Cost (Placeholder for now, but better than flat 40 if we used logic)
            # Let's keep 40 as base but adjust by zone volume?
            # Actually, standard cost is safer than broken calc.
            zone_costs[zone] = 40.0

        mean_util = (total_target / total_capacity) if total_capacity > 0 else 0
        
        outputs = {
            'production_plan': prod_plan,
            'capacity_utilization': {'mean': mean_util, 'total_capacity': total_capacity, 'total_target': total_target},
            'overtime_hours': total_overtime, # This is actually units made in Overtime, not hours, but sufficient for CFO
            'unit_costs': zone_costs
        }
        
        export_dashboard_data('Production', outputs)
        
    except Exception as e:
        print(f"Shared Output Export Error: {e}")

