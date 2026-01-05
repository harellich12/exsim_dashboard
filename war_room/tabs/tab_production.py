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
    from case_parameters import COMMON
    ZONES = COMMON.get('ZONES', ['Center', 'West', 'North', 'East', 'South'])
    FORTNIGHTS = COMMON.get('FORTNIGHTS', list(range(1, 9)))
except ImportError:
    ZONES = ['Center', 'West', 'North', 'East', 'South']
    FORTNIGHTS = list(range(1, 9))

# Zone Colors
ZONE_COLORS = {
    'Center': '#1565C0',  # Blue
    'West': '#EF6C00',    # Orange
    'North': '#2E7D32',   # Green
    'East': '#F9A825',    # Yellow
    'South': '#6D4C41'    # Brown
}

# Production defaults
PROD_CONFIG = {
    'units_per_worker': 50,
    'units_per_machine': 100,
    'overtime_multiplier': 1.25,
    'module_capacity': 5  # machines per module
}


def init_production_state():
    """Initialize Production state with zone-specific data."""
    if 'production_initialized' not in st.session_state:
        st.session_state.production_initialized = True
        
        prod_data = get_state('production_data')
        workers_data = get_state('workers_data')
        
        # Zone calculators - production by zone by fortnight
        zone_data = []
        for zone in ZONES:
            workers = workers_data.get('zones', {}).get(zone, {}).get('workers', 50) if workers_data else 50
            machines = prod_data.get('zones', {}).get(zone, {}).get('machines', 10) if prod_data else (57 if zone == 'Center' else 0)
            modules = prod_data.get('zones', {}).get(zone, {}).get('modules', 5) if prod_data else (72 if zone == 'Center' else 0)
            materials = prod_data.get('zones', {}).get(zone, {}).get('materials', 1000) if prod_data else (1000 if zone == 'Center' else 0)
            
            zone_data.append({
                'Zone': zone,
                'Current_Machines': machines,
                'Current_Workers': workers,
                'Module_Slots': modules,
                'Local_Materials': materials,
                'Machine_Cap': machines * PROD_CONFIG['units_per_machine'],
                'Labor_Cap': workers * PROD_CONFIG['units_per_worker'],
                **{f'Target_FN{fn}': 0 for fn in FORTNIGHTS},
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
                'Transfer_Machines_In': 0,
                'Transfer_Machines_Out': 0
            })
        st.session_state.production_expansion = pd.DataFrame(expansion_data)


def sync_from_uploads():
    """Sync Production data from uploaded files."""
    prod_data = get_state('production_data')
    workers_data = get_state('workers_data')
    
    if prod_data and 'zones' in prod_data:
        for idx, row in st.session_state.production_zones.iterrows():
            zone = row['Zone']
            if zone in prod_data['zones']:
                machines = prod_data['zones'][zone].get('machines', row['Current_Machines'])
                st.session_state.production_zones.at[idx, 'Current_Machines'] = machines
                st.session_state.production_zones.at[idx, 'Machine_Cap'] = machines * PROD_CONFIG['units_per_machine']
    
    if workers_data and 'zones' in workers_data:
        for idx, row in st.session_state.production_zones.iterrows():
            zone = row['Zone']
            if zone in workers_data['zones']:
                workers = workers_data['zones'][zone].get('workers', row['Current_Workers'])
                st.session_state.production_zones.at[idx, 'Current_Workers'] = workers
                st.session_state.production_zones.at[idx, 'Labor_Cap'] = workers * PROD_CONFIG['units_per_worker']


def calculate_zone_production():
    """Calculate real output and alerts per zone per fortnight."""
    zones_df = st.session_state.production_zones
    exp_df = st.session_state.production_expansion
    
    results = []
    for idx, row in zones_df.iterrows():
        zone = row['Zone']
        exp_row = exp_df[exp_df['Zone'] == zone].iloc[0]
        
        # Effective capacity
        machines = row['Current_Machines'] + exp_row['Buy_Machines'] + exp_row['Transfer_Machines_In'] - exp_row['Transfer_Machines_Out']
        machine_cap = machines * PROD_CONFIG['units_per_machine']
        labor_cap = row['Labor_Cap']
        materials = row['Local_Materials']
        
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
    
    # Zone status
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Machines", zone_row['Current_Machines'])
    with col2:
        st.metric("Workers", zone_row['Current_Workers'])
    with col3:
        st.metric("Machine Cap", f"{zone_row['Machine_Cap']:,}")
    with col4:
        st.metric("Materials", f"{zone_row['Local_Materials']:,}")
    
    # Production schedule grid
    st.markdown("### 📅 Production Schedule")
    
    EDITABLE_STYLE = {'backgroundColor': '#E3F2FD', 'color': '#1565C0'}
    
    # Combined Grid for Inputs & Outputs with Client-Side Logic
    
    # Prepare data with hidden constraints for JS
    combined_data = []
    machine_cap = zone_row['Machine_Cap']
    labor_cap = zone_row['Labor_Cap']
    materials = zone_row['Local_Materials']
    
    # Calculate initial python values for first render
    results = calculate_zone_production()
    zone_results = results[results['Zone'] == selected_zone].iloc[0]
    
    for fn in FORTNIGHTS:
        combined_data.append({
            'Fortnight': f'FN{fn}',
            'Target': zone_row.get(f'Target_FN{fn}', 0),
            'Overtime': zone_row.get(f'Overtime_FN{fn}', False),
            'Machine_Cap': machine_cap,
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
            let labor_cap = Number(params.data.Labor_Cap);
            let materials = Number(params.data.Materials);
            
            let effective_cap = machine_cap;
            if (overtime) {
                effective_cap = Math.floor(machine_cap * 1.25);
            }
            
            return Math.min(target, effective_cap, labor_cap, materials);
        }
    """)
    
    # JS Logic for Alerts
    alert_getter = JsCode("""
        function(params) {
            let target = Number(params.data.Target) || 0;
            let overtime = params.data.Overtime;
            let machine_cap = Number(params.data.Machine_Cap);
            let labor_cap = Number(params.data.Labor_Cap);
            let materials = Number(params.data.Materials);
            
            let effective_cap = machine_cap;
            if (overtime) {
                effective_cap = Math.floor(machine_cap * 1.25);
            }
            
            if (target > materials) return '📦 SHIPMENT NEEDED';
            if (target > effective_cap) return '⚙️ Machine Limit';
            if (target > labor_cap) return '👷 Labor Limit';
            return '✅ OK';
        }
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
    
    grid_response = AgGrid(
        combined_df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=350,
        allow_unsafe_jscode=True,
        key=f'zone_calc_grid_{selected_zone}'
    )
    
    if grid_response.data is not None:
        updated = pd.DataFrame(grid_response.data)
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
    
    zones_df = st.session_state.production_zones[['Zone', 'Current_Machines', 'Current_Workers', 
                                                  'Module_Slots', 'Local_Materials', 'Machine_Cap', 'Labor_Cap']]
    
    st.dataframe(zones_df, width='stretch', hide_index=True)
    
    # Section B: Expansion by Zone
    st.markdown("### Section B: Expansion by Zone")
    
    EDITABLE_STYLE = {'backgroundColor': '#E3F2FD', 'color': '#1565C0'}
    
    exp_df = st.session_state.production_expansion.copy()
    
    gb = GridOptionsBuilder.from_dataframe(exp_df)
    gb.configure_column('Zone', editable=False, width=80)
    gb.configure_column('Buy_Machines', headerName='Buy Machines', editable=True, width=120, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Buy_Modules', headerName='Buy Modules', editable=True, width=110, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Transfer_Machines_In', headerName='Transfer In', editable=True, width=110, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Transfer_Machines_Out', headerName='Transfer Out', editable=True, width=110, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
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
        
        modules = zone_row['Module_Slots']
        machines = zone_row['Current_Machines'] + exp_row['Buy_Machines']
        
        if modules == 0 and machines > 0:
            st.warning(f"⚠️ **{zone}**: Buy module first - zone has no slots!")
        elif machines > modules * PROD_CONFIG['module_capacity']:
            st.error(f"🔴 **{zone}**: Too many machines for slots! Buy {((machines // PROD_CONFIG['module_capacity']) + 1 - modules)} more modules")
    
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
    
    data_status = []
    if prod_data:
        data_status.append("✅ Production Data")
    if workers_data:
        data_status.append("✅ Workers Data")
    
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
