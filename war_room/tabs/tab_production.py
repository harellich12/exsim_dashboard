"""
ExSim War Room - Production Tab
Interactive grid for Target Production per Zone.
Visualization: Capacity Stack (Combo Chart).
Uses proper session state caching to prevent data loss.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

from utils.state_manager import get_state, set_state

ZONES = ['Center', 'West', 'North', 'East', 'South']


def init_production_state():
    """Initialize Production state with defaults or from uploaded data."""
    if 'production_initialized' not in st.session_state:
        st.session_state.production_initialized = True
        
        # Get data from uploads
        prod_data = get_state('production_data')
        workers_data = get_state('workers_data')
        
        data = []
        for zone in ZONES:
            machine_cap = prod_data.get('machine_capacity', 5000) if prod_data else 5000
            labor_cap = workers_data.get('zones', {}).get(zone, {}).get('workers', 50) * 50 if workers_data else 2500
            
            data.append({
                'Zone': zone,
                'Target_Production': 0,
                'Machine_Capacity': int(machine_cap),
                'Labor_Capacity': int(labor_cap),
                'Bottleneck': '✅ OK'
            })
        
        st.session_state.production_df = pd.DataFrame(data)


def render_production_tab():
    """Render the Production tab with subtabs."""
    init_production_state()
    
    st.header("🏭 Production Dashboard - Capacity Planning")
    
    # Load data status
    prod_data = get_state('production_data')
    workers_data = get_state('workers_data')
    
    if prod_data or workers_data:
        st.success("✅ Production/Workers data loaded from uploads")
    else:
        st.info("💡 Upload Production and Workers files for accurate capacity data")
    
    # SUBTABS
    subtab1, subtab2 = st.tabs(["📋 Production Planning", "📊 Capacity Analysis"])
    
    with subtab1:
        st.subheader("📋 Production Targets by Zone")
        st.caption("Edit Target Production. Bottleneck status updates automatically.")
        
        gb = GridOptionsBuilder.from_dataframe(st.session_state.production_df)
        gb.configure_column('Zone', editable=False, pinned='left')
        gb.configure_column('Target_Production', editable=True, type=['numericColumn'])
        gb.configure_column('Machine_Capacity', editable=False)
        gb.configure_column('Labor_Capacity', editable=False)
        gb.configure_column('Bottleneck', editable=False)
        gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
        grid_options = gb.build()
        
        grid_response = AgGrid(
            st.session_state.production_df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.MODEL_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            fit_columns_on_grid_load=True,
            height=250,
            key='production_grid'
        )
        
        # Update state and calculate bottlenecks
        if grid_response.data is not None:
            updated_df = pd.DataFrame(grid_response.data)
            
            for idx, row in updated_df.iterrows():
                target = row['Target_Production']
                machine = row['Machine_Capacity']
                labor = row['Labor_Capacity']
                min_cap = min(machine, labor)
                
                if target > min_cap:
                    if machine < labor:
                        updated_df.at[idx, 'Bottleneck'] = '⚠️ Machine'
                    else:
                        updated_df.at[idx, 'Bottleneck'] = '⚠️ Labor'
                else:
                    updated_df.at[idx, 'Bottleneck'] = '✅ OK'
            
            st.session_state.production_df = updated_df
            set_state('PRODUCTION_PLAN', updated_df)
        
        # Summary metrics
        col1, col2, col3 = st.columns(3)
        with col1:
            total_target = st.session_state.production_df['Target_Production'].sum()
            st.metric("Total Target", f"{total_target:,.0f} units")
        with col2:
            bottlenecks = st.session_state.production_df['Bottleneck'].str.contains('⚠️').sum()
            st.metric("Bottlenecks", bottlenecks, delta_color="inverse")
        with col3:
            total_machine = st.session_state.production_df['Machine_Capacity'].sum()
            utilization = total_target / total_machine * 100 if total_machine > 0 else 0
            st.metric("Utilization", f"{utilization:.1f}%")
    
    with subtab2:
        st.subheader("📊 Capacity Stack Chart")
        
        fig = go.Figure()
        
        # Stacked bars for capacity
        fig.add_trace(go.Bar(
            name='Machine Capacity',
            x=st.session_state.production_df['Zone'],
            y=st.session_state.production_df['Machine_Capacity'],
            marker_color='steelblue'
        ))
        
        fig.add_trace(go.Bar(
            name='Labor Capacity',
            x=st.session_state.production_df['Zone'],
            y=st.session_state.production_df['Labor_Capacity'],
            marker_color='lightblue'
        ))
        
        # Line for target
        fig.add_trace(go.Scatter(
            name='Target Production',
            x=st.session_state.production_df['Zone'],
            y=st.session_state.production_df['Target_Production'],
            mode='lines+markers',
            line=dict(color='red', width=3),
            marker=dict(size=10)
        ))
        
        fig.update_layout(
            barmode='group',
            title='Capacity vs Target Production',
            xaxis_title='Zone',
            yaxis_title='Units',
            height=400
        )
        
        st.plotly_chart(fig, use_container_width=True)
