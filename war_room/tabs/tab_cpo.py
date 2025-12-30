"""
ExSim War Room - CPO (People) Tab
Interactive grid for Hiring and Salaries.
Visualizations: Strike Risk, Payroll Pie.
Uses proper session state caching to prevent data loss.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

from utils.state_manager import get_state, set_state

ZONES = ['Center', 'West', 'North', 'East', 'South']
BENEFITS = [
    "Personal days (per person and period)",
    "Budget for additional training (% of payroll)",
    "Health and safety budget (% of payroll)",
    "Union representatives (total)",
    "Reduction of working hours (hours)",
    "Profit sharing (%)",
    "Health insurance (%)"
]


def init_cpo_state():
    """Initialize CPO state."""
    if 'cpo_initialized' not in st.session_state:
        st.session_state.cpo_initialized = True
        st.session_state.cpo_inflation = 3.0
        
        workers_data = get_state('workers_data')
        
        # Workforce grid
        data = []
        for zone in ZONES:
            current = workers_data.get('zones', {}).get(zone, {}).get('workers', 50) if workers_data else 50
            salary = workers_data.get('zones', {}).get(zone, {}).get('salary', 25) if workers_data else 25
            data.append({
                'Zone': zone,
                'Current_Staff': int(current),
                'Hire': 0,
                'Fire': 0,
                'Current_Salary': salary,
                'New_Salary': salary
            })
        st.session_state.cpo_workforce_df = pd.DataFrame(data)
        
        # Benefits grid
        benefits_data = [{'Benefit': b, 'Amount': 0} for b in BENEFITS]
        st.session_state.cpo_benefits_df = pd.DataFrame(benefits_data)


def render_cpo_tab():
    """Render the CPO (People) tab with subtabs."""
    init_cpo_state()
    
    st.header("👥 CPO Dashboard - People & Compensation")
    
    workers_data = get_state('workers_data')
    if workers_data:
        st.success("✅ Workers data loaded from upload")
    else:
        st.info("💡 Upload Workers Balance file for accurate staffing data")
    
    # SUBTABS
    subtab1, subtab2, subtab3 = st.tabs(["👥 Workforce", "🎁 Benefits", "📊 Analytics"])
    
    with subtab1:
        # Inflation rate
        inflation = st.number_input(
            "Inflation Rate (%)", 
            min_value=0.0, max_value=20.0, 
            value=st.session_state.cpo_inflation, 
            step=0.5,
            key='cpo_inflation_input'
        )
        st.session_state.cpo_inflation = inflation
        
        st.subheader("📋 Workforce & Salaries by Zone")
        st.caption("Edit Hire/Fire and New Salary. Strike risk calculated automatically.")
        
        # Add calculated columns
        display_df = st.session_state.cpo_workforce_df.copy()
        display_df['Inflation_Floor'] = round(display_df['Current_Salary'] * (1 + inflation / 100), 2)
        display_df['Strike_Risk'] = display_df.apply(
            lambda r: '🔴 High' if r['New_Salary'] < r['Inflation_Floor'] else '🟢 Low', axis=1
        )
        
        gb = GridOptionsBuilder.from_dataframe(display_df)
        gb.configure_column('Zone', editable=False, pinned='left')
        gb.configure_column('Current_Staff', editable=False)
        gb.configure_column('Hire', editable=True, type=['numericColumn'])
        gb.configure_column('Fire', editable=True, type=['numericColumn'])
        gb.configure_column('Current_Salary', editable=False)
        gb.configure_column('New_Salary', editable=True, type=['numericColumn'])
        gb.configure_column('Inflation_Floor', editable=False)
        gb.configure_column('Strike_Risk', editable=False)
        gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
        grid_options = gb.build()
        
        grid_response = AgGrid(
            display_df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.MODEL_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            fit_columns_on_grid_load=True,
            height=250,
            key='cpo_workforce_grid'
        )
        
        if grid_response.data is not None:
            updated = pd.DataFrame(grid_response.data)
            st.session_state.cpo_workforce_df = updated[['Zone', 'Current_Staff', 'Hire', 'Fire', 'Current_Salary', 'New_Salary']]
            
            # Calculate payroll
            total_payroll = 0
            for _, row in updated.iterrows():
                net_staff = row['Current_Staff'] + row['Hire'] - row['Fire']
                total_payroll += net_staff * row['New_Salary'] * 1000
                total_payroll += row['Hire'] * 5000  # Hiring cost
            
            set_state('TOTAL_PAYROLL_CASH', total_payroll)
        
        st.metric("Total Payroll Cost", f"${get_state('TOTAL_PAYROLL_CASH', 0):,.0f}")
    
    with subtab2:
        st.subheader("🎁 Benefits Package")
        
        gb = GridOptionsBuilder.from_dataframe(st.session_state.cpo_benefits_df)
        gb.configure_column('Benefit', editable=False, width=350)
        gb.configure_column('Amount', editable=True, type=['numericColumn'])
        gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
        grid_options = gb.build()
        
        ben_response = AgGrid(
            st.session_state.cpo_benefits_df,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.MODEL_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            fit_columns_on_grid_load=True,
            height=280,
            key='cpo_benefits_grid'
        )
        
        if ben_response.data is not None:
            st.session_state.cpo_benefits_df = pd.DataFrame(ben_response.data)
    
    with subtab3:
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("⚠️ Strike Risk Chart")
            
            df = st.session_state.cpo_workforce_df.copy()
            df['Inflation_Floor'] = df['Current_Salary'] * (1 + st.session_state.cpo_inflation / 100)
            
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=df['Zone'], y=df['New_Salary'], mode='lines+markers', name='Proposed Salary', line=dict(color='blue')))
            fig.add_trace(go.Scatter(x=df['Zone'], y=df['Inflation_Floor'], mode='lines+markers', name='Inflation Floor', line=dict(color='red', dash='dash')))
            fig.update_layout(title='Salary vs Inflation Floor', height=300, yaxis_title='$/hour')
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            st.subheader("💰 Payroll Distribution")
            
            payroll = get_state('TOTAL_PAYROLL_CASH', 100000)
            pie_data = pd.DataFrame({
                'Category': ['Base Pay', 'Benefits', 'Hiring Costs'],
                'Amount': [payroll * 0.7, payroll * 0.2, payroll * 0.1]
            })
            fig = px.pie(pie_data, values='Amount', names='Category', title='Payroll Breakdown')
            fig.update_layout(height=300)
            st.plotly_chart(fig, use_container_width=True)
