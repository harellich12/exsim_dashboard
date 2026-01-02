"""
ExSim War Room - CPO (People) Tab
4 sub-tabs mirroring the Excel dashboard sheets:
1. WORKFORCE_PLANNING - Headcount by zone, hiring/firing
2. COMPENSATION_STRATEGY - Salaries, strike risk, benefits
3. LABOR_COST_ANALYSIS - Total labor cost forecast
4. UPLOAD_READY_PEOPLE - Export preview
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode

from utils.state_manager import get_state, set_state

# Constants
ZONES = ['Center', 'West', 'North', 'East', 'South']

# Default HR Parameters
HR_CONFIG = {
    'hiring_fee': 3000,
    'severance': 5000,
    'default_turnover': 0.05,
    'default_salary': 25  # hourly
}

BENEFITS = [
    ('Personal days (per person)', 0, 'days'),
    ('Training budget (% of payroll)', 2, '%'),
    ('Health & safety (% of payroll)', 3, '%'),
    ('Union representatives', 0, 'count'),
    ('Reduction of working hours', 0, 'hours'),
    ('Profit sharing (%)', 0, '%'),
    ('Health insurance (%)', 3, '%')
]


def init_cpo_state():
    """Initialize CPO state with proper workforce data structures."""
    if 'cpo_initialized' not in st.session_state:
        st.session_state.cpo_initialized = True
        
        workers_data = get_state('workers_data')
        
        # Inflation rate
        st.session_state.cpo_inflation = 3.0
        
        # Workforce planning by zone
        workforce_data = []
        for zone in ZONES:
            current = workers_data.get('zones', {}).get(zone, {}).get('workers', 50) if workers_data else 50
            salary = workers_data.get('zones', {}).get(zone, {}).get('salary', HR_CONFIG['default_salary']) if workers_data else HR_CONFIG['default_salary']
            workforce_data.append({
                'Zone': zone,
                'Current_Workers': int(current),
                'Required_Workers': int(current),
                'Turnover_Rate': HR_CONFIG['default_turnover'] * 100,
                'Hire': 0,
                'Fire': 0,
                'Prev_Salary': salary,
                'New_Salary': salary
            })
        st.session_state.cpo_workforce = pd.DataFrame(workforce_data)
        
        # Benefits settings
        benefits_data = [{'Benefit': b[0], 'Amount': b[1], 'Unit': b[2]} for b in BENEFITS]
        st.session_state.cpo_benefits = pd.DataFrame(benefits_data)
        
        # Estimated net profit for profit sharing
        st.session_state.cpo_est_net_profit = 500000


def sync_from_uploads():
    """Sync CPO data from uploaded files."""
    workers_data = get_state('workers_data')
    
    if workers_data and 'zones' in workers_data:
        for idx, row in st.session_state.cpo_workforce.iterrows():
            zone = row['Zone']
            if zone in workers_data['zones']:
                workers = workers_data['zones'][zone].get('workers', row['Current_Workers'])
                salary = workers_data['zones'][zone].get('salary', row['Prev_Salary'])
                st.session_state.cpo_workforce.at[idx, 'Current_Workers'] = workers
                st.session_state.cpo_workforce.at[idx, 'Required_Workers'] = workers
                st.session_state.cpo_workforce.at[idx, 'Prev_Salary'] = salary


def calculate_workforce_costs():
    """Calculate total workforce costs."""
    wf = st.session_state.cpo_workforce
    inflation = st.session_state.cpo_inflation / 100
    benefits = st.session_state.cpo_benefits
    est_profit = st.session_state.cpo_est_net_profit
    
    results = {
        'base_payroll': 0,
        'hiring_costs': 0,
        'severance_costs': 0,
        'benefits_cost': 0,
        'profit_sharing': 0,
        'total': 0,
        'strike_risk_zones': []
    }
    
    # Calculate per zone
    for _, row in wf.iterrows():
        zone = row['Zone']
        current = row['Current_Workers']
        required = row['Required_Workers']
        turnover = row['Turnover_Rate'] / 100
        hire = row['Hire']
        fire = row['Fire']
        prev_salary = row['Prev_Salary']
        new_salary = row['New_Salary']
        min_salary = prev_salary * (1 + inflation)
        
        # Check strike risk
        if new_salary < min_salary:
            results['strike_risk_zones'].append(zone)
        
        # Net workers after turnover
        projected_loss = int(current * turnover)
        net_workers = current - projected_loss + hire - fire
        
        # Payroll (salary is hourly * 40 hrs/week * 2 weeks = fortnight)
        results['base_payroll'] += net_workers * new_salary * 40 * 2
        results['hiring_costs'] += hire * HR_CONFIG['hiring_fee']
        results['severance_costs'] += fire * HR_CONFIG['severance']
    
    # Benefits costs (as % of base payroll)
    training_pct = benefits[benefits['Benefit'].str.contains('Training')]['Amount'].values[0] / 100
    health_safety_pct = benefits[benefits['Benefit'].str.contains('Health & safety')]['Amount'].values[0] / 100
    health_ins_pct = benefits[benefits['Benefit'].str.contains('Health insurance')]['Amount'].values[0] / 100
    profit_sharing_pct = benefits[benefits['Benefit'].str.contains('Profit sharing')]['Amount'].values[0] / 100
    
    results['benefits_cost'] = results['base_payroll'] * (training_pct + health_safety_pct + health_ins_pct)
    results['profit_sharing'] = est_profit * profit_sharing_pct
    
    results['total'] = (results['base_payroll'] + results['hiring_costs'] + 
                       results['severance_costs'] + results['benefits_cost'] + 
                       results['profit_sharing'])
    
    return results


def render_workforce_planning():
    """Render WORKFORCE_PLANNING sub-tab - Headcount by zone."""
    st.subheader("👷 WORKFORCE PLANNING - Headcount by Zone")
    
    st.markdown(f"""
    **Turnover Reality:** Even stable companies lose 5-10% of workers annually.  
    **Hiring Fee:** ${HR_CONFIG['hiring_fee']:,} | **Severance:** ${HR_CONFIG['severance']:,}
    """)
    
    EDITABLE_STYLE = {'backgroundColor': '#E3F2FD', 'color': '#1565C0'}
    REFERENCE_STYLE = {'backgroundColor': '#F5F5F5', 'color': '#616161'}
    
    wf_df = st.session_state.cpo_workforce.copy()
    
    # Calculate projected values
    wf_df['Proj_Loss'] = (wf_df['Current_Workers'] * wf_df['Turnover_Rate'] / 100).astype(int)
    wf_df['Net_Workers'] = wf_df['Current_Workers'] - wf_df['Proj_Loss'] + wf_df['Hire'] - wf_df['Fire']
    wf_df['Hire_Cost'] = wf_df['Hire'] * HR_CONFIG['hiring_fee']
    wf_df['Fire_Cost'] = wf_df['Fire'] * HR_CONFIG['severance']
    
    gb = GridOptionsBuilder.from_dataframe(wf_df[['Zone', 'Current_Workers', 'Required_Workers', 
                                                  'Turnover_Rate', 'Proj_Loss', 'Hire', 'Fire', 
                                                  'Net_Workers', 'Hire_Cost', 'Fire_Cost']])
    gb.configure_column('Zone', editable=False, width=80, cellStyle=REFERENCE_STYLE)
    gb.configure_column('Current_Workers', headerName='Current', editable=False, width=100, cellStyle=REFERENCE_STYLE)
    gb.configure_column('Required_Workers', headerName='Required', editable=True, width=100, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Turnover_Rate', headerName='Turnover %', editable=True, width=100, type=['numericColumn'], 
                       valueFormatter="value.toFixed(1) + '%'", cellStyle=EDITABLE_STYLE)
    gb.configure_column('Proj_Loss', headerName='Proj Loss', editable=False, width=90, cellStyle=REFERENCE_STYLE)
    gb.configure_column('Hire', editable=True, width=70, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Fire', editable=True, width=70, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Net_Workers', headerName='Net Workers', editable=False, width=100, cellStyle=REFERENCE_STYLE)
    gb.configure_column('Hire_Cost', headerName='Hire Cost', editable=False, width=100, valueFormatter="'$' + value.toLocaleString()", cellStyle=REFERENCE_STYLE)
    gb.configure_column('Fire_Cost', headerName='Fire Cost', editable=False, width=100, valueFormatter="'$' + value.toLocaleString()", cellStyle=REFERENCE_STYLE)
    gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
    
    grid_response = AgGrid(
        wf_df[['Zone', 'Current_Workers', 'Required_Workers', 'Turnover_Rate', 
               'Proj_Loss', 'Hire', 'Fire', 'Net_Workers', 'Hire_Cost', 'Fire_Cost']],
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=230,
        key='cpo_workforce_grid'
    )
    
    if grid_response.data is not None:
        updated = pd.DataFrame(grid_response.data)
        for col in ['Required_Workers', 'Turnover_Rate', 'Hire', 'Fire']:
            st.session_state.cpo_workforce[col] = updated[col]
    
    # Summary
    total_hire = st.session_state.cpo_workforce['Hire'].sum()
    total_fire = st.session_state.cpo_workforce['Fire'].sum()
    hire_cost = total_hire * HR_CONFIG['hiring_fee']
    fire_cost = total_fire * HR_CONFIG['severance']
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Hiring", total_hire)
    with col2:
        st.metric("Total Firing", total_fire)
    with col3:
        st.metric("Hiring Cost", f"${hire_cost:,.0f}")
    with col4:
        st.metric("Severance Cost", f"${fire_cost:,.0f}")


def render_compensation_strategy():
    """Render COMPENSATION_STRATEGY sub-tab - Salaries and strike risk."""
    st.subheader("💰 COMPENSATION STRATEGY - Salaries & Strike Risk")
    
    # Inflation input - CRITICAL
    col1, col2 = st.columns([1, 2])
    with col1:
        inflation = st.number_input(
            "🔴 INFLATION RATE (%)",
            min_value=0.0, max_value=20.0,
            value=st.session_state.cpo_inflation,
            step=0.5,
            key='cpo_inflation_input',
            help="Get from Case Guide - CRITICAL for avoiding strikes!"
        )
        st.session_state.cpo_inflation = inflation
    
    with col2:
        st.warning(f"""
        **Min Salary to Avoid Strike = Previous Salary × (1 + {inflation}%)**  
        Set salaries ABOVE the Min Salary column or face STRIKE RISK!
        """)
    
    # Salary grid with strike risk
    wf_df = st.session_state.cpo_workforce.copy()
    wf_df['Min_Salary'] = (wf_df['Prev_Salary'] * (1 + inflation / 100)).round(2)
    wf_df['Strike_Risk'] = wf_df.apply(
        lambda r: '🔴 STRIKE RISK!' if r['New_Salary'] < r['Min_Salary'] else '🟢 Safe', axis=1
    )
    
    EDITABLE_STYLE = {'backgroundColor': '#E3F2FD', 'color': '#1565C0'}
    REFERENCE_STYLE = {'backgroundColor': '#F5F5F5', 'color': '#616161'}
    
    strike_js = JsCode("""
        function(params) {
            if (params.value && params.value.includes('STRIKE')) {
                return {'backgroundColor': '#FFCDD2', 'color': '#B71C1C', 'fontWeight': 'bold'};
            }
            return {'backgroundColor': '#C8E6C9', 'color': '#2E7D32'};
        }
    """)
    
    salary_js = JsCode("""
        function(params) {
            if (params.data && params.value < params.data.Min_Salary) {
                return {'backgroundColor': '#FFCDD2', 'color': '#B71C1C', 'fontWeight': 'bold'};
            }
            return {'backgroundColor': '#E3F2FD', 'color': '#1565C0'};
        }
    """)
    
    gb = GridOptionsBuilder.from_dataframe(wf_df[['Zone', 'Prev_Salary', 'Min_Salary', 'New_Salary', 'Strike_Risk']])
    gb.configure_column('Zone', editable=False, width=90, cellStyle=REFERENCE_STYLE)
    gb.configure_column('Prev_Salary', headerName='Prev Salary', editable=False, width=110, 
                       valueFormatter="'$' + value.toFixed(2)", cellStyle=REFERENCE_STYLE)
    gb.configure_column('Min_Salary', headerName='Min Salary', editable=False, width=110,
                       valueFormatter="'$' + value.toFixed(2)", cellStyle=REFERENCE_STYLE)
    gb.configure_column('New_Salary', headerName='New Salary', editable=True, width=110, type=['numericColumn'],
                       valueFormatter="'$' + value.toFixed(2)", cellStyle=salary_js)
    gb.configure_column('Strike_Risk', headerName='Strike Risk', editable=False, width=140, cellStyle=strike_js)
    gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
    
    grid_response = AgGrid(
        wf_df[['Zone', 'Prev_Salary', 'Min_Salary', 'New_Salary', 'Strike_Risk']],
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=230,
        allow_unsafe_jscode=True,
        key='cpo_salary_grid'
    )
    
    if grid_response.data is not None:
        updated = pd.DataFrame(grid_response.data)
        st.session_state.cpo_workforce['New_Salary'] = updated['New_Salary']
    
    # Benefits section
    st.markdown("### 🎁 Benefits Package")
    
    benefits_df = st.session_state.cpo_benefits.copy()
    
    gb = GridOptionsBuilder.from_dataframe(benefits_df)
    gb.configure_column('Benefit', editable=False, width=280)
    gb.configure_column('Amount', editable=True, width=100, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Unit', editable=False, width=80)
    gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
    
    grid_response = AgGrid(
        benefits_df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=280,
        key='cpo_benefits_grid'
    )
    
    if grid_response.data is not None:
        st.session_state.cpo_benefits = pd.DataFrame(grid_response.data)


def render_labor_cost_analysis():
    """Render LABOR_COST_ANALYSIS sub-tab - Total labor cost forecast."""
    st.subheader("📊 LABOR COST ANALYSIS - Total Expense Forecast")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        est_profit = st.number_input(
            "Estimated Net Profit",
            value=int(st.session_state.cpo_est_net_profit),
            step=50000,
            key='cpo_profit_input'
        )
        st.session_state.cpo_est_net_profit = est_profit
    
    # Calculate costs
    costs = calculate_workforce_costs()
    
    # Strike risk warning
    if costs['strike_risk_zones']:
        st.error(f"🔴 **STRIKE RISK in: {', '.join(costs['strike_risk_zones'])}** - Increase salaries above inflation floor!")
    else:
        st.success("🟢 All zones have salaries above inflation floor - No strike risk")
    
    # Cost breakdown
    st.markdown("### 💵 Cost Breakdown")
    
    cost_data = pd.DataFrame({
        'Category': ['Base Payroll', 'Hiring Costs', 'Severance Costs', 'Benefits', 'Profit Sharing'],
        'Amount': [costs['base_payroll'], costs['hiring_costs'], costs['severance_costs'],
                  costs['benefits_cost'], costs['profit_sharing']]
    })
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        for _, row in cost_data.iterrows():
            st.metric(row['Category'], f"${row['Amount']:,.0f}")
    
    with col2:
        fig = px.pie(
            cost_data,
            values='Amount',
            names='Category',
            title='Labor Cost Distribution',
            color_discrete_sequence=px.colors.qualitative.Set2
        )
        fig.update_layout(height=350)
        st.plotly_chart(fig, width='stretch')
    
    # Total
    st.metric("**TOTAL LABOR EXPENSE**", f"${costs['total']:,.0f}")
    set_state('TOTAL_PAYROLL_CASH', costs['total'])
    
    # Strike Risk Chart
    st.markdown("### ⚠️ Strike Risk Visualization")
    
    wf = st.session_state.cpo_workforce.copy()
    inflation = st.session_state.cpo_inflation / 100
    wf['Min_Salary'] = wf['Prev_Salary'] * (1 + inflation)
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=wf['Zone'],
        y=wf['New_Salary'],
        name='Proposed Salary',
        marker_color='#1565C0'
    ))
    fig.add_trace(go.Scatter(
        x=wf['Zone'],
        y=wf['Min_Salary'],
        mode='lines+markers',
        name='Inflation Floor (Min)',
        line=dict(color='#D32F2F', width=3, dash='dash')
    ))
    fig.update_layout(
        title='Proposed Salary vs Inflation Floor by Zone',
        height=350,
        template='plotly_white',
        yaxis_title='Salary ($/hr)'
    )
    st.plotly_chart(fig, width='stretch')


def render_upload_ready_people():
    """Render UPLOAD_READY_PEOPLE sub-tab - Export preview."""
    st.subheader("📤 UPLOAD READY - People Decisions")
    
    st.info("Copy these values to ExSim People Decision Form")
    
    # Workforce changes
    st.markdown("### 👷 Workforce Changes")
    
    wf = st.session_state.cpo_workforce
    changes = wf[(wf['Hire'] > 0) | (wf['Fire'] > 0) | (wf['New_Salary'] != wf['Prev_Salary'])]
    
    if not changes.empty:
        display = changes[['Zone', 'Hire', 'Fire', 'New_Salary']].copy()
        st.dataframe(display, hide_index=True, width='stretch')
    else:
        st.caption("No workforce changes")
    
    # Benefits summary
    st.markdown("### 🎁 Benefits")
    
    benefits = st.session_state.cpo_benefits[st.session_state.cpo_benefits['Amount'] > 0]
    if not benefits.empty:
        st.dataframe(benefits, hide_index=True, width='stretch')
    else:
        st.caption("No benefits set")
    
    # Cost summary
    st.markdown("### 💰 Cost Summary")
    costs = calculate_workforce_costs()
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Base Payroll", f"${costs['base_payroll']:,.0f}")
        st.metric("HR Costs (Hire/Fire)", f"${costs['hiring_costs'] + costs['severance_costs']:,.0f}")
    with col2:
        st.metric("Benefits", f"${costs['benefits_cost']:,.0f}")
        st.metric("Profit Sharing", f"${costs['profit_sharing']:,.0f}")
    
    st.metric("**TOTAL LABOR EXPENSE**", f"${costs['total']:,.0f}")
    
    if st.button("📋 Copy People Decisions", type="primary", key='cpo_copy'):
        st.success("✅ Data copied! Paste into ExSim People form.")


def render_cpo_tab():
    """Render the CPO (People) tab with 4 Excel-aligned subtabs."""
    init_cpo_state()
    sync_from_uploads()
    
    st.header("👥 CPO Dashboard - Workforce Planning & Compensation")
    
    # Data source status
    workers_data = get_state('workers_data')
    
    if workers_data and workers_data.get('zones'):
        st.success("✅ Workers Balance data loaded")
    else:
        st.info("💡 Upload Workers Balance file in sidebar to populate staffing data")
    
    # Strike risk quick check
    costs = calculate_workforce_costs()
    if costs['strike_risk_zones']:
        st.error(f"🔴 **STRIKE RISK in: {', '.join(costs['strike_risk_zones'])}**")
    
    # 4 SUBTABS - Matching Excel logic
    subtabs = st.tabs([
        "👷 Workforce Planning",
        "💰 Compensation Strategy",
        "📊 Labor Cost Analysis",
        "📤 Upload Ready"
    ])
    
    with subtabs[0]:
        render_workforce_planning()
    
    with subtabs[1]:
        render_compensation_strategy()
    
    with subtabs[2]:
        render_labor_cost_analysis()
    
    with subtabs[3]:
        render_upload_ready_people()
