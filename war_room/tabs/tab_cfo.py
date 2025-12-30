"""
ExSim War Room - CFO (Finance) Tab
Aggregator dashboard with Credit/Mortgages/Dividends.
Visualizations: Liquidity Waterfall, Solvency Gauge.
Uses proper session state caching to prevent data loss.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

from utils.state_manager import get_state, set_state, get_summary_metrics

FORTNIGHTS = list(range(1, 9))


def init_cfo_state():
    """Initialize CFO state."""
    if 'cfo_initialized' not in st.session_state:
        st.session_state.cfo_initialized = True
        
        balance = get_state('balance_data')
        st.session_state.cfo_opening_cash = balance.get('net_sales', 500000) if balance else 500000
        st.session_state.cfo_total_assets = balance.get('total_assets', 4000000) if balance else 4000000
        st.session_state.cfo_total_liabilities = balance.get('total_liabilities', 1500000) if balance else 1500000
        st.session_state.cfo_dividends = 0
        
        # Credit lines grid
        credit_data = [{'Item': 'Credit Lines'}]
        for fn in FORTNIGHTS:
            credit_data[0][f'FN{fn}'] = 0
        st.session_state.cfo_credit_df = pd.DataFrame(credit_data)
        
        # Investments grid
        invest_data = [{'Item': 'Investments'}]
        for fn in FORTNIGHTS:
            invest_data[0][f'FN{fn}'] = 0
        st.session_state.cfo_invest_df = pd.DataFrame(invest_data)
        
        # Mortgages grid
        mortgage_data = [
            {'Loan': 'Loan 1', 'Amount': 0, 'Rate': 0.08, 'Payment1': 0, 'Payment2': 0},
            {'Loan': 'Loan 2', 'Amount': 0, 'Rate': 0.08, 'Payment1': 0, 'Payment2': 0}
        ]
        st.session_state.cfo_mortgage_df = pd.DataFrame(mortgage_data)


def render_cfo_tab():
    """Render the CFO (Finance) tab with subtabs."""
    init_cfo_state()
    
    st.header("💰 CFO Dashboard - Financial Control")
    
    balance = get_state('balance_data')
    if balance:
        st.success("✅ Balance data loaded from upload")
    else:
        st.info("💡 Upload Balance Statements for accurate financials")
    
    # Cross-tab summary
    metrics = get_summary_metrics()
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("Procurement", f"${metrics['Procurement']:,.0f}")
    with col2:
        st.metric("Logistics", f"${metrics['Logistics']:,.0f}")
    with col3:
        st.metric("Payroll", f"${metrics['Payroll']:,.0f}")
    with col4:
        st.metric("ESG", f"${metrics['ESG CapEx']:,.0f}")
    with col5:
        st.metric("Total Out", f"${metrics['Total Out']:,.0f}", delta_color="inverse")
    
    # SUBTABS
    subtab1, subtab2, subtab3 = st.tabs(["💳 Financing", "📊 Cash Flow", "📈 Solvency"])
    
    with subtab1:
        col1, col2 = st.columns(2)
        with col1:
            opening = st.number_input("Opening Cash", value=int(st.session_state.cfo_opening_cash), step=10000, key='cfo_open')
            st.session_state.cfo_opening_cash = opening
            assets = st.number_input("Total Assets", value=int(st.session_state.cfo_total_assets), step=100000, key='cfo_assets')
            st.session_state.cfo_total_assets = assets
        with col2:
            liabilities = st.number_input("Total Liabilities", value=int(st.session_state.cfo_total_liabilities), step=100000, key='cfo_liab')
            st.session_state.cfo_total_liabilities = liabilities
            dividends = st.number_input("Dividends", value=int(st.session_state.cfo_dividends), step=10000, key='cfo_div')
            st.session_state.cfo_dividends = dividends
        
        st.markdown("---")
        st.subheader("💳 Credit Lines by Fortnight")
        
        gb = GridOptionsBuilder.from_dataframe(st.session_state.cfo_credit_df)
        gb.configure_column('Item', editable=False)
        for fn in FORTNIGHTS:
            gb.configure_column(f'FN{fn}', editable=True, type=['numericColumn'], width=80)
        gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
        grid_options = gb.build()
        
        credit_response = AgGrid(
            st.session_state.cfo_credit_df, gridOptions=grid_options,
            update_mode=GridUpdateMode.MODEL_CHANGED, data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            fit_columns_on_grid_load=True, height=80, key='cfo_credit_grid'
        )
        if credit_response.data is not None:
            st.session_state.cfo_credit_df = pd.DataFrame(credit_response.data)
        
        st.subheader("🏠 Mortgages")
        gb = GridOptionsBuilder.from_dataframe(st.session_state.cfo_mortgage_df)
        gb.configure_column('Loan', editable=False)
        gb.configure_column('Amount', editable=True, type=['numericColumn'])
        gb.configure_column('Rate', editable=False)
        gb.configure_column('Payment1', editable=True, type=['numericColumn'])
        gb.configure_column('Payment2', editable=True, type=['numericColumn'])
        gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
        
        mort_response = AgGrid(
            st.session_state.cfo_mortgage_df, gridOptions=gb.build(),
            update_mode=GridUpdateMode.MODEL_CHANGED, data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            fit_columns_on_grid_load=True, height=100, key='cfo_mort_grid'
        )
        if mort_response.data is not None:
            st.session_state.cfo_mortgage_df = pd.DataFrame(mort_response.data)
    
    with subtab2:
        st.subheader("💧 Liquidity Waterfall")
        
        credit_total = sum(st.session_state.cfo_credit_df.iloc[0].get(f'FN{fn}', 0) for fn in FORTNIGHTS)
        cash_in = st.session_state.cfo_opening_cash + credit_total
        cash_out = metrics['Total Out'] + st.session_state.cfo_dividends
        closing = cash_in - cash_out
        
        fig = go.Figure(go.Waterfall(
            measure=["absolute", "relative", "relative", "relative", "relative", "total"],
            x=["Opening", "Credit", "Procurement", "Operations", "Dividends+ESG", "Closing"],
            y=[
                st.session_state.cfo_opening_cash,
                credit_total,
                -metrics['Procurement'],
                -(metrics['Logistics'] + metrics['Payroll']),
                -(metrics['ESG CapEx'] + st.session_state.cfo_dividends),
                closing
            ],
            increasing={"marker": {"color": "green"}},
            decreasing={"marker": {"color": "red"}},
            totals={"marker": {"color": "blue"}}
        ))
        fig.update_layout(title="Cash Flow Waterfall", height=400)
        st.plotly_chart(fig, use_container_width=True)
        
        if closing > 0:
            st.success(f"✅ Projected Closing Cash: ${closing:,.0f}")
        else:
            st.error(f"❌ Cash Shortfall: ${closing:,.0f}")
    
    with subtab3:
        st.subheader("📏 Solvency Analysis")
        
        assets = st.session_state.cfo_total_assets
        liabilities = st.session_state.cfo_total_liabilities
        new_debt = st.session_state.cfo_mortgage_df['Amount'].sum()
        
        current_ratio = liabilities / assets if assets > 0 else 0
        post_ratio = (liabilities + new_debt) / assets if assets > 0 else 0
        
        fig = go.Figure()
        fig.add_trace(go.Bar(x=['Current', 'Post-Decision'], y=[current_ratio, post_ratio],
                             marker_color=['steelblue', 'orange' if post_ratio > 0.6 else 'steelblue']))
        fig.add_hline(y=0.6, line_dash="dash", line_color="red", annotation_text="60% Risk Threshold")
        fig.update_layout(title="Debt Ratio", yaxis_tickformat=".0%", height=350)
        st.plotly_chart(fig, use_container_width=True)
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Current Debt Ratio", f"{current_ratio:.1%}")
        with col2:
            st.metric("Post-Decision Ratio", f"{post_ratio:.1%}", delta=f"{(post_ratio - current_ratio):.1%}")
