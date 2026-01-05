"""
ExSim War Room - CFO (Finance) Tab
5 sub-tabs mirroring the Excel dashboard sheets:
1. LIQUIDITY_MONITOR - Cash flow by fortnight
2. PROFIT_CONTROL - Income statement projection vs actuals
3. BALANCE_SHEET_HEALTH - Debt ratio tracking
4. DEBT_MANAGER - Mortgage calculator
5. UPLOAD_READY_FINANCE - Export preview
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode

from utils.state_manager import get_state, set_state, get_summary_metrics

# Import centralized constants from case_parameters
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
try:
    from case_parameters import COMMON
    FORTNIGHTS = COMMON.get('FORTNIGHTS', list(range(1, 9)))
except ImportError:
    FORTNIGHTS = list(range(1, 9))

ST_LIMIT = 500000  # Short-term debt limit
LT_LIMIT = 2000000  # Long-term debt limit
DEFAULT_TAX_RATE = 0.25


def init_cfo_state():
    """Initialize CFO state with proper financial data structures."""
    if 'cfo_initialized' not in st.session_state:
        st.session_state.cfo_initialized = True
        
        # Balance data from upload
        balance = get_state('balance_data')
        
        # Section A: Initialization (from initial_cash_flow.xlsx)
        st.session_state.cfo_cash_at_end_last_period = 500000
        st.session_state.cfo_tax_payments = 50000
        st.session_state.cfo_dividend_payments = 0
        st.session_state.cfo_asset_purchases = 0
        
        # Section B: Operational Cash Flow (8 fortnights)
        st.session_state.cfo_cash_flow = pd.DataFrame({
            'Item': ['Sales Receipts', 'Procurement Spend', 'Fixed Overhead (S&A)', 
                    'Receivables (HARD)', 'Payables (HARD)'],
            **{f'FN{fn}': [100000, 0, 25000, 0, 0] for fn in FORTNIGHTS}
        })
        
        # Section C: Financing Decisions (8 fortnights)
        st.session_state.cfo_financing = pd.DataFrame({
            'Item': ['Credit Line (+/-)', 'Investments (+/-)', 
                    'Mortgage Inflow', 'Dividends Paid'],
            **{f'FN{fn}': [0, 0, 0, 0] for fn in FORTNIGHTS}
        })
        
        # Balance data
        if balance:
            st.session_state.cfo_net_sales = balance.get('net_sales', 1000000)
            st.session_state.cfo_cogs = balance.get('cogs', 600000)
            st.session_state.cfo_gross_margin = balance.get('gross_income', 400000)
            st.session_state.cfo_net_profit = balance.get('net_profit', 100000)
            st.session_state.cfo_total_assets = balance.get('total_assets', 4000000)
            st.session_state.cfo_total_liabilities = balance.get('total_liabilities', 1500000)
            st.session_state.cfo_retained_earnings = balance.get('retained_earnings', 200000)
            st.session_state.cfo_gross_margin_pct = balance.get('gross_margin_pct', 0.4)
            st.session_state.cfo_net_margin_pct = balance.get('net_margin_pct', 0.1)
        else:
            st.session_state.cfo_net_sales = 1000000
            st.session_state.cfo_cogs = 600000
            st.session_state.cfo_gross_margin = 400000
            st.session_state.cfo_net_profit = 100000
            st.session_state.cfo_total_assets = 4000000
            st.session_state.cfo_total_liabilities = 1500000
            st.session_state.cfo_retained_earnings = 200000
            st.session_state.cfo_gross_margin_pct = 0.4
            st.session_state.cfo_net_margin_pct = 0.1
        
        # Mortgages
        st.session_state.cfo_mortgages = pd.DataFrame({
            'Loan': ['Loan 1', 'Loan 2', 'Loan 3'],
            'Amount': [0, 0, 0],
            'Interest_Rate': [0.08, 0.08, 0.08],
            'Payment_FN1': [0, 0, 0],
            'Payment_FN2': [0, 0, 0]
        })


def sync_from_uploads():
    """Sync CFO data from uploaded files."""
    balance = get_state('balance_data')
    ar_ap_data = get_state('ar_ap_data')
    
    if balance:
        st.session_state.cfo_net_sales = balance.get('net_sales', st.session_state.cfo_net_sales)
        st.session_state.cfo_total_assets = balance.get('total_assets', st.session_state.cfo_total_assets)
        st.session_state.cfo_total_liabilities = balance.get('total_liabilities', st.session_state.cfo_total_liabilities)
    
    if ar_ap_data:
        # Update hard data rows in cash flow
        for fn in FORTNIGHTS:
            if ar_ap_data.get('receivables'):
                st.session_state.cfo_cash_flow.at[3, f'FN{fn}'] = ar_ap_data['receivables'][fn-1]
            if ar_ap_data.get('payables'):
                st.session_state.cfo_cash_flow.at[4, f'FN{fn}'] = -ar_ap_data['payables'][fn-1]


def calculate_cash_flow():
    """Calculate cash flow for each fortnight."""
    # Starting cash
    starting_cash = (st.session_state.cfo_cash_at_end_last_period 
                    - st.session_state.cfo_tax_payments 
                    - st.session_state.cfo_dividend_payments
                    - st.session_state.cfo_asset_purchases)
    
    results = []
    opening = starting_cash
    
    for fn in FORTNIGHTS:
        fn_col = f'FN{fn}'
        
        # Operational
        sales = st.session_state.cfo_cash_flow.at[0, fn_col]
        procurement = st.session_state.cfo_cash_flow.at[1, fn_col]
        overhead = st.session_state.cfo_cash_flow.at[2, fn_col]
        receivables = st.session_state.cfo_cash_flow.at[3, fn_col]
        payables = st.session_state.cfo_cash_flow.at[4, fn_col]
        
        # Financing
        credit = st.session_state.cfo_financing.at[0, fn_col]
        investments = st.session_state.cfo_financing.at[1, fn_col]
        mortgage = st.session_state.cfo_financing.at[2, fn_col]
        dividends = st.session_state.cfo_financing.at[3, fn_col]
        
        # Net = Inflows - Outflows
        net_flow = (sales + receivables + credit + mortgage 
                   - procurement - overhead - abs(payables) - investments - dividends)
        closing = opening + net_flow
        
        # Solvency check
        if closing < 0:
            status = "🔴 INSOLVENT"
        elif closing > 200000:
            status = "🟡 Excess Cash"
        else:
            status = "🟢 OK"
        
        results.append({
            'Fortnight': f'FN{fn}',
            'Opening': opening,
            'Net_Flow': net_flow,
            'Closing': closing,
            'Status': status
        })
        
        opening = closing
    
    return pd.DataFrame(results)


def render_liquidity_monitor():
    """Render LIQUIDITY_MONITOR sub-tab - Cash flow engine."""
    st.subheader("💧 LIQUIDITY MONITOR - Cash Flow Engine")
    
    # Section A: Initialization
    with st.expander("📋 Section A: Initialization (Initial Cash Flow Bridge)", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            cash_end = st.number_input(
                "Cash at End of Last Period",
                value=int(st.session_state.cfo_cash_at_end_last_period),
                step=10000,
                key='cfo_cash_end'
            )
            st.session_state.cfo_cash_at_end_last_period = cash_end
            
            tax = st.number_input(
                "Less: Tax Payments",
                value=int(st.session_state.cfo_tax_payments),
                step=5000,
                key='cfo_tax'
            )
            st.session_state.cfo_tax_payments = tax
        
        with col2:
            div = st.number_input(
                "Less: Dividend Payments",
                value=int(st.session_state.cfo_dividend_payments),
                step=5000,
                key='cfo_div_init'
            )
            st.session_state.cfo_dividend_payments = div
            
            assets = st.number_input(
                "Less: Asset Purchases",
                value=int(st.session_state.cfo_asset_purchases),
                step=10000,
                key='cfo_assets_purchase'
            )
            st.session_state.cfo_asset_purchases = assets
        
        starting_cash = cash_end - tax - div - assets
        st.metric("**STARTING CASH FOR FN1**", f"${starting_cash:,.0f}")
    
    # Section B: Operational Cash Flow
    st.markdown("### 📊 Section B: Operational Cash Flow")
    
    # Define colors for AgGrid
    EDITABLE_STYLE = {'backgroundColor': '#E3F2FD', 'color': '#1565C0'}
    REFERENCE_STYLE = {'backgroundColor': '#F5F5F5', 'color': '#616161'}
    HARD_DATA_STYLE = {'backgroundColor': '#B0BEC5', 'color': '#37474F'}
    
    cash_df = st.session_state.cfo_cash_flow.copy()
    
    gb = GridOptionsBuilder.from_dataframe(cash_df)
    gb.configure_column('Item', editable=False, pinned='left', width=180)
    
    for fn in FORTNIGHTS:
        gb.configure_column(f'FN{fn}', editable=True, width=100, 
                           type=['numericColumn'],
                           cellStyle=EDITABLE_STYLE)
    
    gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
    
    grid_response = AgGrid(
        cash_df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=200,
        key='cfo_cashflow_grid'
    )
    
    if grid_response.data is not None:
        st.session_state.cfo_cash_flow = pd.DataFrame(grid_response.data)
    
    # Section C: Financing Decisions
    st.markdown("### 💳 Section C: Financing Decisions")
    st.caption(f"📌 Short-Term Debt Limit: ${ST_LIMIT:,.0f}")
    
    financing_df = st.session_state.cfo_financing.copy()
    
    gb = GridOptionsBuilder.from_dataframe(financing_df)
    gb.configure_column('Item', editable=False, pinned='left', width=180)
    
    for fn in FORTNIGHTS:
        gb.configure_column(f'FN{fn}', editable=True, width=100,
                           type=['numericColumn'],
                           cellStyle=EDITABLE_STYLE)
    
    gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
    
    grid_response = AgGrid(
        financing_df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=170,
        key='cfo_financing_grid'
    )
    
    if grid_response.data is not None:
        st.session_state.cfo_financing = pd.DataFrame(grid_response.data)
    
    # Section D: Cash Balance
    st.markdown("### 📈 Section D: Cash Balance")
    
    results_df = calculate_cash_flow()
    
    # Solvency status styling
    status_js = JsCode("""
        function(params) {
            if (params.value && params.value.includes('INSOLVENT')) {
                return {'backgroundColor': '#FFCDD2', 'color': '#B71C1C', 'fontWeight': 'bold'};
            } else if (params.value && params.value.includes('Excess')) {
                return {'backgroundColor': '#FFF9C4', 'color': '#F57F17'};
            }
            return {'backgroundColor': '#C8E6C9', 'color': '#2E7D32'};
        }
    """)
    
    closing_js = JsCode("""
        function(params) {
            if (params.value < 0) {
                return {'backgroundColor': '#FFCDD2', 'color': '#B71C1C', 'fontWeight': 'bold'};
            } else if (params.value > 200000) {
                return {'backgroundColor': '#C8E6C9', 'color': '#1B5E20'};
            }
            return {'backgroundColor': '#E3F2FD', 'color': '#1565C0'};
        }
    """)
    
    gb = GridOptionsBuilder.from_dataframe(results_df)
    gb.configure_column('Fortnight', editable=False, width=80)
    gb.configure_column('Opening', editable=False, width=120, valueFormatter="'$' + value.toLocaleString()")
    gb.configure_column('Net_Flow', headerName='Net Flow', editable=False, width=120, valueFormatter="'$' + value.toLocaleString()")
    gb.configure_column('Closing', editable=False, width=120, valueFormatter="'$' + value.toLocaleString()", cellStyle=closing_js)
    gb.configure_column('Status', editable=False, width=130, cellStyle=status_js)
    
    AgGrid(
        results_df,
        gridOptions=gb.build(),
        fit_columns_on_grid_load=True,
        height=300,
        allow_unsafe_jscode=True,
        key='cfo_results_grid'
    )
    
    # Liquidity Chart
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=results_df['Fortnight'],
        y=results_df['Closing'],
        mode='lines+markers',
        name='Closing Cash',
        line=dict(color='#1565C0', width=3),
        marker=dict(size=10)
    ))
    fig.add_hline(y=0, line_dash="dash", line_color="red", annotation_text="Solvency Floor")
    fig.add_hline(y=200000, line_dash="dash", line_color="green", annotation_text="Target Buffer ($200k)")
    fig.update_layout(
        title="Liquidity Forecast",
        height=350,
        template='plotly_white',
        yaxis_tickformat='$,.0f'
    )
    st.plotly_chart(fig, width='stretch')


def render_profit_control():
    """Render PROFIT_CONTROL sub-tab - Income statement projection."""
    st.subheader("📊 PROFIT CONTROL - Income Statement Forecast")
    
    # Historical margins
    with st.expander("📋 Historical Margins (From Balance Statements)", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Gross Margin %", f"{st.session_state.cfo_gross_margin_pct:.1%}")
        with col2:
            st.metric("Net Margin %", f"{st.session_state.cfo_net_margin_pct:.1%}")
    
    st.markdown("### Income Statement Comparison")
    
    # Build income statement DataFrame
    last_revenue = st.session_state.cfo_net_sales
    last_cogs = st.session_state.cfo_cogs
    last_gross = st.session_state.cfo_gross_margin
    last_net = st.session_state.cfo_net_profit
    gross_margin_pct = st.session_state.cfo_gross_margin_pct
    
    income_data = pd.DataFrame({
        'Line_Item': ['Net Sales / Revenue', 'Cost of Goods Sold', 'Gross Margin', 
                     'S&A Expenses', 'Depreciation', 'Interest', 'Net Profit'],
        'Last_Round': [last_revenue, last_cogs, last_gross, 200000, 50000, 20000, last_net],
        'This_Round': [last_revenue * 1.1, 0, 0, 200000, 50000, 20000, 0]
    })
    
    # Calculate projected values
    proj_revenue = income_data.at[0, 'This_Round']
    proj_cogs = proj_revenue * (1 - gross_margin_pct)
    proj_gross = proj_revenue - proj_cogs
    proj_expenses = income_data.at[3, 'This_Round'] + income_data.at[4, 'This_Round'] + income_data.at[5, 'This_Round']
    proj_net = proj_gross - proj_expenses
    
    income_data.at[1, 'This_Round'] = proj_cogs
    income_data.at[2, 'This_Round'] = proj_gross
    income_data.at[6, 'This_Round'] = proj_net
    
    # Calculate variance
    income_data['Variance'] = income_data.apply(
        lambda row: (row['This_Round'] - row['Last_Round']) / row['Last_Round'] 
        if row['Last_Round'] != 0 else 0, axis=1
    )
    
    # Styling
    EDITABLE_STYLE = {'backgroundColor': '#E3F2FD', 'color': '#1565C0'}
    CALC_STYLE = {'backgroundColor': '#E8F5E9', 'color': '#2E7D32'}
    REF_STYLE = {'backgroundColor': '#F5F5F5', 'color': '#616161'}
    
    # JS Logic for Variance
    # Variance = (This_Round - Last_Round) / Last_Round
    variance_getter = JsCode("""
        function(params) {
            let last = Number(params.data.Last_Round);
            let this_round = Number(params.data.This_Round);
            if (last === 0) return 0;
            return (this_round - last) / last;
        }
    """)

    variance_js = JsCode("""
        function(params) {
            if (Math.abs(params.value) > 0.2) {
                return {'backgroundColor': '#FFCDD2', 'color': '#B71C1C', 'fontWeight': 'bold'};
            }
            return {'backgroundColor': '#E8F5E9', 'color': '#2E7D32'};
        }
    """)
    
    gb = GridOptionsBuilder.from_dataframe(income_data)
    gb.configure_column('Line_Item', headerName='Line Item', editable=False, width=180)
    gb.configure_column('Last_Round', headerName='Last Round', editable=False, width=150, 
                       valueFormatter="'$' + value.toLocaleString()", cellStyle=REF_STYLE)
    gb.configure_column('This_Round', headerName='This Round', editable=True, width=150,
                       valueFormatter="'$' + value.toLocaleString()", cellStyle=EDITABLE_STYLE)
    gb.configure_column('Variance', editable=False, width=120,
                       valueFormatter="(value * 100).toFixed(1) + '%'", 
                       cellStyle=variance_js,
                       valueGetter=variance_getter)
    
    AgGrid(
        income_data,
        gridOptions=gb.build(),
        fit_columns_on_grid_load=True,
        height=280,
        allow_unsafe_jscode=True,
        key='cfo_income_grid'
    )
    
    # Profit realism check
    historical_margin = st.session_state.cfo_net_margin_pct
    projected_margin = proj_net / proj_revenue if proj_revenue > 0 else 0
    
    if projected_margin > historical_margin + 0.05:
        st.warning("⚠️ **WARNING: Unrealistic profit jump!** Projected net margin exceeds historical by >5%")
    elif proj_net < 0:
        st.error("❌ **LOSS PROJECTED** - Review expenses or revenue forecast")
    else:
        st.success(f"✅ Projected Net Profit: ${proj_net:,.0f} ({projected_margin:.1%} margin)")


def render_balance_sheet_health():
    """Render BALANCE_SHEET_HEALTH sub-tab - Debt ratio tracking."""
    st.subheader("🏦 BALANCE SHEET HEALTH - Debt Ratio Analysis")
    
    st.markdown("""
    **Debt Ratio Thresholds:**
    - 🟢 **< 40%**: Healthy - Best available rates
    - 🟡 **40-60%**: Moderate - Standard rates  
    - 🔴 **> 60%**: Critical - Premium rates, may refuse credit
    """)
    
    col1, col2 = st.columns(2)
    with col1:
        assets = st.number_input(
            "Total Assets",
            value=int(st.session_state.cfo_total_assets),
            step=100000,
            key='cfo_assets_bs'
        )
        st.session_state.cfo_total_assets = assets
        
        retained = st.number_input(
            "Retained Earnings",
            value=int(st.session_state.cfo_retained_earnings),
            step=10000,
            key='cfo_retained'
        )
        st.session_state.cfo_retained_earnings = retained
    
    with col2:
        liabilities = st.number_input(
            "Total Liabilities",
            value=int(st.session_state.cfo_total_liabilities),
            step=100000,
            key='cfo_liab_bs'
        )
        st.session_state.cfo_total_liabilities = liabilities
        
        new_debt = st.number_input(
            "Planned New Debt",
            value=0,
            step=50000,
            key='cfo_new_debt'
        )
    
    # Calculate ratios
    current_ratio = liabilities / assets if assets > 0 else 0
    post_ratio = (liabilities + new_debt) / assets if assets > 0 else 0
    
    # Status flags
    if post_ratio > 0.6:
        st.error("🔴 CRITICAL: Debt too high - Credit rating at risk!")
    elif retained < 0:
        st.warning("⚠️ CRITICAL: Equity Erosion - Retained earnings negative")
    elif post_ratio > 0.4:
        st.warning("🟡 MODERATE: Approaching debt threshold")
    else:
        st.success("🟢 HEALTHY: Debt ratio within safe limits")
    
    # Gauge chart
    fig = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=post_ratio * 100,
        delta={'reference': current_ratio * 100, 'valueformat': '.1f'},
        gauge={
            'axis': {'range': [0, 100], 'tickformat': '.0f'},
            'bar': {'color': '#1565C0'},
            'steps': [
                {'range': [0, 40], 'color': '#C8E6C9'},
                {'range': [40, 60], 'color': '#FFF9C4'},
                {'range': [60, 100], 'color': '#FFCDD2'}
            ],
            'threshold': {
                'line': {'color': 'red', 'width': 4},
                'thickness': 0.75,
                'value': 60
            }
        },
        title={'text': 'Debt Ratio (%)'}
    ))
    fig.update_layout(height=300)
    st.plotly_chart(fig, width='stretch')
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Current Debt Ratio", f"{current_ratio:.1%}")
    with col2:
        st.metric("Post-Decision Ratio", f"{post_ratio:.1%}", delta=f"{(post_ratio - current_ratio):.1%}")


def render_debt_manager():
    """Render DEBT_MANAGER sub-tab - Mortgage calculator."""
    st.subheader("🏠 DEBT MANAGER - Mortgage Calculator")
    
    st.info(f"💡 Long-Term Debt Limit: ${LT_LIMIT:,.0f}")
    
    EDITABLE_STYLE = {'backgroundColor': '#E3F2FD', 'color': '#1565C0'}
    
    mortgage_df = st.session_state.cfo_mortgages.copy()
    
    gb = GridOptionsBuilder.from_dataframe(mortgage_df)
    gb.configure_column('Loan', editable=False, width=80)
    gb.configure_column('Amount', editable=True, width=120, 
                       type=['numericColumn'],
                       valueFormatter="'$' + value.toLocaleString()",
                       cellStyle=EDITABLE_STYLE)
    gb.configure_column('Interest_Rate', headerName='Interest Rate', editable=True, width=100,
                       valueFormatter="(value * 100).toFixed(1) + '%'")
    gb.configure_column('Payment_FN1', headerName='Payment FN1', editable=True, width=110,
                       type=['numericColumn'],
                       valueFormatter="'$' + value.toLocaleString()",
                       cellStyle=EDITABLE_STYLE)
    gb.configure_column('Payment_FN2', headerName='Payment FN2', editable=True, width=110,
                       type=['numericColumn'],
                       valueFormatter="'$' + value.toLocaleString()",
                       cellStyle=EDITABLE_STYLE)
    
    gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
    
    grid_response = AgGrid(
        mortgage_df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=150,
        key='cfo_mortgage_grid'
    )
    
    if grid_response.data is not None:
        st.session_state.cfo_mortgages = pd.DataFrame(grid_response.data)
    
    # Totals
    total_borrowed = mortgage_df['Amount'].sum()
    total_payments = mortgage_df['Payment_FN1'].sum() + mortgage_df['Payment_FN2'].sum()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Borrowed", f"${total_borrowed:,.0f}")
    with col2:
        st.metric("Total Payments", f"${total_payments:,.0f}")
    with col3:
        remaining = total_borrowed - total_payments
        st.metric("Outstanding", f"${remaining:,.0f}")
    
    if total_borrowed > LT_LIMIT:
        st.error(f"⚠️ Exceeds Long-Term Limit by ${total_borrowed - LT_LIMIT:,.0f}")


def render_upload_ready_finance():
    """Render UPLOAD_READY_FINANCE sub-tab - Export preview."""
    st.subheader("📤 UPLOAD READY - Finance Decisions")
    
    st.info("Copy these values to ExSim Finance Decision Form")
    
    # Credit Lines summary
    st.markdown("### 💳 Credit Lines")
    credit_summary = []
    for fn in FORTNIGHTS:
        val = st.session_state.cfo_financing.at[0, f'FN{fn}']
        if val != 0:
            credit_summary.append({'Fortnight': f'FN{fn}', 'Amount': val})
    
    if credit_summary:
        st.dataframe(pd.DataFrame(credit_summary), hide_index=True, width='stretch')
    else:
        st.caption("No credit line changes")
    
    # Investments summary
    st.markdown("### 📈 Investments")
    invest_summary = []
    for fn in FORTNIGHTS:
        val = st.session_state.cfo_financing.at[1, f'FN{fn}']
        if val != 0:
            invest_summary.append({'Fortnight': f'FN{fn}', 'Amount': val})
    
    if invest_summary:
        st.dataframe(pd.DataFrame(invest_summary), hide_index=True, width='stretch')
    else:
        st.caption("No investment changes")
    
    # Mortgages summary
    st.markdown("### 🏠 Mortgages")
    mortgages = st.session_state.cfo_mortgages[st.session_state.cfo_mortgages['Amount'] > 0]
    if not mortgages.empty:
        st.dataframe(mortgages, hide_index=True, width='stretch')
    else:
        st.caption("No mortgages")
    
    # Dividends summary
    st.markdown("### 💰 Dividends")
    div_total = sum(st.session_state.cfo_financing.at[3, f'FN{fn}'] for fn in FORTNIGHTS)
    st.metric("Total Dividends", f"${div_total:,.0f}")
    
    # CSV download button
    import io
    output = io.StringIO()
    output.write("=== CREDIT LINES ===\n")
    if credit_summary:
        pd.DataFrame(credit_summary).to_csv(output, index=False)
    output.write("\n=== INVESTMENTS ===\n")
    if invest_summary:
        pd.DataFrame(invest_summary).to_csv(output, index=False)
    output.write("\n=== MORTGAGES ===\n")
    mortgages.to_csv(output, index=False)
    output.write(f"\n=== DIVIDENDS ===\nTotal,{div_total}\n")
    csv_data = output.getvalue()
    
    st.download_button(
        label="📥 Download Decisions as CSV",
        data=csv_data,
        file_name="finance_decisions.csv",
        mime="text/csv",
        type="primary",
        key='cfo_csv_download'
    )


def render_cross_reference():
    """Render CROSS_REFERENCE sub-tab - Upstream data visibility."""
    st.subheader("🔗 CROSS REFERENCE - Upstream Cost Drivers")
    st.caption("Live visibility into Revenue Projections and Cost Centers.")
    
    # Load shared data
    try:
        from shared_outputs import import_dashboard_data
        cmo_data = import_dashboard_data('CMO') or {}
        cpo_data = import_dashboard_data('CPO') or {}
        purch_data = import_dashboard_data('Purchasing') or {}
        clo_data = import_dashboard_data('CLO') or {}
    except ImportError:
        st.error("Could not load shared_outputs module")
        cmo_data = {}
        cpo_data = {}
        purch_data = {}
        clo_data = {}
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📈 Revenue (CMO)")
        st.info("Sales Forecast based on current prices.")
        
        rev = cmo_data.get('est_revenue', 0)
        mkt_cost = cmo_data.get('marketing_cost', 0)
        
        st.metric("Proj. Revenue", f"${rev:,.0f}")
        st.metric("Marketing Budget", f"${mkt_cost:,.0f}")

    with col2:
        st.markdown("### 💸 Cost Structure (Variable)")
        st.info("Aggregated costs from all departments.")
        
        labor = cpo_data.get('total_labor_cost', 0)
        material = purch_data.get('supplier_spend', 0)
        logistics = clo_data.get('logistics_costs', 0)
        
        total_variable = labor + material + logistics
        
        # DataFrame for breakdown
        cost_df = pd.DataFrame([
            {"Category": "Labor (CPO)", "Amount": labor},
            {"Category": "Materials (Purchasing)", "Amount": material},
            {"Category": "Logistics (CLO)", "Amount": logistics},
            {"Category": "TOTAL VARIABLE", "Amount": total_variable}
        ])
        
        st.dataframe(cost_df.style.format({"Amount": "${:,.0f}"}), hide_index=True)


def render_cfo_tab():
    """Render the CFO (Finance) tab with 5 Excel-aligned subtabs."""
    init_cfo_state()
    sync_from_uploads()
    
    # Header with Download Button
    col_header, col_download = st.columns([4, 1])
    with col_header:
        st.header("💰 CFO Dashboard - Financial Control & Liquidity")
    with col_download:
        try:
            from utils.report_bridge import ReportBridge
            excel_buffer = ReportBridge.export_finance_dashboard()
            st.download_button(
                label="📥 Download Live",
                data=excel_buffer,
                file_name="Finance_Dashboard_Live.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"Export: {e}")
    
    # Data source status
    balance = get_state('balance_data')
    ar_ap = get_state('ar_ap_data')
    
    data_status = []
    if balance:
        data_status.append("✅ Balance Statements")
    if ar_ap:
        data_status.append("✅ AR/AP Schedule")
    
    if data_status:
        st.success(" | ".join(data_status))
    else:
        st.info("💡 Upload Balance Statements and AR/AP data in sidebar to populate financials")
    
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
    
    # 6 SUBTABS (Updated)
    subtabs = st.tabs([
        "💧 Liquidity Monitor",
        "📊 Profit Control",
        "🏦 Balance Sheet",
        "🏠 Debt Manager",
        "📤 Upload Ready",
        "🔗 Cross Reference"
    ])
    
    with subtabs[0]:
        render_liquidity_monitor()
    
    with subtabs[1]:
        render_profit_control()
    
    with subtabs[2]:
        render_balance_sheet_health()
    
    with subtabs[3]:
        render_debt_manager()
    
    with subtabs[4]:
        render_upload_ready_finance()
        
    with subtabs[5]:
        render_cross_reference()
