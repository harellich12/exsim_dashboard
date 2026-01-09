"""
ExSim War Room - CFO (Finance) Tab
6 sub-tabs mirroring the Excel dashboard sheets:
1. LIQUIDITY_MONITOR - Cash flow by fortnight
2. PROFIT_CONTROL - Income statement projection vs actuals
3. BALANCE_SHEET_HEALTH - Debt ratio tracking
4. DEBT_MANAGER - Mortgage calculator with Table VIII.1 data
5. UPLOAD_READY_FINANCE - Export preview
6. CROSS_REFERENCE - Upstream data visibility
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
    from case_parameters import COMMON, PRODUCTION, FINANCE, FINANCIAL_STATEMENTS
    FORTNIGHTS = COMMON.get('FORTNIGHTS', list(range(1, 9)))
except ImportError:
    FORTNIGHTS = list(range(1, 9))
    PRODUCTION = {}
    FINANCE = {}
    FINANCIAL_STATEMENTS = {}

# Get limits from case_parameters (Table VIII.1) or use fallback
if FINANCE:
    credit_line = FINANCE.get('LINE_OF_CREDIT', {})
    mortgages = FINANCE.get('MORTGAGES', {})
    ST_LIMIT = credit_line.get('calculated_limit', 230216)  # 33% of net assets
    LT_LIMIT = mortgages.get('limit', 800000)
    CREDIT_RATE = credit_line.get('interest_rate_per_period', 0.10)
    MORTGAGE_RATE = mortgages.get('interest_rate_per_period', 0.06)
    DEPOSIT_RATE = FINANCE.get('SHORT_TERM_DEPOSITS', {}).get('interest_rate_per_period', 0.04)
    EMERGENCY_RATE = FINANCE.get('EMERGENCY_LOAN', {}).get('interest_rate_per_period', 0.30)
else:
    ST_LIMIT = 230216  # From Table VIII.1
    LT_LIMIT = 800000
    CREDIT_RATE = 0.10
    MORTGAGE_RATE = 0.06
    DEPOSIT_RATE = 0.04
    EMERGENCY_RATE = 0.30

DEFAULT_TAX_RATE = 0.25


def calculate_projected_costs():
    """
    Calculate machine depreciation and module costs from case_parameters.
    Returns (depreciation, admin_cost, rent_cost) tuple.
    """
    machinery = PRODUCTION.get('MACHINES', PRODUCTION.get('MACHINERY', {}))
    initial_machines = PRODUCTION.get('INITIAL_MACHINES', {})
    facilities = PRODUCTION.get('FACILITIES', {})
    
    total_depr = 0.0
    
    # Machine Depreciation: price / lifespan for each machine
    for region, sections in initial_machines.items():
        for section, machines in sections.items():
            for m_type, count in machines.items():
                m_config = machinery.get(m_type)
                # Handle M3_ALPHA vs M3-ALPHA naming inconsistency
                if not m_config and "_" in m_type:
                    m_config = machinery.get(m_type.replace("_", "-"))
                if m_config:
                    price = m_config.get('price', 0)
                    lifespan = m_config.get('lifespan_periods', 10)
                    if lifespan > 0:
                        total_depr += (count * price / lifespan)
    
    # Module Costs from case_parameters
    initial_modules = facilities.get('INITIAL_MODULES', {})
    total_module_count = sum(initial_modules.values())
    admin_per_module = facilities.get('ADMIN_COST_PER_MODULE_PER_PERIOD', 10000)
    rent_per_module = facilities.get('MODULE_RENT_COST_PER_PERIOD', 7500)
    
    total_admin = total_module_count * admin_per_module
    total_rent = total_module_count * rent_per_module
    
    return total_depr, total_admin, total_rent



def init_cfo_state():
    """Initialize CFO state with proper financial data structures from case_parameters."""
    if 'cfo_initialized' not in st.session_state:
        st.session_state.cfo_initialized = True
        
        # Balance data from upload
        balance = get_state('balance_data')
        
        # Section A: Initialization (from Chapter IX Financial Statements)
        # Verify CASH matches Table IX.2 ($219,615)
        st.session_state.cfo_cash_at_end_last_period = FINANCIAL_STATEMENTS.get('BALANCE_SHEET', {}).get('ASSETS', {}).get('CURRENT', {}).get('cash', 219615)
        st.session_state.cfo_tax_payments = FINANCIAL_STATEMENTS.get('INCOME_STATEMENT', {}).get('TAXES', 52346)
        st.session_state.cfo_dividend_payments = 0
        st.session_state.cfo_asset_purchases = 0
        
        # Get initial AP/AR from case_parameters (Tables VIII.3 & VIII.4)
        initial_ap = FINANCE.get('INITIAL_AP', {})
        initial_ar = FINANCE.get('INITIAL_AR', {})
        
        # Section B: Operational Cash Flow (8 fortnights) with AP/AR pre-populated
        ar_values = {f'FN{fn}': initial_ar.get(fn, 0) for fn in FORTNIGHTS}
        ap_values = {f'FN{fn}': initial_ap.get(fn, 0) for fn in FORTNIGHTS}
        
        # Calculate Sales Receipts from CMO data (Table VIII.2 - Customer payments in FN 2,4,6,8)
        # First try to get CMO-based receivables, otherwise use Initial AR as starting point
        try:
            cmo_receivables = calculate_receivables_from_cmo()
            sales_receipts = {f'FN{fn}': cmo_receivables.get(fn, 0) for fn in FORTNIGHTS}
            # If CMO data gives zeros, use Initial AR for FN2 and estimate others
            if sum(sales_receipts.values()) == 0:
                # Use Period 6 Net Sales / 4 for each sales fortnight (FN2, FN4, FN6, FN8)
                p6_sales = FINANCIAL_STATEMENTS.get('INCOME_STATEMENT', {}).get('NET_SALES', 1183541)
                sales_per_fn = p6_sales / 4  # ~$296K per sales FN
                sales_receipts = {
                    'FN1': 0, 'FN2': sales_per_fn, 'FN3': 0, 'FN4': sales_per_fn,
                    'FN5': 0, 'FN6': sales_per_fn, 'FN7': 0, 'FN8': sales_per_fn
                }
        except:
            # Fallback: Use Period 6 Net Sales divided by 4 sales fortnights
            p6_sales = FINANCIAL_STATEMENTS.get('INCOME_STATEMENT', {}).get('NET_SALES', 1183541)
            sales_per_fn = p6_sales / 4
            sales_receipts = {
                'FN1': 0, 'FN2': sales_per_fn, 'FN3': 0, 'FN4': sales_per_fn,
                'FN5': 0, 'FN6': sales_per_fn, 'FN7': 0, 'FN8': sales_per_fn
            }
        
        st.session_state.cfo_cash_flow = pd.DataFrame({
            'Item': ['Sales Receipts', 'Procurement Spend', 'Fixed Overhead (S&A)', 
                    'Receivables (HARD)', 'Payables (HARD)'],
            **{f'FN{fn}': [sales_receipts[f'FN{fn}'], 0, 25000, ar_values[f'FN{fn}'], ap_values[f'FN{fn}']] for fn in FORTNIGHTS}
        })
        
        # Credit line starting balance from Table VIII.1
        credit_balance = FINANCE.get('LINE_OF_CREDIT', {}).get('current_balance', 113000)
        deposit_balance = FINANCE.get('SHORT_TERM_DEPOSITS', {}).get('current_balance', 200000)
        
        # Section C: Financing Decisions (8 fortnights)
        st.session_state.cfo_financing = pd.DataFrame({
            'Item': ['Credit Line (+/-)', 'Investments (+/-)', 
                    'Mortgage Inflow', 'Dividends Paid'],
            **{f'FN{fn}': [0, 0, 0, 0] for fn in FORTNIGHTS}
        })
        
        # Store credit/deposit balances
        st.session_state.cfo_credit_line_balance = credit_balance
        st.session_state.cfo_deposit_balance = deposit_balance
        
        # Balance data - Now using REAL data from Chapter IX instead of placeholders
        income = FINANCIAL_STATEMENTS.get('INCOME_STATEMENT', {})
        balance_sheet = FINANCIAL_STATEMENTS.get('BALANCE_SHEET', {})
        
        # If balance data uploaded, use it (priority), otherwise use Chapter IX defaults
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
            st.session_state.cfo_net_sales = income.get('NET_SALES', 1183541)
            st.session_state.cfo_cogs = income.get('COGS', 481439)
            st.session_state.cfo_gross_margin = income.get('GROSS_INCOME', 702101)
            st.session_state.cfo_net_profit = income.get('NET_PROFIT', 52346)
            
            total_assets = balance_sheet.get('ASSETS', {}).get('TOTAL_ASSETS', 1924943)
            st.session_state.cfo_total_assets = total_assets
            
            total_liabilities = balance_sheet.get('LIABILITIES_EQUITY', {}).get('LIABILITIES', {}).get('total_liabilities', 839314)
            st.session_state.cfo_total_liabilities = total_liabilities
            
            retained = balance_sheet.get('LIABILITIES_EQUITY', {}).get('EQUITY', {}).get('retained_earnings', 183281)
            st.session_state.cfo_retained_earnings = retained
            
            # Calculate margins
            sales = st.session_state.cfo_net_sales
            if sales > 0:
                st.session_state.cfo_gross_margin_pct = st.session_state.cfo_gross_margin / sales
                st.session_state.cfo_net_margin_pct = st.session_state.cfo_net_profit / sales
            else:
                st.session_state.cfo_gross_margin_pct = 0.59  # ~$702k/$1.18M
                st.session_state.cfo_net_margin_pct = 0.04    # ~$52k/$1.18M

        
        # Mortgages from Table VIII.1 - pre-populated with actual data
        mortgage_balance = FINANCE.get('MORTGAGES', {}).get('current_balance', 500000)
        st.session_state.cfo_mortgages = pd.DataFrame({
            'Loan': ['Current Mortgage', 'New Loan 1', 'New Loan 2'],
            'Amount': [mortgage_balance, 0, 0],
            'Interest_Rate': [MORTGAGE_RATE, MORTGAGE_RATE, MORTGAGE_RATE],
            'Payment_Period': [10, 0, 0],  # Period 10 for first payment
            'Payment_Amount': [240000, 0, 0]  # From payment schedule
        })



def calculate_receivables_from_cmo():
    """
    Calculate receivables schedule based on CMO payment terms (Table III.3).
    Sales occur in even fortnights (2, 4, 6, 8), collection is delayed by payment term.
    
    Returns: dict {fn: amount} for fortnights 1-8
    """
    try:
        from shared_outputs import import_dashboard_data
        from case_parameters import COMMON
        
        cmo_data = import_dashboard_data('CMO') or {}
        payment_terms_config = COMMON.get('PAYMENT_TERMS', {
            'A': {'fortnights': 0, 'discount': 0.130},
            'B': {'fortnights': 2, 'discount': 0.075},
            'C': {'fortnights': 4, 'discount': 0.025},
            'D': {'fortnights': 8, 'discount': 0.000}
        })
        
        receivables = {fn: 0 for fn in range(1, 9)}
        
        demand_forecast = cmo_data.get('demand_forecast', {})
        pricing = cmo_data.get('pricing', {})
        payment_terms = cmo_data.get('payment_terms', {})
        
        for zone in demand_forecast.keys():
            # Convert to float for type safety (JSON may serialize as strings)
            demand = float(demand_forecast.get(zone, 0)) if demand_forecast.get(zone) else 0
            price = float(pricing.get(zone, 100)) if pricing.get(zone) else 100
            term_code = payment_terms.get(zone, 'D')  # Default to D (no discount, 8 fortnight delay)
            
            term_config = payment_terms_config.get(term_code, payment_terms_config['D'])
            delay = term_config['fortnights']
            discount = term_config['discount']
            
            # Revenue per period after discount
            revenue_per_zone = demand * price * (1 - discount)
            # Sales occur in even fortnights (2, 4, 6, 8) - split revenue evenly
            sales_per_fortnight = revenue_per_zone / 4
            
            for sale_fn in [2, 4, 6, 8]:
                collect_fn = sale_fn + delay
                if 1 <= collect_fn <= 8:
                    receivables[collect_fn] += sales_per_fortnight
        
        return receivables
        
    except Exception as e:
        print(f"Error calculating receivables from CMO: {e}")
        return {fn: 0 for fn in range(1, 9)}

def sync_from_uploads():
    """Sync CFO data from uploaded files."""
    balance = get_state('balance_data')
    ar_ap_data = get_state('ar_ap_data')
    
    # Guard: ensure cfo state exists
    if 'cfo_cash_flow' not in st.session_state:
        return
    
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
    
    # Get costs from module-level function
    calc_depr, calc_admin, calc_rent = calculate_projected_costs()
    
    # Use calculated or defaults
    current_depr = calc_depr if calc_depr > 0 else 50000
    current_admin = calc_admin if calc_admin > 0 else 60000
    current_rent = calc_rent if calc_rent > 0 else 45000

    income_data = pd.DataFrame({
        'Line_Item': ['Net Sales / Revenue', 'Cost of Goods Sold', 'Gross Margin', 
                     'Module Admin (S&A)', 'Module Rent', 'Depreciation', 'Interest', 'Net Profit'],
        'Last_Round': [last_revenue, last_cogs, last_gross, current_admin, current_rent, current_depr, 20000, last_net],
        'This_Round': [last_revenue * 1.1, 0, 0, current_admin, current_rent, current_depr, 20000, 0]
    })
    
    # Calculate projected values
    proj_revenue = income_data.at[0, 'This_Round']
    proj_cogs = proj_revenue * (1 - gross_margin_pct)
    proj_gross = proj_revenue - proj_cogs
    # Sum indices 3, 4, 5, 6 (Admin, Rent, Depr, Interest)
    proj_expenses = income_data.at[3, 'This_Round'] + income_data.at[4, 'This_Round'] + \
                    income_data.at[5, 'This_Round'] + income_data.at[6, 'This_Round']
    proj_net = proj_gross - proj_expenses
    
    income_data.at[1, 'This_Round'] = proj_cogs
    income_data.at[2, 'This_Round'] = proj_gross
    income_data.at[7, 'This_Round'] = proj_net
    
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
    """Render DEBT_MANAGER sub-tab - Financing options from Table VIII.1."""
    st.subheader("🏦 DEBT MANAGER - Financing Options (Table VIII.1)")
    
    # Financing Options Reference Table
    st.markdown("### 📊 Available Financing Options")
    
    financing_df = pd.DataFrame([
        {'Option': 'Line of Credit', 'Interest': f'{CREDIT_RATE*100:.0f}%', 
         'Limit': f'33% net assets (${ST_LIMIT:,.0f})', 'Current': f'${FINANCE.get("LINE_OF_CREDIT", {}).get("current_balance", 113000):,.0f}',
         'Timing': 'Per fortnight'},
        {'Option': 'Short-term Deposits', 'Interest': f'{DEPOSIT_RATE*100:.0f}%', 
         'Limit': 'No limit', 'Current': f'${FINANCE.get("SHORT_TERM_DEPOSITS", {}).get("current_balance", 200000):,.0f}',
         'Timing': 'Per fortnight'},
        {'Option': 'Mortgages', 'Interest': f'{MORTGAGE_RATE*100:.0f}%', 
         'Limit': f'${LT_LIMIT:,.0f}', 'Current': f'${FINANCE.get("MORTGAGES", {}).get("current_balance", 500000):,.0f}',
         'Timing': 'Per period'},
        {'Option': '⚠️ Emergency Loan', 'Interest': f'{EMERGENCY_RATE*100:.0f}%', 
         'Limit': 'Auto if negative cash', 'Current': '$0',
         'Timing': 'AVOID!'}
    ])
    st.dataframe(financing_df, use_container_width=True, hide_index=True)
    
    # Emergency Loan Warning
    st.error("⚠️ **Emergency Loan Warning**: 30% interest rate! Deliberate use = de facto bankruptcy. Maintain positive cash at all times!")
    
    # Credit Line Calculator
    st.markdown("### 💳 Credit Line Calculator")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        net_assets = st.number_input(
            "Net Fixed Assets ($)", 
            value=FINANCE.get('LINE_OF_CREDIT', {}).get('net_fixed_assets_p6', 697625),
            step=10000,
            help="From Period 6 balance sheet"
        )
    with col2:
        credit_limit = net_assets * 0.33
        st.metric("Credit Limit (33%)", f"${credit_limit:,.0f}")
    with col3:
        current_used = FINANCE.get('LINE_OF_CREDIT', {}).get('current_balance', 113000)
        available = credit_limit - current_used
        st.metric("Available Credit", f"${available:,.0f}")
    
    if available < 50000:
        st.warning("⚠️ Low available credit! Consider reducing credit line usage or increasing assets.")
    
    # Mortgage Manager
    st.markdown("### 🏠 Mortgage Manager")
    st.info(f"💡 Long-Term Debt Limit: ${LT_LIMIT:,.0f} | Interest Rate: {MORTGAGE_RATE*100:.0f}% per period")
    
    EDITABLE_STYLE = {'backgroundColor': '#E3F2FD', 'color': '#1565C0'}
    
    mortgage_df = st.session_state.cfo_mortgages.copy()
    
    gb = GridOptionsBuilder.from_dataframe(mortgage_df)
    gb.configure_column('Loan', editable=False, width=120)
    gb.configure_column('Amount', editable=True, width=120, 
                       type=['numericColumn'],
                       valueFormatter="'$' + value.toLocaleString()",
                       cellStyle=EDITABLE_STYLE)
    gb.configure_column('Interest_Rate', headerName='Interest Rate', editable=False, width=100,
                       valueFormatter="(value * 100).toFixed(1) + '%'")
    gb.configure_column('Payment_Period', headerName='Payment Period', editable=True, width=110,
                       cellStyle=EDITABLE_STYLE)
    gb.configure_column('Payment_Amount', headerName='Payment Amount', editable=True, width=120,
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
    
    # Mortgage Payment Schedule from Table VIII.1
    st.markdown("### 📅 Payment Schedule (from Table VIII.1)")
    schedule = FINANCE.get('MORTGAGES', {}).get('payment_schedule', [])
    if schedule:
        schedule_df = pd.DataFrame(schedule)
        schedule_df.columns = ['Period', 'Amount']
        schedule_df['Amount'] = schedule_df['Amount'].apply(lambda x: f"${x:,.0f}")
        st.dataframe(schedule_df, use_container_width=True, hide_index=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Borrowed", f"${total_borrowed:,.0f}")
    with col2:
        interest_per_period = total_borrowed * MORTGAGE_RATE
        st.metric("Interest (per period)", f"${interest_per_period:,.0f}")
    with col3:
        st.metric("Remaining Limit", f"${max(0, LT_LIMIT - total_borrowed):,.0f}")
    
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
        
        # Ensure safely converted floats
        try:
            labor = float(labor)
        except (ValueError, TypeError):
            labor = 0.0
            
        try:
            material = float(material)
        except (ValueError, TypeError):
            material = 0.0
            
        try:
            logistics = float(logistics)
        except (ValueError, TypeError):
            logistics = 0.0
        
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
    
    # ---------------------------------------------------------
    # EXSIM SHARED OUTPUTS - EXPORT
    # ---------------------------------------------------------
    try:
        from shared_outputs import export_dashboard_data
        
        # Calculate final outputs for export
        # 'cash_flow_projection', 'debt_levels', 'liquidity_status'
        
        # Liquidity Status (Projected FN8 or FN1? Or min?)
        # Let's take the status of the first period or the worst status?
        # Typically the 'Next Period' is most relevant (FN1)
        results_df = calculate_cash_flow()
        fn1_status = results_df.at[0, 'Status']
        final_cash = results_df.at[0, 'Closing']
        
        # Debt Levels
        # Total Borrowed / Total Assets? Or just total amount?
        # Schema: "debt_levels": "835497" (string?) or number. 
        # Using string in example, allow flexible.
        total_debt = st.session_state.cfo_total_liabilities
        
        outputs = {
            'cash_flow_projection': {'final_cash': final_cash, 'tax_payments': st.session_state.cfo_tax_payments},
            'debt_levels': total_debt,
            'liquidity_status': fn1_status
        }
        
        export_dashboard_data('CFO', outputs)
        
    except Exception as e:
        print(f"Shared Output Export Error: {e}")

