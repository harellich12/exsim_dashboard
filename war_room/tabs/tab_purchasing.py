"""
ExSim War Room - Purchasing Tab
5 sub-tabs mirroring the Excel dashboard sheets:
1. SUPPLIER_CONFIG - Supplier data configuration
2. COST_ANALYSIS - Ordering vs holding cost analysis
3. MRP_ENGINE - Material requirements planning
4. CASH_FLOW_PREVIEW - Procurement spending tracker
5. UPLOAD_READY_PROCUREMENT - Export preview
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode

from utils.state_manager import get_state, set_state

# Constants
FORTNIGHTS = list(range(1, 9))

# Default supplier configuration
SUPPLIERS = {
    'Part_A': {
        'A1': {'lead_time': 1, 'cost': 50, 'payment_terms': 2, 'batch_size': 500},
        'A2': {'lead_time': 2, 'cost': 45, 'payment_terms': 3, 'batch_size': 1000}
    },
    'Part_B': {
        'B1': {'lead_time': 1, 'cost': 30, 'payment_terms': 1, 'batch_size': 300},
        'B2': {'lead_time': 3, 'cost': 25, 'payment_terms': 2, 'batch_size': 600}
    }
}


def init_purchasing_state():
    """Initialize Purchasing state with MRP data structures."""
    if 'purchasing_initialized' not in st.session_state:
        st.session_state.purchasing_initialized = True
        
        raw_data = get_state('raw_materials_data')
        
        # Supplier configuration
        supplier_data = [
            {'Supplier': 'A1', 'Part': 'Part A', 'Lead_Time': 1, 'Cost': 50, 'Payment_Terms': 2, 'Batch_Size': 500},
            {'Supplier': 'A2', 'Part': 'Part A', 'Lead_Time': 2, 'Cost': 45, 'Payment_Terms': 3, 'Batch_Size': 1000},
            {'Supplier': 'B1', 'Part': 'Part B', 'Lead_Time': 1, 'Cost': 30, 'Payment_Terms': 1, 'Batch_Size': 300},
            {'Supplier': 'B2', 'Part': 'Part B', 'Lead_Time': 3, 'Cost': 25, 'Payment_Terms': 2, 'Batch_Size': 600}
        ]
        st.session_state.purchasing_suppliers = pd.DataFrame(supplier_data)
        
        # MRP Engine data
        opening_inv = raw_data.get('opening_inventory', 1000) if raw_data else 1000
        
        mrp_data = {
            'Item': ['Target_Production', 'Opening_Inventory', 'Gross_Requirement', 
                    'Scheduled_Arrivals', 'Projected_Inventory', 'Net_Deficit'],
            **{f'FN{fn}': [0, opening_inv if fn == 1 else 0, 0, 0, 0, 0] for fn in FORTNIGHTS}
        }
        st.session_state.purchasing_mrp = pd.DataFrame(mrp_data)
        
        # Orders by supplier
        orders_data = []
        for supplier in ['A1', 'A2', 'B1', 'B2']:
            orders_data.append({
                'Supplier': supplier,
                **{f'FN{fn}': 0 for fn in FORTNIGHTS}
            })
        st.session_state.purchasing_orders = pd.DataFrame(orders_data)
        
        # Cost analysis
        st.session_state.purchasing_ordering_cost = 5000
        st.session_state.purchasing_holding_cost = 2000


def calculate_mrp():
    """Calculate MRP projections."""
    mrp = st.session_state.purchasing_mrp.copy()
    orders = st.session_state.purchasing_orders
    suppliers = st.session_state.purchasing_suppliers
    
    for fn in FORTNIGHTS:
        fn_col = f'FN{fn}'
        
        # Get target production (row 0)
        target = mrp.at[0, fn_col]
        
        # Gross requirement = target
        mrp.at[2, fn_col] = target
        
        # Calculate scheduled arrivals from orders + lead times
        arrivals = 0
        for _, order_row in orders.iterrows():
            supplier = order_row['Supplier']
            sup_info = suppliers[suppliers['Supplier'] == supplier].iloc[0]
            lead_time = int(sup_info['Lead_Time'])
            
            # Orders placed in earlier FN arrive now
            order_fn = fn - lead_time
            if 1 <= order_fn <= 8:
                arrivals += order_row[f'FN{order_fn}']
        
        mrp.at[3, fn_col] = arrivals
        
        # Opening inventory (previous closing or initial)
        if fn == 1:
            opening = mrp.at[1, 'FN1']
        else:
            opening = mrp.at[4, f'FN{fn-1}']
        
        mrp.at[1, fn_col] = opening
        
        # Projected inventory
        projected = opening + arrivals - target
        mrp.at[4, fn_col] = projected
        
        # Net deficit
        deficit = max(0, -projected)
        mrp.at[5, fn_col] = deficit
    
    return mrp


def render_supplier_config():
    """Render SUPPLIER_CONFIG sub-tab."""
    st.subheader("🏪 SUPPLIER CONFIG - Supplier Data")
    
    st.markdown("""
    Configure your supplier data from the case study.  
    **Lead Time:** Fortnights until delivery | **Batch Size:** Minimum order quantity
    """)
    
    EDITABLE_STYLE = {'backgroundColor': '#E3F2FD', 'color': '#1565C0'}
    
    suppliers_df = st.session_state.purchasing_suppliers.copy()
    
    gb = GridOptionsBuilder.from_dataframe(suppliers_df)
    gb.configure_column('Supplier', editable=False, width=80)
    gb.configure_column('Part', editable=False, width=80)
    gb.configure_column('Lead_Time', editable=True, width=100, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Cost', editable=True, width=80, type=['numericColumn'], 
                       valueFormatter="'$' + value", cellStyle=EDITABLE_STYLE)
    gb.configure_column('Payment_Terms', editable=True, width=120, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Batch_Size', editable=True, width=100, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
    
    grid_response = AgGrid(
        suppliers_df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=200,
        key='supplier_config_grid'
    )
    
    if grid_response.data is not None:
        st.session_state.purchasing_suppliers = pd.DataFrame(grid_response.data)


def render_cost_analysis():
    """Render COST_ANALYSIS sub-tab."""
    st.subheader("📊 COST ANALYSIS - Ordering vs Holding")
    
    col1, col2 = st.columns(2)
    with col1:
        ordering = st.number_input(
            "Ordering Cost ($/order)",
            value=int(st.session_state.purchasing_ordering_cost),
            step=500,
            key='ordering_cost_input'
        )
        st.session_state.purchasing_ordering_cost = ordering
    
    with col2:
        holding = st.number_input(
            "Holding Cost ($/unit/period)",
            value=int(st.session_state.purchasing_holding_cost),
            step=200,
            key='holding_cost_input'
        )
        st.session_state.purchasing_holding_cost = holding
    
    # Calculate ratio
    total = ordering + holding
    ratio = ordering / total * 100 if total > 0 else 50
    
    # Interpretation
    if ratio > 70:
        interpretation = "🔴 **Ordering too frequently** - INCREASE batch sizes"
        color = '#FFCDD2'
    elif ratio < 30:
        interpretation = "🟡 **Holding too much inventory** - DECREASE batch sizes"
        color = '#FFF9C4'
    else:
        interpretation = "🟢 **Balanced approach** - Maintain current policy"
        color = '#C8E6C9'
    
    st.markdown(f"### Ordering Cost Ratio: {ratio:.1f}%")
    st.markdown(interpretation)
    
    # Gauge chart
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=ratio,
        gauge={
            'axis': {'range': [0, 100]},
            'bar': {'color': '#1565C0'},
            'steps': [
                {'range': [0, 30], 'color': '#FFF9C4'},
                {'range': [30, 70], 'color': '#C8E6C9'},
                {'range': [70, 100], 'color': '#FFCDD2'}
            ]
        },
        title={'text': 'Ordering Cost Ratio (%)'}
    ))
    fig.update_layout(height=300)
    st.plotly_chart(fig, use_container_width=True)


def render_mrp_engine():
    """Render MRP_ENGINE sub-tab."""
    st.subheader("📦 MRP ENGINE - Material Requirements Planning")
    
    st.info("""
    **Time Travel Rule:** Order in FN X → Arrives in FN X + Lead Time  
    RED = Negative inventory = ORDER NEEDED!
    """)
    
    EDITABLE_STYLE = {'backgroundColor': '#E3F2FD', 'color': '#1565C0'}
    REFERENCE_STYLE = {'backgroundColor': '#F5F5F5', 'color': '#616161'}
    
    # Section A: Target Production
    st.markdown("### Section A: Target Production")
    
    mrp_df = st.session_state.purchasing_mrp.copy()
    target_row = mrp_df[mrp_df['Item'] == 'Target_Production'].copy()
    
    gb = GridOptionsBuilder.from_dataframe(target_row)
    gb.configure_column('Item', editable=False, width=150)
    for fn in FORTNIGHTS:
        gb.configure_column(f'FN{fn}', editable=True, width=80, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
    
    grid_response = AgGrid(
        target_row,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=80,
        key='mrp_target_grid'
    )
    
    if grid_response.data is not None:
        updated = pd.DataFrame(grid_response.data)
        for fn in FORTNIGHTS:
            mrp_df.at[0, f'FN{fn}'] = updated.iloc[0][f'FN{fn}']
        st.session_state.purchasing_mrp = mrp_df
    
    # Section B: Net Requirements
    st.markdown("### Section B: Net Requirements")
    
    calculated_mrp = calculate_mrp()
    display_mrp = calculated_mrp[calculated_mrp['Item'].isin(['Projected_Inventory', 'Net_Deficit'])]
    
    inv_js = JsCode("""
        function(params) {
            if (params.value < 0) {
                return {'backgroundColor': '#FFCDD2', 'color': '#B71C1C', 'fontWeight': 'bold'};
            }
            return {'backgroundColor': '#C8E6C9', 'color': '#2E7D32'};
        }
    """)
    
    gb = GridOptionsBuilder.from_dataframe(display_mrp)
    gb.configure_column('Item', editable=False, width=150)
    for fn in FORTNIGHTS:
        gb.configure_column(f'FN{fn}', editable=False, width=80, cellStyle=inv_js)
    
    AgGrid(
        display_mrp,
        gridOptions=gb.build(),
        fit_columns_on_grid_load=True,
        height=100,
        allow_unsafe_jscode=True,
        key='mrp_display_grid'
    )
    
    # Section C: Sourcing Strategy
    st.markdown("### Section C: Sourcing Strategy (Orders by Supplier)")
    
    orders_df = st.session_state.purchasing_orders.copy()
    
    gb = GridOptionsBuilder.from_dataframe(orders_df)
    gb.configure_column('Supplier', editable=False, width=80)
    for fn in FORTNIGHTS:
        gb.configure_column(f'FN{fn}', editable=True, width=80, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
    
    grid_response = AgGrid(
        orders_df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=180,
        key='mrp_orders_grid'
    )
    
    if grid_response.data is not None:
        st.session_state.purchasing_orders = pd.DataFrame(grid_response.data)


def render_cash_flow_preview():
    """Render CASH_FLOW_PREVIEW sub-tab."""
    st.subheader("💵 CASH FLOW PREVIEW - Procurement Spending")
    
    orders = st.session_state.purchasing_orders
    suppliers = st.session_state.purchasing_suppliers
    
    # Calculate spend per fortnight
    spend_data = {'Fortnight': [f'FN{fn}' for fn in FORTNIGHTS]}
    
    total_spend = []
    for fn in FORTNIGHTS:
        fn_spend = 0
        for _, order_row in orders.iterrows():
            supplier = order_row['Supplier']
            sup_info = suppliers[suppliers['Supplier'] == supplier].iloc[0]
            qty = order_row[f'FN{fn}']
            cost = sup_info['Cost']
            fn_spend += qty * cost
        total_spend.append(fn_spend)
    
    spend_data['Spend'] = total_spend
    spend_data['Cumulative'] = [sum(total_spend[:i+1]) for i in range(len(total_spend))]
    
    spend_df = pd.DataFrame(spend_data)
    
    st.dataframe(spend_df, use_container_width=True, hide_index=True)
    
    # Chart
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=spend_df['Fortnight'],
        y=spend_df['Spend'],
        name='Per-FN Spend',
        marker_color='#1565C0'
    ))
    fig.add_trace(go.Scatter(
        x=spend_df['Fortnight'],
        y=spend_df['Cumulative'],
        name='Cumulative',
        mode='lines+markers',
        line=dict(color='#EF6C00', width=3)
    ))
    fig.update_layout(
        title='Procurement Cash Flow',
        height=350,
        template='plotly_white',
        yaxis_tickformat='$,.0f'
    )
    st.plotly_chart(fig, use_container_width=True)
    
    total = sum(total_spend)
    set_state('PROCUREMENT_COST', total)
    st.metric("**TOTAL PROCUREMENT SPEND**", f"${total:,.0f}")


def render_upload_ready_procurement():
    """Render UPLOAD_READY_PROCUREMENT sub-tab."""
    st.subheader("📤 UPLOAD READY - Procurement Decisions")
    
    st.info("Copy these values to ExSim Procurement Decision Form")
    
    # Orders summary
    st.markdown("### 📦 Order Summary")
    
    orders = st.session_state.purchasing_orders
    active_orders = orders[(orders[[f'FN{fn}' for fn in FORTNIGHTS]].sum(axis=1) > 0)]
    
    if not active_orders.empty:
        st.dataframe(active_orders, hide_index=True, use_container_width=True)
    else:
        st.caption("No orders placed")
    
    # Totals
    total_spend = get_state('PROCUREMENT_COST', 0)
    st.metric("Total Procurement", f"${total_spend:,.0f}")
    
    if st.button("📋 Copy Procurement Decisions", type="primary", key='purchasing_copy'):
        st.success("✅ Data copied! Paste into ExSim Procurement form.")


def render_purchasing_tab():
    """Render the Purchasing tab with 5 Excel-aligned subtabs."""
    init_purchasing_state()
    
    st.header("📦 Purchasing Dashboard - MRP & Sourcing")
    
    # Data status
    raw_data = get_state('raw_materials_data')
    if raw_data:
        st.success("✅ Raw Materials data loaded")
    else:
        st.info("💡 Upload Raw Materials file in sidebar for inventory data")
    
    # 5 SUBTABS
    subtabs = st.tabs([
        "🏪 SUPPLIER_CONFIG",
        "📊 COST_ANALYSIS",
        "📦 MRP_ENGINE",
        "💵 CASH_FLOW_PREVIEW",
        "📤 UPLOAD_READY"
    ])
    
    with subtabs[0]:
        render_supplier_config()
    
    with subtabs[1]:
        render_cost_analysis()
    
    with subtabs[2]:
        render_mrp_engine()
    
    with subtabs[3]:
        render_cash_flow_preview()
    
    with subtabs[4]:
        render_upload_ready_procurement()
