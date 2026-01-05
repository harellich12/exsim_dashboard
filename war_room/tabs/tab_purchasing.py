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

# Import centralized constants from case_parameters
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
try:
    from case_parameters import COMMON, PURCHASING, PRODUCTION
    FORTNIGHTS = COMMON.get('FORTNIGHTS', list(range(1, 9)))
    
    # Parts and supplier data from case
    PARTS_DATA = PURCHASING.get('PARTS', {})
    PIECES_DATA = PURCHASING.get('PIECES', {})
    INITIAL_INVENTORY = PRODUCTION.get('INITIAL_INVENTORY', {})
except ImportError:
    FORTNIGHTS = list(range(1, 9))
    PARTS_DATA = {}
    PIECES_DATA = {}
    INITIAL_INVENTORY = {}

# Build supplier configuration from case_parameters
def build_suppliers_from_case():
    """Build supplier table from case_parameters."""
    suppliers = []
    for part_name, part_info in PARTS_DATA.items():
        batch_size = part_info.get('batch_size', 30)
        ordering_cost = part_info.get('ordering_cost', 0)
        for sup_code, sup_info in part_info.get('suppliers', {}).items():
            price = sup_info.get('price', 0)
            discount = sup_info.get('discount', 0)
            discount_thresh = sup_info.get('discount_threshold', 0)
            payment = sup_info.get('payment_fortnights', 0)
            reliability = sup_info.get('delivery_rate', 1.0)
            suppliers.append({
                'Supplier': f'Supplier {sup_code}',
                'Part': part_name,
                'Lead_Time': 0,  # Immediate delivery (all suppliers deliver same fortnight)
                'Cost': price,   # Cost per batch
                'Price_Per_Batch': price,
                'Batch_Size': batch_size,
                'Payment_Terms': payment,  # Keep numeric for MRP calcs
                'Discount': f'{discount*100:.0f}%' if discount else 'None',
                'Discount_Threshold': str(discount_thresh) if discount else '-',
                'Reliability': f'{reliability*100:.0f}%'
            })
    return suppliers

SUPPLIERS = build_suppliers_from_case() if PARTS_DATA else [
    {'Supplier': 'Supplier A', 'Part': 'Part A', 'Lead_Time': 0, 'Cost': 125, 
     'Price_Per_Batch': 125, 'Batch_Size': 30, 'Payment_Terms': 0,
     'Discount': 'None', 'Discount_Threshold': '-', 'Reliability': '100%'},
    {'Supplier': 'Supplier B', 'Part': 'Part A', 'Lead_Time': 0, 'Cost': 100, 
     'Price_Per_Batch': 100, 'Batch_Size': 30, 'Payment_Terms': 2,
     'Discount': '16%', 'Discount_Threshold': 150, 'Reliability': '80%'},
]



def init_purchasing_state():
    """Initialize Purchasing state with MRP data structures."""
    if 'purchasing_initialized' not in st.session_state:
        st.session_state.purchasing_initialized = True
        
        raw_data = get_state('materials_data')
        
        # Supplier configuration from case_parameters
        st.session_state.purchasing_suppliers = pd.DataFrame(SUPPLIERS)
        
        # MRP Engine data - Use initial inventory from case if available
        opening_inv = 1000  # Default
        
        # Try to get initial inventory from case_parameters first
        if INITIAL_INVENTORY:
            center_inv = INITIAL_INVENTORY.get('Center', {})
            west_inv = INITIAL_INVENTORY.get('West', {})
            part_a = center_inv.get('Part A', 0) + west_inv.get('Part A', 0)
            part_b = center_inv.get('Part B', 0) + west_inv.get('Part B', 0)
            opening_inv = part_a + part_b if (part_a + part_b) > 0 else 1000
        
        # Override with uploaded data if available
        if raw_data and 'parts' in raw_data:
            part_a = raw_data['parts'].get('Part A (Unit)', {}).get('stock', 0)
            part_b = raw_data['parts'].get('Part B (Unit)', {}).get('stock', 0)
            if part_a == 0 and part_b == 0:
                total_stock = sum(p.get('stock', 0) for p in raw_data['parts'].values())
                opening_inv = total_stock if total_stock > 0 else opening_inv
            else:
                opening_inv = part_a + part_b
        
        mrp_data = {
            'Item': ['Target_Production', 'Opening_Inventory', 'Gross_Requirement', 
                    'Scheduled_Arrivals', 'Projected_Inventory', 'Net_Deficit'],
            **{f'FN{fn}': [0, opening_inv if fn == 1 else 0, 0, 0, 0, 0] for fn in FORTNIGHTS}
        }
        st.session_state.purchasing_mrp = pd.DataFrame(mrp_data)
        
        # Orders by supplier - dynamically build from SUPPLIERS list
        supplier_names = list(set(s['Supplier'] for s in SUPPLIERS)) if SUPPLIERS else ['Supplier A', 'Supplier B', 'Supplier C']
        orders_data = []
        for supplier in supplier_names:
            orders_data.append({
                'Supplier': supplier,
                **{f'FN{fn}': 0 for fn in FORTNIGHTS}
            })
        st.session_state.purchasing_orders = pd.DataFrame(orders_data)
        
        # Cost analysis - use actual values from case if available
        part_a_cost = PARTS_DATA.get('Part A', {}).get('ordering_cost', 2300)
        st.session_state.purchasing_ordering_cost = part_a_cost
        st.session_state.purchasing_holding_cost = 2000



def sync_from_production():
    """Sync 'Target_Production' from shared outputs."""
    try:
        from shared_outputs import import_dashboard_data
        prod_data = import_dashboard_data('Production')
        if prod_data and 'production_plan' in prod_data:
            prod_plan = prod_data['production_plan']
            # Sum target across all zones for each FN?
            # Production Plan export format: {'Center': {'Target': 1000}}
            # It seems we only have a single 'Target' value per zone (total?) or maybe per FN if we expanded it.
            # If the export is just a total target, we might need to distribute it.
            # But wait, looking at tab_production.py export:
            # prod_plan[zone] = {'Target': total_target}
            # total_target is sum of FN1..4. Use average or fill FN1?
            
            # Better approach: If we want per-FN granularity, we should have exported it.
            # For now, let's take the total target sum and spread it or just put it in FN1/FN2?
            # Or assume uniform distribution.
            
            total_target_all_zones = sum(z.get('Target', 0) for z in prod_plan.values())
            
            if total_target_all_zones > 0:
                # Distribute uniformly?
                mrp = st.session_state.purchasing_mrp
                # Avoid overwriting if user manually edited? 
                # Actually, MRP 'Target_Production' should be driven by Production Plan.
                # Let's overwrite.
                per_fn = total_target_all_zones / 4 # Assume 4 active FNs
                for i in range(1, 5):
                     mrp.at[0, f'FN{i}'] = per_fn
                st.session_state.purchasing_mrp = mrp
                
    except ImportError:
        pass


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
    gb.configure_column('Lead_Time', headerName='Lead Time', editable=True, width=100, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Cost', editable=True, width=80, type=['numericColumn'], 
                       valueFormatter="'$' + value", cellStyle=EDITABLE_STYLE)
    gb.configure_column('Payment_Terms', headerName='Payment Terms', editable=True, width=120, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
    gb.configure_column('Batch_Size', headerName='Batch Size', editable=True, width=100, type=['numericColumn'], cellStyle=EDITABLE_STYLE)
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
    st.plotly_chart(fig, width='stretch')


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
    
    st.dataframe(spend_df, width='stretch', hide_index=True)
    
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
    st.plotly_chart(fig, width='stretch')
    
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
        st.dataframe(active_orders, hide_index=True, width='stretch')
    else:
        st.caption("No orders placed")
    
    # Totals
    total_spend = get_state('PROCUREMENT_COST', 0)
    st.metric("Total Procurement", f"${total_spend:,.0f}")
    
    # CSV download button
    orders_export = st.session_state.purchasing_orders.copy()
    csv_data = orders_export.to_csv(index=False)
    st.download_button(
        label="📥 Download Decisions as CSV",
        data=csv_data,
        file_name="procurement_decisions.csv",
        mime="text/csv",
        type="primary",
        key='purchasing_csv_download'
    )


def render_cross_reference():
    """Render CROSS_REFERENCE sub-tab - Upstream data visibility."""
    st.subheader("🔗 CROSS REFERENCE - Upstream Support")
    st.caption("Live visibility into Production requirements and Finance limits.")
    
    # Load shared data
    try:
        from shared_outputs import import_dashboard_data
        prod_data = import_dashboard_data('Production') or {}
        cfo_data = import_dashboard_data('CFO') or {}
    except ImportError:
        st.error("Could not load shared_outputs module")
        prod_data = {}
        cfo_data = {}
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 🏭 Production (Material Needs)")
        st.info("Ensures you buy enough raw materials.")
        
        # Extract Production Plan Target Sum
        try:
            prod_plan = prod_data.get('production_plan', {})
            total_target = sum([d.get('Target', 0) for d in prod_plan.values()]) if isinstance(prod_plan, dict) else 0
            utilization = prod_data.get('capacity_utilization', {}).get('mean', 0)
        except:
            total_target = 0
            utilization = 0
            
        st.metric("Total Production Target", f"{total_target:,.0f} units")
        st.metric("Avg Capacity Utilization", f"{utilization*100:.1f}%")
        
        if total_target > 0:
            st.success("✅ Production Plan is Active")
        else:
            st.warning("⚠️ No Production Plan detected")

    with col2:
        st.markdown("### 💰 Finance (Liquidity)")
        st.info("Check if cash is available for bulk buying.")
        
        liquidity = cfo_data.get('liquidity_status', 'Unknown')
        cash_proj = cfo_data.get('cash_flow_projection', "N/A")
        
        st.metric("Liquidity Status", liquidity)
        
        if "CRITICAL" in liquidity:
            st.error(f"🔴 {liquidity} - Reduce Order Size!")
        elif "Stable" in liquidity:
            st.success(f"🟢 {liquidity} - OK to Order")
        else:
            st.info(f"ℹ️ {liquidity}")


def render_purchasing_tab():
    """Render the Purchasing tab with 5 Excel-aligned subtabs."""
    init_purchasing_state()
    sync_from_production() # AUTO-SYNC
    
    # Header with Download Button
    col_header, col_download = st.columns([4, 1])
    with col_header:
        st.header("📦 Purchasing Dashboard - MRP & Sourcing")
    with col_download:
        try:
            from utils.report_bridge import create_download_button
            create_download_button('Purchasing', 'Purchasing')
        except Exception as e:
            st.error(f"Export: {e}")
    
    # Data status
    raw_data = get_state('materials_data')
    if raw_data:
        st.success("✅ Raw Materials data loaded")
    else:
        st.info("💡 Upload Raw Materials file in sidebar for inventory data")
    
    # 6 SUBTABS (Updated to include Cross Reference)
    subtabs = st.tabs([
        "🏪 Supplier Config",
        "📊 Cost Analysis",
        "📦 MRP Engine",
        "💵 Cash Flow",
        "📤 Upload Ready",
        "🔗 Cross Reference"
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
        
    with subtabs[5]:
        render_cross_reference()
    
    # ---------------------------------------------------------
    # EXSIM SHARED OUTPUTS - EXPORT
    # ---------------------------------------------------------
    try:
        from shared_outputs import export_dashboard_data
        
        # Calculate final outputs for export
        # 'supplier_spend', 'inventory_levels'
        # Supplier Spend: Total procurement cost
        total_spend = get_state('PROCUREMENT_COST', 0)
        if total_spend == 0:
            # Re-calc if not set (page refresh)
            pass 
        
        # Inventory Levels: projected closing for FN1
        mrp = st.session_state.purchasing_mrp
        proj_inv = mrp.at[4, 'FN1']
        
        outputs = {
            'supplier_spend': total_spend,
            'inventory_levels': {'projected_closing': proj_inv}
        }
        
        export_dashboard_data('Purchasing', outputs)
        
    except Exception as e:
        print(f"Shared Output Export Error: {e}")
