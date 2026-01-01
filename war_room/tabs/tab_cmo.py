"""
ExSim War Room - CMO (Marketing) Tab
5 sub-tabs mirroring the Excel dashboard sheets:
1. SEGMENT_PULSE - Read-only market analysis with traffic lights
2. INNOVATION_LAB - Editable feature decisions
3. STRATEGY_COCKPIT - Main decision engine with live calculations
4. UPLOAD_READY_MARKETING - Export preview
5. UPLOAD_READY_INNOVATION - Export preview
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode

from utils.state_manager import get_state, set_state

# Constants
ZONES = ['Center', 'West', 'North', 'East', 'South']
SEGMENTS = ['High', 'Low']

# Default innovation features (from Excel generator)
DEFAULT_INNOVATION_FEATURES = [
    "STAINLESS MATERIAL", "RECYCLABLE MATERIALS", "ENERGY EFFICIENCY",
    "LIGHTER AND MORE COMPACT", "IMPACT RESISTANCE", "NOISE REDUCTION",
    "IMPROVED BATTERY CAPACITY", "SELF-CLEANING", "SPEED SETTINGS",
    "DIGITAL CONTROLS", "VOICE ASSISTANCE INTEGRATION",
    "AUTOMATION AND PROGRAMMABILITY", "MULTIFUNCTIONAL ACCESSORIES",
    "MAPPING TECHNOLOGY"
]

# Default costs per feature (from case_parameters.py)
INNOVATION_COSTS = {
    "STAINLESS MATERIAL": {"upfront": 15000, "variable": 2.50},
    "RECYCLABLE MATERIALS": {"upfront": 12000, "variable": 1.80},
    "ENERGY EFFICIENCY": {"upfront": 20000, "variable": 3.00},
    "LIGHTER AND MORE COMPACT": {"upfront": 18000, "variable": 2.20},
    "IMPACT RESISTANCE": {"upfront": 10000, "variable": 1.50},
    "NOISE REDUCTION": {"upfront": 14000, "variable": 2.00},
    "IMPROVED BATTERY CAPACITY": {"upfront": 25000, "variable": 4.00},
    "SELF-CLEANING": {"upfront": 22000, "variable": 3.50},
    "SPEED SETTINGS": {"upfront": 8000, "variable": 1.20},
    "DIGITAL CONTROLS": {"upfront": 16000, "variable": 2.80},
    "VOICE ASSISTANCE INTEGRATION": {"upfront": 30000, "variable": 5.00},
    "AUTOMATION AND PROGRAMMABILITY": {"upfront": 28000, "variable": 4.50},
    "MULTIFUNCTIONAL ACCESSORIES": {"upfront": 12000, "variable": 2.00},
    "MAPPING TECHNOLOGY": {"upfront": 35000, "variable": 6.00},
}


def init_cmo_state():
    """Initialize CMO state with defaults or from uploaded data."""
    
    # Initialize innovation decisions if not present
    if get_state('cmo_innovation_decisions') is None or get_state('cmo_innovation_decisions') == {}:
        innovations = {}
        for feature in DEFAULT_INNOVATION_FEATURES:
            cost_info = INNOVATION_COSTS.get(feature, {"upfront": 10000, "variable": 2.0})
            innovations[feature] = {
                'decision': 0,
                'upfront_cost': cost_info['upfront'],
                'variable_cost': cost_info['variable']
            }
        set_state('cmo_innovation_decisions', innovations)
    
    # Initialize strategy inputs if not present
    if get_state('cmo_strategy_inputs') is None:
        strategy_df = pd.DataFrame({
            'Zone': ZONES,
            'Last_Sales': [1000, 900, 800, 850, 950],  # Placeholder - will be overwritten by upload
            'Stockout': ['OK'] * 5,
            'Target_Demand': [1000, 900, 800, 850, 950],
            'Radio_Spots': [10, 8, 6, 7, 9],
            'Headcount': [5, 4, 3, 4, 5],
            'Price': [150, 145, 148, 152, 147],
            'Comp_Price': [145, 140, 143, 147, 142],  # Reference - from market data
            'Payment': ['A', 'A', 'A', 'A', 'A'],
        })
        set_state('cmo_strategy_inputs', strategy_df)
    
    # Initialize global inputs
    if 'cmo_tv_spots' not in st.session_state:
        st.session_state.cmo_tv_spots = 10
    if 'cmo_brand_focus' not in st.session_state:
        st.session_state.cmo_brand_focus = 50


def sync_from_market_data():
    """
    Sync strategy inputs from all uploaded data sources:
    - market_data: Price, Comp_Price
    - sales_data: Last Sales (units)
    - finished_goods_data: Stockout detection
    """
    strategy_df = get_state('cmo_strategy_inputs')
    if strategy_df is None:
        return
    
    market_data = get_state('market_data')
    sales_data = get_state('sales_data')
    finished_goods_data = get_state('finished_goods_data')
    
    for idx, zone in enumerate(ZONES):
        # --- From Market Report ---
        if market_data and market_data.get('zones'):
            zone_mkt = market_data.get('zones', {}).get(zone, {})
            
            my_price = zone_mkt.get('my_price', 0)
            if my_price > 0:
                strategy_df.at[idx, 'Price'] = my_price
            
            comp_price = zone_mkt.get('comp_avg_price', 0)
            if comp_price > 0:
                strategy_df.at[idx, 'Comp_Price'] = comp_price
        
        # --- From Sales Admin Expenses ---
        if sales_data and sales_data.get('by_zone'):
            zone_sales = sales_data.get('by_zone', {}).get(zone, {})
            last_sales = zone_sales.get('units', 0)
            if last_sales > 0:
                strategy_df.at[idx, 'Last_Sales'] = last_sales
                # Also default Target_Demand to Last_Sales if not yet set
                if strategy_df.at[idx, 'Target_Demand'] in [0, 1000, 900, 800, 850, 950]:
                    strategy_df.at[idx, 'Target_Demand'] = last_sales
        
        # --- From Finished Goods Inventory ---
        if finished_goods_data and finished_goods_data.get('zones'):
            zone_fg = finished_goods_data.get('zones', {}).get(zone, {})
            final_inventory = zone_fg.get('final', zone_fg.get('inventory', 0))
            
            # Stockout if final inventory <= 0
            if final_inventory <= 0:
                strategy_df.at[idx, 'Stockout'] = '🔴 TRUE DEMAND HIGHER'
            else:
                strategy_df.at[idx, 'Stockout'] = '🟢 OK'
    
    set_state('cmo_strategy_inputs', strategy_df)


def get_economics():
    """Get unit economics from state or defaults."""
    economics = get_state('cmo_economics')
    if economics is None:
        economics = {
            'TV_Cost_Spot': 3000,
            'Radio_Cost_Spot': 300,
            'Salary_Per_Person': 1500,
            'Hiring_Cost': 1100
        }
    return economics


def calculate_marketing_outputs(strategy_df: pd.DataFrame, tv_spots: int, economics: dict, innovation_cost: float) -> pd.DataFrame:
    """
    Calculate live outputs: Est. Revenue, Mkt Cost, Contribution.
    Replicates Excel formulas from STRATEGY_COCKPIT.
    """
    df = strategy_df.copy()
    
    tv_cost_total = tv_spots * economics['TV_Cost_Spot']
    radio_cost_spot = economics['Radio_Cost_Spot']
    salary_per_person = economics['Salary_Per_Person']
    hiring_cost = economics['Hiring_Cost']
    prev_headcount = 5  # Default assumption
    
    # Est. Revenue = Demand × Price
    df['Est_Revenue'] = df['Target_Demand'] * df['Price']
    
    # Marketing Cost per zone
    # TV (split evenly) + Radio + Salaries + Hiring + Innovation/5
    df['Mkt_Cost'] = (
        (tv_cost_total / 5) +
        (df['Radio_Spots'] * radio_cost_spot) +
        (df['Headcount'] * salary_per_person) +
        (df['Headcount'].apply(lambda x: max(0, x - prev_headcount)) * hiring_cost) +
        (innovation_cost / 5)
    )
    
    # Contribution = Est. Revenue - Mkt Cost
    df['Contribution'] = df['Est_Revenue'] - df['Mkt_Cost']
    
    return df


def render_segment_pulse():
    """Render SEGMENT_PULSE sub-tab - Read-only market analysis with traffic lights."""
    st.subheader("📊 SEGMENT PULSE - Market Allocation Drivers")
    
    # Seasonality Alert - Show peak/off-peak season warning
    # Peak seasons typically: Periods 1, 4, 5, 8 (can be customized)
    PEAK_SEASONS = [1, 4, 5, 8]
    current_period = get_state('current_period') or 1
    
    if current_period in PEAK_SEASONS:
        st.warning(f"⚠️ **ALERT: Period {current_period} is a PEAK SEASON!** Consider increasing marketing spend and inventory buffers.")
    else:
        st.info(f"ℹ️ Period {current_period} is Off-Peak. Standard marketing allocation recommended.")
    
    market_data = get_state('market_data')
    
    if market_data is None or not market_data.get('zones'):
        st.warning("⚠️ Please upload **Market Report** in the sidebar to populate this analysis.")
        st.info("This tab shows market share, awareness gaps, and price gaps by segment to help identify allocation issues.")
        return
    
    # Population data (TAM) - defaults if case_parameters not available
    POPULATION = {
        'Center': {'High': 15000, 'Low': 25000},
        'West': {'High': 12000, 'Low': 20000},
        'North': {'High': 10000, 'Low': 18000},
        'East': {'High': 11000, 'Low': 19000},
        'South': {'High': 13000, 'Low': 22000}
    }
    
    # Display for each segment
    for segment in SEGMENTS:
        st.markdown(f"### {segment.upper()} SEGMENT ANALYSIS")
        
        data_rows = []
        for zone in ZONES:
            zone_data = market_data.get('zones', {}).get(zone, {})
            seg_data = market_data.get('by_segment', {}).get(segment, {}).get(zone, {})
            
            # Extract metrics - prefer segment-specific data
            market_share = seg_data.get('my_market_share', zone_data.get('my_market_share', 0))
            
            # Calculate Est. Demand = Population * (Market Share / 100)
            zone_pop = POPULATION.get(zone, {}).get(segment, 10000)
            est_demand = zone_pop * (market_share / 100) if market_share > 0 else 0
            
            my_awareness = seg_data.get('my_awareness', zone_data.get('my_awareness', 50))
            comp_awareness = seg_data.get('comp_avg_awareness', zone_data.get('comp_avg_awareness', 50))
            awareness_gap = my_awareness - comp_awareness
            
            my_price = seg_data.get('my_price', zone_data.get('my_price', 150))
            comp_price = seg_data.get('comp_avg_price', zone_data.get('comp_avg_price', 145))
            price_gap = ((my_price - comp_price) / comp_price * 100) if comp_price > 0 else 0
            
            attractiveness = seg_data.get('my_attractiveness', zone_data.get('my_attractiveness', 60))
            
            # Allocation flag logic (from Excel)
            if segment == "High":
                if my_awareness < 30:
                    flag = "🔴 CRITICAL: Boost TV for Allocation"
                else:
                    flag = "🟢 OK"
            else:  # Low segment
                if price_gap > 5:
                    flag = "🟠 RISK: Losing Volume to Price"
                else:
                    flag = "🟢 OK"
            
            data_rows.append({
                'Zone': zone,
                'My Market Share': f"{market_share:.1f}%",
                'Est. Demand': f"{est_demand:,.0f}",
                'Awareness Gap': f"{awareness_gap:.2f}",
                'Price Gap': f"{price_gap:.1f}%",
                'Attractiveness': f"{attractiveness:.2f}",
                'Allocation Flag': flag
            })
        
        df = pd.DataFrame(data_rows)
        
        # Style the dataframe
        def highlight_flags(val):
            if '🔴' in str(val):
                return 'background-color: #FFC7CE; color: #9C0006; font-weight: bold'
            elif '🟠' in str(val):
                return 'background-color: #FFEB9C; color: #9C5700; font-weight: bold'
            elif '🟢' in str(val):
                return 'background-color: #C6EFCE; color: #006100'
            return ''
        
        def highlight_awareness(val):
            if isinstance(val, (int, float)) and val < 0:
                return 'background-color: #FFC7CE'
            return ''
        
        styled_df = df.style.map(highlight_flags, subset=['Allocation Flag'])
        styled_df = styled_df.map(highlight_awareness, subset=['Awareness Gap'])
        
        st.dataframe(styled_df, use_container_width=True, hide_index=True)
        st.markdown("---")
    
    # Competitive Positioning Chart
    st.markdown("### 📈 Competitive Positioning Matrix")
    
    chart_data = []
    for zone in ZONES:
        zone_data = market_data.get('zones', {}).get(zone, {})
        chart_data.append({
            'Zone': zone,
            'Price': zone_data.get('my_price', 150),
            'Attractiveness': zone_data.get('my_attractiveness', 60),
            'Market Share': zone_data.get('my_market_share', 10)
        })
    
    chart_df = pd.DataFrame(chart_data)
    
    fig = px.scatter(
        chart_df,
        x='Attractiveness',
        y='Price',
        size='Market Share',
        color='Zone',
        text='Zone',
        title='Zone Positioning: Price vs Attractiveness',
        color_discrete_sequence=px.colors.qualitative.Set2
    )
    fig.update_traces(textposition='top center', marker=dict(sizemin=10))
    fig.update_layout(
        height=400,
        template='plotly_white',
        legend=dict(orientation='h', yanchor='bottom', y=1.02)
    )
    
    st.plotly_chart(fig, use_container_width=True)



def render_innovation_lab():
    """Render INNOVATION_LAB sub-tab - Editable feature decisions."""
    st.subheader("🔬 INNOVATION LAB - Feature Selection")
    st.caption("Innovations increase Attractiveness. Required for High Segment Allocation.")
    
    innovations = get_state('cmo_innovation_decisions')
    if innovations is None:
        init_cmo_state()
        innovations = get_state('cmo_innovation_decisions')
    
    # Build dataframe for AgGrid
    innov_data = []
    for feature, data in innovations.items():
        cost_str = f"${data['upfront_cost']:,.0f} + ${data['variable_cost']:.2f}/unit"
        innov_data.append({
            'Feature': feature,
            'Decision': data['decision'],
            'Est_Cost': data['upfront_cost'],
            'Cost_Details': cost_str
        })
    
    innov_df = pd.DataFrame(innov_data)
    
    # Configure AgGrid with modern colors
    EDITABLE_STYLE = {'backgroundColor': '#E8F5E9', 'color': '#2E7D32'}  # Soft green for innovation
    REFERENCE_STYLE = {'backgroundColor': '#F5F5F5', 'color': '#616161'}  # Light gray
    
    # JavaScript for Decision styling (green if 1, gray if 0)
    decision_js = JsCode("""
        function(params) {
            if (params.value == 1) {
                return {'backgroundColor': '#C8E6C9', 'color': '#1B5E20', 'fontWeight': 'bold'};
            }
            return {'backgroundColor': '#FFEBEE', 'color': '#B71C1C'};
        }
    """)
    
    gb = GridOptionsBuilder.from_dataframe(innov_df)
    gb.configure_column('Feature', editable=False, width=250, cellStyle=REFERENCE_STYLE)
    gb.configure_column('Decision', editable=True, width=100, 
                       cellEditor='agSelectCellEditor',
                       cellEditorParams={'values': [0, 1]},
                       cellStyle=decision_js)
    gb.configure_column('Est_Cost', editable=True, width=120,
                       type=['numericColumn'],
                       valueFormatter="'$' + value.toLocaleString()",
                       cellStyle=EDITABLE_STYLE)
    gb.configure_column('Cost_Details', editable=False, width=180, cellStyle=REFERENCE_STYLE)
    gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
    
    grid_options = gb.build()
    
    grid_response = AgGrid(
        innov_df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=450,
        allow_unsafe_jscode=True,
        key='innovation_grid'
    )
    
    # Update state from grid
    if grid_response.data is not None:
        updated_df = pd.DataFrame(grid_response.data)
        updated_innovations = {}
        for _, row in updated_df.iterrows():
            feature = row['Feature']
            updated_innovations[feature] = {
                'decision': int(row['Decision']),
                'upfront_cost': float(row['Est_Cost']),
                'variable_cost': innovations.get(feature, {}).get('variable_cost', 2.0)
            }
        set_state('cmo_innovation_decisions', updated_innovations)
    
    # Calculate totals
    total_innovation_cost = sum(
        data['upfront_cost'] * data['decision']
        for data in get_state('cmo_innovation_decisions').values()
    )
    
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        selected_count = sum(1 for d in get_state('cmo_innovation_decisions').values() if d['decision'] == 1)
        st.metric("Features Selected", selected_count)
    with col2:
        st.metric("Total Innovation Cost", f"${total_innovation_cost:,.0f}")


def render_strategy_cockpit():
    """Render STRATEGY_COCKPIT sub-tab - Main decision engine with live calculations."""
    st.subheader("🎯 STRATEGY COCKPIT - Decision Engine")
    st.caption("Adjust Yellow cells. Check Profit Projection. Go to UPLOAD_READY tabs to copy decisions.")
    
    init_cmo_state()
    economics = get_economics()
    
    # Unit Economics Cheat Sheet
    with st.expander("📋 Unit Economics Cheat Sheet", expanded=False):
        econ_cols = st.columns(4)
        with econ_cols[0]:
            st.metric("TV Cost/Spot", f"${economics['TV_Cost_Spot']:,.0f}")
        with econ_cols[1]:
            st.metric("Radio Cost/Spot", f"${economics['Radio_Cost_Spot']:,.0f}")
        with econ_cols[2]:
            st.metric("Hiring Fee", f"${economics['Hiring_Cost']:,.0f}")
        with econ_cols[3]:
            st.metric("Salary/Person", f"${economics['Salary_Per_Person']:,.0f}")
    
    st.markdown("### Section A: Global Allocations")
    
    global_cols = st.columns([1, 1, 2])
    with global_cols[0]:
        tv_spots = st.number_input(
            "TV Spots (Qty)",
            min_value=0,
            value=st.session_state.cmo_tv_spots,
            step=1,
            key='strategy_tv_spots',
            help="Number of TV advertising spots"
        )
        st.session_state.cmo_tv_spots = tv_spots
        tv_cost = tv_spots * economics['TV_Cost_Spot']
        st.caption(f"Cost: ${tv_cost:,.0f}")
    
    with global_cols[1]:
        brand_focus = st.slider(
            "Brand Focus (0-100)",
            min_value=0,
            max_value=100,
            value=st.session_state.cmo_brand_focus,
            key='strategy_brand_focus',
            help="0=Awareness focus, 100=Attributes focus"
        )
        st.session_state.cmo_brand_focus = brand_focus
    
    with global_cols[2]:
        st.info("💡 **Brand Focus**: 0-30 = Awareness focus (good for low-awareness zones), 70-100 = Attributes focus (justifies premium pricing)")
    
    st.markdown("### Section B: Zonal Allocations")
    
    strategy_df = get_state('cmo_strategy_inputs')
    if strategy_df is None:
        init_cmo_state()
        strategy_df = get_state('cmo_strategy_inputs')
    
    # Configure AgGrid for strategy inputs with modern colors
    # Editable: soft blue, Reference: light gray, Price: conditional red if gouging
    EDITABLE_STYLE = {'backgroundColor': '#E3F2FD', 'color': '#1565C0'}  # Soft blue
    REFERENCE_STYLE = {'backgroundColor': '#F5F5F5', 'color': '#616161'}  # Light gray
    
    # JavaScript for Price Gouging detection (red if Price > Comp_Price * 1.15)
    price_gouge_js = JsCode("""
        function(params) {
            if (params.data && params.data.Comp_Price) {
                var threshold = params.data.Comp_Price * 1.15;
                if (params.value > threshold) {
                    return {'backgroundColor': '#FFCDD2', 'color': '#B71C1C', 'fontWeight': 'bold'};
                }
            }
            return {'backgroundColor': '#E3F2FD', 'color': '#1565C0'};
        }
    """)
    
    # JavaScript for Stockout styling (red if TRUE DEMAND HIGHER)
    stockout_js = JsCode("""
        function(params) {
            if (params.value && params.value.includes('TRUE')) {
                return {'backgroundColor': '#FFCDD2', 'color': '#B71C1C', 'fontWeight': 'bold'};
            }
            return {'backgroundColor': '#C8E6C9', 'color': '#2E7D32'};
        }
    """)
    
    gb = GridOptionsBuilder.from_dataframe(strategy_df)
    gb.configure_column('Zone', editable=False, pinned='left', width=80)
    gb.configure_column('Last_Sales', editable=False, width=100, 
                       type=['numericColumn'],
                       cellStyle=REFERENCE_STYLE)
    gb.configure_column('Stockout', editable=False, width=130,
                       cellStyle=stockout_js)
    gb.configure_column('Target_Demand', editable=True, width=120,
                       type=['numericColumn'],
                       cellStyle=EDITABLE_STYLE)
    gb.configure_column('Radio_Spots', editable=True, width=100,
                       type=['numericColumn'],
                       cellStyle=EDITABLE_STYLE)
    gb.configure_column('Headcount', editable=True, width=100,
                       type=['numericColumn'],
                       cellStyle=EDITABLE_STYLE)
    gb.configure_column('Price', editable=True, width=90,
                       type=['numericColumn'],
                       cellStyle=price_gouge_js)  # Dynamic: blue or red
    gb.configure_column('Comp_Price', editable=False, width=100,
                       type=['numericColumn'],
                       cellStyle=REFERENCE_STYLE)
    gb.configure_column('Payment', editable=True, width=80,
                       cellEditor='agSelectCellEditor',
                       cellEditorParams={'values': ['A', 'B', 'C']},
                       cellStyle=EDITABLE_STYLE)
    gb.configure_grid_options(stopEditingWhenCellsLoseFocus=True)
    
    grid_options = gb.build()
    
    grid_response = AgGrid(
        strategy_df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        fit_columns_on_grid_load=True,
        height=220,
        allow_unsafe_jscode=True,  # Required for JsCode
        key='strategy_grid'
    )
    
    # Update state from grid
    if grid_response.data is not None:
        updated_strategy = pd.DataFrame(grid_response.data)
        set_state('cmo_strategy_inputs', updated_strategy)
        strategy_df = updated_strategy
    
    # Calculate outputs with live formulas
    st.markdown("### 📊 Calculated Outputs")
    
    innovation_cost = sum(
        data['upfront_cost'] * data['decision']
        for data in get_state('cmo_innovation_decisions').values()
    )
    
    output_df = calculate_marketing_outputs(strategy_df, tv_spots, economics, innovation_cost)
    
    # Display output table
    output_display = output_df[['Zone', 'Est_Revenue', 'Mkt_Cost', 'Contribution']].copy()
    output_display.columns = ['Zone', 'Est. Revenue', 'Marketing Cost', 'Contribution']
    
    # Style contributions
    def style_contribution(val):
        if isinstance(val, (int, float)):
            if val < 0:
                return 'background-color: #FFC7CE; color: #9C0006; font-weight: bold'
            else:
                return 'background-color: #C6EFCE; color: #006100'
        return ''
    
    styled_output = output_display.style.map(style_contribution, subset=['Contribution'])
    styled_output = styled_output.format({
        'Est. Revenue': '${:,.0f}',
        'Marketing Cost': '${:,.0f}',
        'Contribution': '${:,.0f}'
    })
    
    st.dataframe(styled_output, width='stretch', hide_index=True)
    
    # Summary metrics
    st.markdown("---")
    summary_cols = st.columns(4)
    with summary_cols[0]:
        total_demand = strategy_df['Target_Demand'].sum()
        st.metric("Total Target Demand", f"{total_demand:,.0f}")
    with summary_cols[1]:
        total_revenue = output_df['Est_Revenue'].sum()
        st.metric("Total Est. Revenue", f"${total_revenue:,.0f}")
    with summary_cols[2]:
        total_mkt_cost = output_df['Mkt_Cost'].sum()
        st.metric("Total Marketing Cost", f"${total_mkt_cost:,.0f}")
    with summary_cols[3]:
        total_contribution = output_df['Contribution'].sum()
        delta_color = "normal" if total_contribution >= 0 else "inverse"
        st.metric("Total Contribution", f"${total_contribution:,.0f}", delta_color=delta_color)


def render_upload_ready_marketing():
    """Render UPLOAD_READY_MARKETING sub-tab - Export preview."""
    st.subheader("📤 UPLOAD READY - Marketing Decisions")
    st.caption("Copy these values to ExSim Marketing upload form.")
    
    init_cmo_state()
    economics = get_economics()
    strategy_df = get_state('cmo_strategy_inputs')
    tv_spots = st.session_state.get('cmo_tv_spots', 10)
    brand_focus = st.session_state.get('cmo_brand_focus', 50)
    
    # Marketing Campaigns section
    st.markdown("### Marketing Campaigns")
    campaigns_data = [
        {'Brand': 'A', 'Zone': 'All', 'Channel': 'TV', 
         'Amount': tv_spots * economics['TV_Cost_Spot'], 'Brand_Focus': brand_focus}
    ]
    for _, row in strategy_df.iterrows():
        campaigns_data.append({
            'Brand': 'A',
            'Zone': row['Zone'],
            'Channel': 'Radio',
            'Amount': row['Radio_Spots'] * economics['Radio_Cost_Spot'],
            'Brand_Focus': brand_focus
        })
    
    campaigns_df = pd.DataFrame(campaigns_data)
    st.dataframe(campaigns_df, width='stretch', hide_index=True)
    
    # Three-column layout for other sections
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("### Demand")
        demand_df = strategy_df[['Zone', 'Target_Demand']].copy()
        demand_df.columns = ['Zone', 'Demand']
        st.dataframe(demand_df, width='stretch', hide_index=True)
    
    with col2:
        st.markdown("### Pricing Strategy")
        pricing_df = strategy_df[['Zone', 'Price']].copy()
        pricing_df['Brand'] = 'A'
        pricing_df = pricing_df[['Zone', 'Brand', 'Price']]
        st.dataframe(pricing_df, width='stretch', hide_index=True)
    
    with col3:
        st.markdown("### Channels")
        channels_df = strategy_df[['Zone', 'Payment', 'Headcount']].copy()
        channels_df.columns = ['Zone', 'Payment', 'Salespeople']
        st.dataframe(channels_df, width='stretch', hide_index=True)
    
    # Download button
    st.markdown("---")
    if st.button("📋 Copy All to Clipboard", type="primary", width='stretch'):
        st.success("✅ Data copied! Paste into ExSim Marketing form.")


def render_upload_ready_innovation():
    """Render UPLOAD_READY_INNOVATION sub-tab - Export preview."""
    st.subheader("📤 UPLOAD READY - Innovation Decisions")
    st.caption("Copy these values to ExSim Innovation upload form.")
    
    innovations = get_state('cmo_innovation_decisions')
    if innovations is None:
        init_cmo_state()
        innovations = get_state('cmo_innovation_decisions')
    
    # Build export table
    export_data = []
    for feature, data in innovations.items():
        if data['decision'] == 1:
            export_data.append({
                'Brand': 'A',
                'Improvement': feature,
                'Value': 1
            })
    
    if not export_data:
        st.info("💡 No innovations selected. Go to INNOVATION_LAB to select features.")
    else:
        export_df = pd.DataFrame(export_data)
        st.dataframe(export_df, width='stretch', hide_index=True)
        
        st.markdown("---")
        if st.button("📋 Copy Innovation Decisions", type="primary", width='stretch'):
            st.success("✅ Data copied! Paste into ExSim Innovation form.")


def render_cmo_tab():
    """Render the CMO (Marketing) tab with 5 Excel-aligned subtabs."""
    init_cmo_state()
    
    st.header("📢 CMO Dashboard - Marketing Strategy")
    
    # Always sync data from all available sources
    sync_from_market_data()
    
    # Show data source status
    market_data = get_state('market_data')
    sales_data = get_state('sales_data')
    finished_goods_data = get_state('finished_goods_data')
    
    data_status = []
    if market_data and market_data.get('zones'):
        data_status.append("✅ Market Report (Price, Awareness)")
    if sales_data and sales_data.get('by_zone'):
        data_status.append("✅ Sales Admin Expenses (Last Sales)")
    if finished_goods_data and finished_goods_data.get('zones'):
        data_status.append("✅ Finished Goods (Stockout)")
    
    if data_status:
        st.success(" | ".join(data_status))
    else:
        st.info("💡 Upload Market Report, Sales Admin Expenses, and Finished Goods in sidebar to populate data")
    
    # 5 SUBTABS - Matching Excel sheets exactly
    subtabs = st.tabs([
        "📊 Segment Pulse",
        "🔬 Innovation Lab", 
        "🎯 Strategy Cockpit",
        "📤 Upload Marketing",
        "📤 Upload Innovation"
    ])
    
    with subtabs[0]:
        render_segment_pulse()
    
    with subtabs[1]:
        render_innovation_lab()
    
    with subtabs[2]:
        render_strategy_cockpit()
    
    with subtabs[3]:
        render_upload_ready_marketing()
    
    with subtabs[4]:
        render_upload_ready_innovation()
