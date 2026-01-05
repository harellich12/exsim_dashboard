"""
ExSim War Room - ESG Tab
4 sub-tabs mirroring the Excel dashboard sheets:
1. IMPACT_CONFIG - CO2 tax rates and initiative settings
2. STRATEGY_SELECTOR - Compare green investment options
3. RESULTS - Summary and recommendations
4. UPLOAD_READY_ESG - Export preview
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode

from utils.state_manager import get_state, set_state

# ESG Initiative defaults
ESG_INITIATIVES = {
    'Solar_PV': {'cost_per_unit': 15000, 'co2_reduction': 0.5, 'type': 'CAPEX', 'unit': 'panel'},
    'Trees': {'cost_per_unit': 50, 'co2_reduction': 0.02, 'type': 'CAPEX', 'unit': 'tree'},
    'Green_Electricity': {'cost_per_kwh': 0.03, 'co2_reduction': 0.5, 'type': 'OpEx', 'unit': 'kWh'},
    'CO2_Credits': {'cost_per_ton': 25, 'co2_reduction': 1.0, 'type': 'OpEx', 'unit': 'credit'}
}


def init_esg_state():
    """Initialize ESG state with green investment data."""
    if 'esg_initialized' not in st.session_state:
        st.session_state.esg_initialized = True
        
        # CO2 configuration
        st.session_state.esg_co2_tax_rate = 30  # $/ton
        st.session_state.esg_current_emissions = 100  # tons/year
        st.session_state.esg_energy_consumption = 500000  # kWh/year
        
        # Initiative quantities
        st.session_state.esg_solar_panels = 0
        st.session_state.esg_trees = 0
        st.session_state.esg_green_electricity_pct = 0
        st.session_state.esg_co2_credits = 0


def calculate_esg_impact():
    """Calculate ESG initiative costs and impacts."""
    tax_rate = st.session_state.esg_co2_tax_rate
    emissions = st.session_state.esg_current_emissions
    energy = st.session_state.esg_energy_consumption
    
    # Solar PV
    solar_qty = st.session_state.esg_solar_panels
    solar_cost = solar_qty * ESG_INITIATIVES['Solar_PV']['cost_per_unit']
    solar_reduction = solar_qty * ESG_INITIATIVES['Solar_PV']['co2_reduction']
    solar_annual_savings = solar_reduction * tax_rate
    solar_payback = solar_cost / solar_annual_savings if solar_annual_savings > 0 else 999
    solar_cost_per_ton = solar_cost / solar_reduction if solar_reduction > 0 else 0
    
    # Trees
    trees_qty = st.session_state.esg_trees
    trees_cost = trees_qty * ESG_INITIATIVES['Trees']['cost_per_unit']
    trees_reduction = trees_qty * ESG_INITIATIVES['Trees']['co2_reduction']
    trees_annual_savings = trees_reduction * tax_rate
    trees_payback = trees_cost / trees_annual_savings if trees_annual_savings > 0 else 999
    trees_cost_per_ton = trees_cost / trees_reduction if trees_reduction > 0 else 0
    
    # Green Electricity
    green_pct = st.session_state.esg_green_electricity_pct / 100
    green_kwh = energy * green_pct
    green_cost = green_kwh * ESG_INITIATIVES['Green_Electricity']['cost_per_kwh']
    green_reduction = green_kwh * ESG_INITIATIVES['Green_Electricity']['co2_reduction'] / 1000  # Convert to tons
    green_cost_per_ton = green_cost / green_reduction if green_reduction > 0 else 0
    
    # CO2 Credits
    credits_qty = st.session_state.esg_co2_credits
    credits_cost = credits_qty * ESG_INITIATIVES['CO2_Credits']['cost_per_ton']
    credits_reduction = credits_qty * ESG_INITIATIVES['CO2_Credits']['co2_reduction']
    credits_cost_per_ton = ESG_INITIATIVES['CO2_Credits']['cost_per_ton']
    
    # Totals
    total_reduction = solar_reduction + trees_reduction + green_reduction + credits_reduction
    remaining_emissions = max(0, emissions - total_reduction)
    tax_liability = remaining_emissions * tax_rate
    
    total_capex = solar_cost + trees_cost
    total_opex = green_cost + credits_cost
    
    return {
        'solar': {'qty': solar_qty, 'cost': solar_cost, 'reduction': solar_reduction, 
                 'payback': solar_payback, 'cost_per_ton': solar_cost_per_ton},
        'trees': {'qty': trees_qty, 'cost': trees_cost, 'reduction': trees_reduction,
                 'payback': trees_payback, 'cost_per_ton': trees_cost_per_ton},
        'green_elec': {'pct': green_pct * 100, 'cost': green_cost, 'reduction': green_reduction,
                       'cost_per_ton': green_cost_per_ton},
        'credits': {'qty': credits_qty, 'cost': credits_cost, 'reduction': credits_reduction,
                   'cost_per_ton': credits_cost_per_ton},
        'total_reduction': total_reduction,
        'remaining_emissions': remaining_emissions,
        'tax_liability': tax_liability,
        'total_capex': total_capex,
        'total_opex': total_opex
    }


def render_impact_config():
    """Render IMPACT_CONFIG sub-tab."""
    st.subheader("⚙️ IMPACT CONFIG - CO2 Tax & Initiative Settings")
    
    st.markdown("### CO2 Configuration")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        tax = st.number_input(
            "CO2 Tax Rate ($/ton)",
            value=int(st.session_state.esg_co2_tax_rate),
            step=5,
            key='esg_tax_input'
        )
        st.session_state.esg_co2_tax_rate = tax
    
    with col2:
        emissions = st.number_input(
            "Current Emissions (tons/year)",
            value=int(st.session_state.esg_current_emissions),
            step=10,
            key='esg_emissions_input'
        )
        st.session_state.esg_current_emissions = emissions
    
    with col3:
        energy = st.number_input(
            "Energy Consumption (kWh/year)",
            value=int(st.session_state.esg_energy_consumption),
            step=50000,
            key='esg_energy_input'
        )
        st.session_state.esg_energy_consumption = energy
    
    # Current tax liability
    current_tax = emissions * tax
    st.metric("Current Annual CO2 Tax", f"${current_tax:,.0f}")
    
    # Initiative settings table
    st.markdown("### Initiative Settings (Reference)")
    
    settings_df = pd.DataFrame([
        {'Initiative': 'Solar PV', 'Type': 'CAPEX', 'Cost': '$15,000/panel', 'CO2 Reduction': '0.5 tons/panel/year'},
        {'Initiative': 'Trees', 'Type': 'CAPEX', 'Cost': '$50/tree', 'CO2 Reduction': '0.02 tons/tree/year'},
        {'Initiative': 'Green Electricity', 'Type': 'OpEx', 'Cost': '$0.03/kWh premium', 'CO2 Reduction': '0.5 tons/1000 kWh'},
        {'Initiative': 'CO2 Credits', 'Type': 'OpEx', 'Cost': '$25/credit', 'CO2 Reduction': '1 ton/credit'}
    ])
    
    st.dataframe(settings_df, width='stretch', hide_index=True)


def render_strategy_selector():
    """Render STRATEGY_SELECTOR sub-tab."""
    st.subheader("🌱 STRATEGY SELECTOR - Green Investment Options")
    
    st.markdown("""
    **Decision Framework:**
    - IF Payback < 3 years → BUY SOLAR
    - IF Short-term cash is low → BUY CREDITS
    - Trees for PR & long-term
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### CAPEX Options")
        
        solar = st.number_input(
            "Solar PV Panels",
            value=int(st.session_state.esg_solar_panels),
            step=1,
            key='esg_solar_input'
        )
        st.session_state.esg_solar_panels = solar
        
        trees = st.number_input(
            "Trees to Plant",
            value=int(st.session_state.esg_trees),
            step=10,
            key='esg_trees_input'
        )
        st.session_state.esg_trees = trees
    
    with col2:
        st.markdown("### OpEx Options")
        
        green_pct = st.slider(
            "Green Electricity (%)",
            0, 100,
            value=int(st.session_state.esg_green_electricity_pct),
            key='esg_green_slider'
        )
        st.session_state.esg_green_electricity_pct = green_pct
        
        credits = st.number_input(
            "CO2 Credits to Buy",
            value=int(st.session_state.esg_co2_credits),
            step=5,
            key='esg_credits_input'
        )
        st.session_state.esg_co2_credits = credits
    
    # Calculate impacts
    impact = calculate_esg_impact()
    
    # Comparison table
    st.markdown("### Cost per Ton Comparison")
    
    comparison_data = pd.DataFrame([
        {'Initiative': 'Solar PV', 'Quantity': str(impact['solar']['qty']), 
         'Total Cost': f"${impact['solar']['cost']:,.0f}", 
         'CO2 Reduced': f"{impact['solar']['reduction']:.1f} tons",
         'Cost/Ton': f"${impact['solar']['cost_per_ton']:,.0f}" if impact['solar']['cost_per_ton'] > 0 else '-',
         'Payback': f"{impact['solar']['payback']:.1f} yrs" if impact['solar']['payback'] < 999 else '-'},
        {'Initiative': 'Trees', 'Quantity': str(impact['trees']['qty']),
         'Total Cost': f"${impact['trees']['cost']:,.0f}",
         'CO2 Reduced': f"{impact['trees']['reduction']:.1f} tons",
         'Cost/Ton': f"${impact['trees']['cost_per_ton']:,.0f}" if impact['trees']['cost_per_ton'] > 0 else '-',
         'Payback': f"{impact['trees']['payback']:.1f} yrs" if impact['trees']['payback'] < 999 else '-'},
        {'Initiative': 'Green Electricity', 'Quantity': f"{impact['green_elec']['pct']:.0f}%",
         'Total Cost': f"${impact['green_elec']['cost']:,.0f}",
         'CO2 Reduced': f"{impact['green_elec']['reduction']:.1f} tons",
         'Cost/Ton': f"${impact['green_elec']['cost_per_ton']:,.0f}" if impact['green_elec']['cost_per_ton'] > 0 else '-',
         'Payback': 'N/A (OpEx)'},
        {'Initiative': 'CO2 Credits', 'Quantity': str(impact['credits']['qty']),
         'Total Cost': f"${impact['credits']['cost']:,.0f}",
         'CO2 Reduced': f"{impact['credits']['reduction']:.1f} tons",
         'Cost/Ton': f"${impact['credits']['cost_per_ton']:,.0f}",
         'Payback': 'N/A (OpEx)'}
    ])
    
    st.dataframe(comparison_data, width='stretch', hide_index=True)


def render_results():
    """Render RESULTS sub-tab."""
    st.subheader("📊 RESULTS - ESG Summary")
    
    impact = calculate_esg_impact()
    
    # Key metrics
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total CO2 Reduction", f"{impact['total_reduction']:.1f} tons")
    with col2:
        st.metric("Remaining Emissions", f"{impact['remaining_emissions']:.1f} tons")
    with col3:
        st.metric("Projected Tax Liability", f"${impact['tax_liability']:,.0f}")
    
    # Investment summary
    st.markdown("### Investment Summary")
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Total CAPEX (One-Time)", f"${impact['total_capex']:,.0f}")
    with col2:
        st.metric("Total OpEx (Annual)", f"${impact['total_opex']:,.0f}")
    
    set_state('ESG_CAPEX', impact['total_capex'])
    set_state('ESG CapEx', impact['total_capex'])
    
    # Recommendations
    st.markdown("### 💡 Recommendations")
    
    if impact['solar']['payback'] < 3:
        st.success("✅ Solar PV has payback < 3 years - Good investment!")
    elif st.session_state.esg_solar_panels > 0:
        st.warning(f"⚠️ Solar payback is {impact['solar']['payback']:.1f} years - Consider reducing panels")
    
    if impact['remaining_emissions'] > 0:
        credits_needed = int(impact['remaining_emissions'])
        credits_cost = credits_needed * ESG_INITIATIVES['CO2_Credits']['cost_per_ton']
        st.info(f"ℹ️ To offset remaining {impact['remaining_emissions']:.1f} tons: Buy {credits_needed} credits (${credits_cost:,.0f})")
    
    # Pie chart
    if impact['total_reduction'] > 0:
        reduction_data = pd.DataFrame({
            'Source': ['Solar', 'Trees', 'Green Elec', 'Credits'],
            'Reduction': [impact['solar']['reduction'], impact['trees']['reduction'],
                         impact['green_elec']['reduction'], impact['credits']['reduction']]
        })
        reduction_data = reduction_data[reduction_data['Reduction'] > 0]
        
        if not reduction_data.empty:
            fig = px.pie(
                reduction_data,
                values='Reduction',
                names='Source',
                title='CO2 Reduction by Initiative',
                color_discrete_sequence=px.colors.qualitative.Set2
            )
            fig.update_layout(height=350)
            st.plotly_chart(fig, width='stretch')


def render_upload_ready_esg():
    """Render UPLOAD_READY_ESG sub-tab."""
    st.subheader("📤 UPLOAD READY - ESG Decisions")
    
    st.info("Copy these values to ExSim ESG Decision Form")
    
    impact = calculate_esg_impact()
    
    # Decisions summary
    st.markdown("### 🌱 ESG Decisions")
    
    decisions = []
    if st.session_state.esg_solar_panels > 0:
        decisions.append({'Initiative': 'Solar PV Panels', 'Quantity': str(st.session_state.esg_solar_panels), 
                         'Cost': f"${impact['solar']['cost']:,.0f}"})
    if st.session_state.esg_trees > 0:
        decisions.append({'Initiative': 'Trees', 'Quantity': str(st.session_state.esg_trees),
                         'Cost': f"${impact['trees']['cost']:,.0f}"})
    if st.session_state.esg_green_electricity_pct > 0:
        decisions.append({'Initiative': 'Green Electricity', 'Quantity': f"{st.session_state.esg_green_electricity_pct}%",
                         'Cost': f"${impact['green_elec']['cost']:,.0f}"})
    if st.session_state.esg_co2_credits > 0:
        decisions.append({'Initiative': 'CO2 Credits', 'Quantity': str(st.session_state.esg_co2_credits),
                         'Cost': f"${impact['credits']['cost']:,.0f}"})
    
    if decisions:
        st.dataframe(pd.DataFrame(decisions), hide_index=True, width='stretch')
    else:
        st.caption("No ESG initiatives selected")
    
    # Totals
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Total CAPEX", f"${impact['total_capex']:,.0f}")
    with col2:
        st.metric("Total OpEx", f"${impact['total_opex']:,.0f}")
    
    st.metric("Projected Remaining Tax", f"${impact['tax_liability']:,.0f}")
    
    # CSV download button
    if decisions:
        export_df = pd.DataFrame(decisions)
        csv_data = export_df.to_csv(index=False)
        st.download_button(
            label="📥 Download Decisions as CSV",
            data=csv_data,
            file_name="esg_decisions.csv",
            mime="text/csv",
            type="primary",
            key='esg_csv_download'
        )
    else:
        st.caption("No decisions to download")


def render_cross_reference():
    """Render CROSS_REFERENCE sub-tab - Upstream data visibility."""
    st.subheader("🔗 CROSS REFERENCE - Upstream Support")
    st.caption("Live visibility into Production impact and Logistics miles.")
    
    # Load shared data
    try:
        from shared_outputs import import_dashboard_data
        prod_data = import_dashboard_data('Production') or {}
        clo_data = import_dashboard_data('CLO') or {}
    except ImportError:
        st.error("Could not load shared_outputs module")
        prod_data = {}
        clo_data = {}
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 🏭 Production (Output)")
        st.info("Higher output = Higher Energy Consumption.")
        
        # Extract Production Plan Target Sum
        try:
            prod_plan = prod_data.get('production_plan', {})
            total_target = sum([d.get('Target', 0) for d in prod_plan.values()]) if isinstance(prod_plan, dict) else 0
            utilization = prod_data.get('capacity_utilization', {}).get('mean', 0)
        except:
            total_target = 0
            utilization = 0
            
        st.metric("Total Production", f"{total_target:,.0f} units")
        st.metric("Avg Utilization", f"{utilization*100:.1f}%")

    with col2:
        st.markdown("### 🚚 Logistics (Scope 3)")
        st.info("Shipping Volume drives Carbon Footprint.")
        
        logistics_cost = clo_data.get('logistics_costs', 0)
        
        st.metric("Logistics Cost Proxy", f"${logistics_cost:,.0f}")
        
        if logistics_cost > 100000:
            st.warning("⚠️ High Logistics Activity - Check Scope 3 Emissions")
        elif logistics_cost > 0:
            st.info("ℹ️ Moderate Logistics Activity")
        else:
            st.caption("No logistics data")


def render_esg_tab():
    """Render the ESG tab with 4 Excel-aligned subtabs."""
    init_esg_state()
    
    # Header with Download Button
    col_header, col_download = st.columns([4, 1])
    with col_header:
        st.header("🌿 ESG Dashboard - Sustainability & CO2 Abatement")
    with col_download:
        try:
            from utils.report_bridge import create_download_button
            create_download_button('ESG', 'ESG')
        except Exception as e:
            st.error(f"Export: {e}")
    
    # Quick summary
    impact = calculate_esg_impact()
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Current Emissions", f"{st.session_state.esg_current_emissions} tons")
    with col2:
        st.metric("Total Reduction", f"{impact['total_reduction']:.1f} tons")
    with col3:
        st.metric("Remaining", f"{impact['remaining_emissions']:.1f} tons")
    with col4:
        st.metric("Tax Liability", f"${impact['tax_liability']:,.0f}")
    
    # 5 SUBTABS (Updated)
    subtabs = st.tabs([
        "⚙️ Impact Config",
        "🌱 Strategy Selector",
        "📊 Results",
        "📤 Upload Ready",
        "🔗 Cross Reference"
    ])
    
    with subtabs[0]:
        render_impact_config()
    
    with subtabs[1]:
        render_strategy_selector()
    
    with subtabs[2]:
        render_results()
    
    with subtabs[3]:
        render_upload_ready_esg()
        
    with subtabs[4]:
        render_cross_reference()
