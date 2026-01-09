"""
ExSim War Room - ESG Tab
5 sub-tabs mirroring the Excel dashboard sheets:
1. IMPACT_CONFIG - CO2 tax rates and initiative settings
2. STRATEGY_SELECTOR - Compare green investment options
3. RESULTS - Summary and recommendations
4. UPLOAD_READY_ESG - Export preview
5. CROSS_REFERENCE - Upstream data visibility
Plus new sections: Machine CO2, Transport CO2, Emissions Intensity, Product Improvements
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode

from utils.state_manager import get_state, set_state

# Import ESG parameters from case_parameters
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent.parent.parent))
try:
    from case_parameters import ESG as ESG_PARAMS
except ImportError:
    ESG_PARAMS = None

# Build ESG_INITIATIVES from case_parameters (with fallback)
if ESG_PARAMS:
    abatement = ESG_PARAMS.get("ABATEMENT", {})
    solar = abatement.get("SOLAR_PANELS", {})
    green = abatement.get("GREEN_ENERGY", {})
    trees = abatement.get("TREES", {})
    credits = abatement.get("CO2_CREDITS", {})
    
    ESG_INITIATIVES = {
        'Solar_PV': {
            'cost_per_unit': solar.get('cost', 420),
            'maintenance_per_period': solar.get('maintenance_per_period', 7),
            'co2_reduction_per_period_kg': solar.get('co2_reduction_per_period_kg', 106.4),
            'type': 'CAPEX',
            'unit': 'panel'
        },
        'Trees': {
            'cost_per_unit': trees.get('cost_per_tree', 6.25),
            'maintenance_per_period_per_80': trees.get('maintenance_per_period_per_80_trees', 16.67),
            'co2_reduction_per_period_per_80_kg': trees.get('co2_absorbed_per_period_per_80_trees_kg', 333),
            'type': 'CAPEX',
            'unit': 'tree'
        },
        'Green_Electricity': {
            'regular_price_kwh': green.get('regular_price_per_kwh', 0.06),
            'premium_rate': green.get('premium_rate', 0.20),
            'co2_per_kwh_kg': green.get('co2_per_kwh_reduction', 0.4),
            'type': 'OpEx',
            'unit': 'kWh'
        },
        'CO2_Credits': {
            'co2_per_credit_kg': credits.get('co2_per_credit_kg', 1000),
            'type': 'OpEx',
            'unit': 'credit'
        }
    }
    MACHINE_CO2 = ESG_PARAMS.get("MACHINE_CO2", {})
    TRANSPORT_CO2 = ESG_PARAMS.get("TRANSPORT_CO2", {})
    IMPROVEMENT_CO2 = ESG_PARAMS.get("IMPROVEMENT_CO2", {})
    TARGETS = ESG_PARAMS.get("TARGETS", {})
else:
    # Fallback defaults (should not be used)
    ESG_INITIATIVES = {
        'Solar_PV': {'cost_per_unit': 420, 'co2_reduction_per_period_kg': 106.4, 'type': 'CAPEX', 'unit': 'panel'},
        'Trees': {'cost_per_unit': 6.25, 'co2_reduction_per_period_per_80_kg': 333, 'type': 'CAPEX', 'unit': 'tree'},
        'Green_Electricity': {'premium_rate': 0.20, 'co2_per_kwh_kg': 0.4, 'type': 'OpEx', 'unit': 'kWh'},
        'CO2_Credits': {'co2_per_credit_kg': 1000, 'type': 'OpEx', 'unit': 'credit'}
    }
    MACHINE_CO2 = {}
    TRANSPORT_CO2 = {}
    IMPROVEMENT_CO2 = {}
    TARGETS = {"ANNUAL_CO2_REDUCTION": 0.15, "PERIOD_6_INTENSITY": 29.93}




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
    """Calculate ESG initiative costs and impacts using case_parameters values."""
    tax_rate = st.session_state.esg_co2_tax_rate
    emissions = st.session_state.esg_current_emissions  # in tons
    energy = st.session_state.esg_energy_consumption  # in kWh
    
    # Solar PV - $420/panel, 106.4 kg CO2 reduction/period
    solar_qty = st.session_state.esg_solar_panels
    solar_upfront = solar_qty * ESG_INITIATIVES['Solar_PV']['cost_per_unit']
    solar_maintenance = solar_qty * ESG_INITIATIVES['Solar_PV'].get('maintenance_per_period', 7)
    solar_cost = solar_upfront + solar_maintenance * 3  # 3 periods per year
    solar_reduction_kg = solar_qty * ESG_INITIATIVES['Solar_PV'].get('co2_reduction_per_period_kg', 106.4) * 3
    solar_reduction = solar_reduction_kg / 1000  # Convert to tons
    solar_annual_savings = solar_reduction * tax_rate
    solar_payback = solar_upfront / solar_annual_savings if solar_annual_savings > 0 else 999
    solar_cost_per_ton = solar_cost / solar_reduction if solar_reduction > 0 else 0
    
    # Trees - $6.25/tree, 333 kg CO2/period per 80 trees
    trees_qty = st.session_state.esg_trees
    trees_upfront = trees_qty * ESG_INITIATIVES['Trees']['cost_per_unit']
    trees_per_80 = trees_qty / 80 if trees_qty > 0 else 0
    trees_maintenance = trees_per_80 * ESG_INITIATIVES['Trees'].get('maintenance_per_period_per_80', 16.67) * 3
    trees_cost = trees_upfront + trees_maintenance
    trees_reduction_kg = trees_per_80 * ESG_INITIATIVES['Trees'].get('co2_reduction_per_period_per_80_kg', 333) * 3
    trees_reduction = trees_reduction_kg / 1000  # Convert to tons
    trees_annual_savings = trees_reduction * tax_rate
    trees_payback = trees_upfront / trees_annual_savings if trees_annual_savings > 0 else 999
    trees_cost_per_ton = trees_cost / trees_reduction if trees_reduction > 0 else 0
    
    # Green Electricity - 20% premium over $0.06/kWh, 0.4 kg CO2/kWh reduction
    green_pct = st.session_state.esg_green_electricity_pct / 100
    green_kwh = energy * green_pct
    regular_price = ESG_INITIATIVES['Green_Electricity'].get('regular_price_kwh', 0.06)
    premium_rate = ESG_INITIATIVES['Green_Electricity'].get('premium_rate', 0.20)
    green_premium_cost = green_kwh * regular_price * premium_rate  # Only the premium portion
    green_cost = green_premium_cost
    green_reduction_kg = green_kwh * ESG_INITIATIVES['Green_Electricity'].get('co2_per_kwh_kg', 0.4)
    green_reduction = green_reduction_kg / 1000  # Convert to tons
    green_cost_per_ton = green_cost / green_reduction if green_reduction > 0 else 0
    
    # CO2 Credits - 1 ton = 1000 kg per credit
    credits_qty = st.session_state.esg_co2_credits
    credit_cost_per_ton = st.session_state.get('esg_credit_price', 25)  # User-adjustable
    credits_cost = credits_qty * credit_cost_per_ton
    credits_reduction = credits_qty  # 1 credit = 1 ton
    credits_cost_per_ton = credit_cost_per_ton
    
    # Totals
    total_reduction = solar_reduction + trees_reduction + green_reduction + credits_reduction
    remaining_emissions = max(0, emissions - total_reduction)
    tax_liability = remaining_emissions * tax_rate
    
    total_capex = solar_upfront + trees_upfront
    total_opex = solar_maintenance * 3 + trees_maintenance + green_cost + credits_cost
    
    return {
        'solar': {'qty': solar_qty, 'cost': solar_cost, 'reduction': solar_reduction, 
                 'payback': solar_payback, 'cost_per_ton': solar_cost_per_ton, 'upfront': solar_upfront},
        'trees': {'qty': trees_qty, 'cost': trees_cost, 'reduction': trees_reduction,
                 'payback': trees_payback, 'cost_per_ton': trees_cost_per_ton, 'upfront': trees_upfront},
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
    
    # Initiative settings table from case_parameters
    st.markdown("### Initiative Settings (from Case Table VII.2)")
    
    settings_df = pd.DataFrame([
        {'Initiative': 'Solar PV', 'Type': 'CAPEX', 'Cost': '$420/panel + $7/period maint', 'CO2 Reduction': '106.4 kg/panel/period'},
        {'Initiative': 'Trees', 'Type': 'CAPEX', 'Cost': '$6.25/tree + $16.67/period per 80', 'CO2 Reduction': '333 kg/period per 80 trees'},
        {'Initiative': 'Green Electricity', 'Type': 'OpEx', 'Cost': '20% premium ($0.012/kWh)', 'CO2 Reduction': '0.4 kg CO2/kWh'},
        {'Initiative': 'CO2 Credits', 'Type': 'OpEx', 'Cost': 'Variable (set below)', 'CO2 Reduction': '1 ton (1000 kg)/credit'}
    ])
    
    st.dataframe(settings_df, use_container_width=True, hide_index=True)
    
    # Add credit price input
    credit_price = st.number_input(
        "CO2 Credit Price ($/ton)", 
        value=st.session_state.get('esg_credit_price', 25),
        step=5,
        key='esg_credit_price_input'
    )
    st.session_state.esg_credit_price = credit_price
    
    # Machine CO2 Emissions Reference (Table VII.1)
    st.markdown("### 🏭 Machine CO2 Emissions (Table VII.1)")
    if MACHINE_CO2:
        machine_data = []
        for machine, data in MACHINE_CO2.items():
            machine_data.append({
                'Machine': machine,
                'Emissions @ Capacity': f"{data['emissions_at_capacity_kg']} kg/fortnight",
                'Capacity': f"{data['capacity']} units/fortnight",
                'kg CO2/unit': f"{data['kg_per_unit']:.2f}"
            })
        st.dataframe(pd.DataFrame(machine_data), use_container_width=True, hide_index=True)
    else:
        st.caption("Machine CO2 data not loaded from case_parameters")
    
    # Product Improvement CO2 Impact (Table VII.1)
    st.markdown("### 📦 Product Improvement CO2 Impact")
    if IMPROVEMENT_CO2:
        improv_data = []
        for id, data in IMPROVEMENT_CO2.items():
            co2 = data['kg_co2']
            impact = "🌿 Reduces" if co2 < 0 else "🔴 Increases"
            improv_data.append({
                'ID': id,
                'Feature': data['name'],
                'Impact': impact,
                'kg CO2/unit': f"{co2:+.3f}"
            })
        st.dataframe(pd.DataFrame(improv_data), use_container_width=True, hide_index=True)
    else:
        st.caption("Improvement CO2 data not loaded")



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
        credit_price = st.session_state.get('esg_credit_price', 25)
        credits_cost = credits_needed * credit_price
        st.info(f"ℹ️ To offset remaining {impact['remaining_emissions']:.1f} tons: Buy {credits_needed} credits (${credits_cost:,.0f})")
    
    # Emissions Intensity Tracker (new section)
    st.markdown("### 📈 Emissions Intensity Tracker")
    
    # Get production estimate from shared outputs or use default
    try:
        from shared_outputs import import_dashboard_data
        prod_data = import_dashboard_data('Production') or {}
        prod_plan = prod_data.get('production_plan', {})
        # Type safety: convert to float (JSON may serialize as strings)
        total_production = sum([float(d.get('Target', 0)) if d.get('Target') else 0 for d in prod_plan.values()]) if isinstance(prod_plan, dict) else 0
    except:
        total_production = 0
    
    if total_production == 0:
        total_production = st.number_input(
            "Est. Annual Production (units)", 
            value=50000, step=5000,
            help="Enter estimated production for intensity calculation"
        )
    
    base_emissions_tons = st.session_state.esg_current_emissions
    base_intensity = (base_emissions_tons * 1000) / total_production if total_production > 0 else 0
    net_emissions = impact['remaining_emissions']
    net_intensity = (net_emissions * 1000) / total_production if total_production > 0 else 0
    
    period_6_baseline = TARGETS.get('PERIOD_6_INTENSITY', 29.93)
    annual_target = TARGETS.get('ANNUAL_CO2_REDUCTION', 0.15)
    target_intensity = period_6_baseline * (1 - annual_target)  # 15% reduction
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Current Intensity", f"{base_intensity:.2f} kg/unit")
    with col2:
        delta = net_intensity - base_intensity
        st.metric("Net Intensity", f"{net_intensity:.2f} kg/unit", delta=f"{delta:.2f}")
    with col3:
        st.metric("Period Target", f"{target_intensity:.2f} kg/unit", help="15% reduction from Period 6 baseline")
    
    if net_intensity <= target_intensity:
        st.success(f"✅ On track to meet 15% reduction target! Current: {net_intensity:.2f} vs Target: {target_intensity:.2f}")
    else:
        gap = net_intensity - target_intensity
        st.warning(f"⚠️ Above target by {gap:.2f} kg/unit. Consider more abatement actions.")
    
    # ESG KPI Score Estimate
    st.markdown("### 🏆 ESG KPI Score Estimate")
    
    # Calculate a rough ENV KPI based on emissions reduction and target achievement
    reduction_pct = (impact['total_reduction'] / base_emissions_tons * 100) if base_emissions_tons > 0 else 0
    intensity_achievement = (1 - net_intensity / period_6_baseline) * 100 if period_6_baseline > 0 else 0
    
    # Simple scoring: 60 base + up to 20 for reduction + up to 20 for intensity
    env_score = min(100, 60 + (reduction_pct * 0.5) + (intensity_achievement * 0.5))
    
    # Get social score from CPO if available
    try:
        from shared_outputs import import_dashboard_data
        cpo_data = import_dashboard_data('CPO') or {}
        social_score = cpo_data.get('workforce_mood', 70)
    except:
        social_score = 70  # Default
    
    esg_score = (env_score * 0.6 + social_score * 0.4)  # 60% environmental, 40% social
    
    col1, col2, col3 = st.columns(3)
    with col1:
        color = "🟢" if env_score >= 80 else ("🟡" if env_score >= 60 else "🔴")
        st.metric(f"{color} ENV KPI", f"{env_score:.0f}%")
    with col2:
        color = "🟢" if social_score >= 80 else ("🟡" if social_score >= 60 else "🔴")
        st.metric(f"{color} Social KPI", f"{social_score:.0f}%")
    with col3:
        color = "🟢" if esg_score >= 80 else ("🟡" if esg_score >= 60 else "🔴")
        st.metric(f"{color} ESG KPI", f"{esg_score:.0f}%")
    
    st.caption("*KPI scores are estimates. Actual scores depend on competitor performance and industry benchmarks.*")
    
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
            st.plotly_chart(fig, use_container_width=True)



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
            # Type safety: convert to float (JSON may serialize as strings)
            total_target = sum([float(d.get('Target', 0)) if d.get('Target') else 0 for d in prod_plan.values()]) if isinstance(prod_plan, dict) else 0
            utilization = float(prod_data.get('capacity_utilization', {}).get('mean', 0) or 0)
        except:
            total_target = 0
            utilization = 0
            
        st.metric("Total Production", f"{total_target:,.0f} units")
        st.metric("Avg Utilization", f"{utilization*100:.1f}%")

    with col2:
        st.markdown("### 🚚 Logistics (Scope 3)")
        st.info("Shipping Volume drives Carbon Footprint.")
        
        raw_cost = clo_data.get('logistics_costs', 0)
        try:
            logistics_cost = float(raw_cost)
        except (ValueError, TypeError):
            logistics_cost = 0
        
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
