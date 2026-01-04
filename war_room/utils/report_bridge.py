"""
Report Bridge - Connects Streamlit live data to Excel dashboard generators.

This module bridges the session state data from the Streamlit app to the 
openpyxl-based dashboard generators, enabling "Live Report Export" functionality.

Now includes:
- Shared outputs integration for cross-dashboard data
- Auto-sync of session state to shared_outputs.json
- Export methods for all 7 dashboards
"""

import io
import streamlit as st
import sys
from pathlib import Path

# Add all dashboard directories to path for imports
PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
sys.path.insert(0, str(PROJECT_ROOT / "CFO Dashboard"))
sys.path.insert(0, str(PROJECT_ROOT / "CLO Dashboard"))
sys.path.insert(0, str(PROJECT_ROOT / "CMO Dashboard"))
sys.path.insert(0, str(PROJECT_ROOT / "Production Manager Dashboard"))
sys.path.insert(0, str(PROJECT_ROOT / "Purchasing Dashboard"))
sys.path.insert(0, str(PROJECT_ROOT / "People Dashboard"))
sys.path.insert(0, str(PROJECT_ROOT / "ESG Dashboard"))

# Import shared outputs for cross-dashboard communication
try:
    from shared_outputs import export_dashboard_data, import_dashboard_data, SharedOutputManager
except ImportError:
    export_dashboard_data = None
    import_dashboard_data = None
    SharedOutputManager = None

# Import COMMON parameters
try:
    from case_parameters import COMMON
    ZONES = COMMON.get('ZONES', ['Center', 'West', 'North', 'East', 'South'])
    FORTNIGHTS = COMMON.get('FORTNIGHTS', list(range(1, 9)))
except ImportError:
    ZONES = ['Center', 'West', 'North', 'East', 'South']
    FORTNIGHTS = list(range(1, 9))


class ReportBridge:
    """Bridge between Streamlit session state and Excel dashboard generators."""
    
    # =========================================================================
    # AUTO-SYNC: Export session state to shared_outputs.json
    # =========================================================================
    
    @staticmethod
    def sync_cfo_to_shared():
        """Sync CFO session state to shared_outputs.json"""
        if not export_dashboard_data:
            return
        
        cash_data = {
            'final_cash': st.session_state.get('cfo_cash_at_end_last_period', 0),
            'tax_payments': st.session_state.get('cfo_tax_payments', 0)
        }
        
        export_dashboard_data('CFO', {
            'cash_flow_projection': cash_data,
            'debt_levels': st.session_state.get('cfo_total_liabilities', 0),
            'liquidity_status': 'OK' if cash_data.get('final_cash', 0) > 0 else 'LOW'
        })
    
    @staticmethod
    def sync_cmo_to_shared():
        """Sync CMO session state to shared_outputs.json"""
        if not export_dashboard_data:
            return
        
        export_dashboard_data('CMO', {
            'demand_forecast': {zone: st.session_state.get(f'cmo_demand_{zone}', 0) for zone in ZONES},
            'marketing_spend': st.session_state.get('cmo_marketing_spend', 0),
            'pricing': {zone: st.session_state.get(f'cmo_price_{zone}', 0) for zone in ZONES},
            'innovation_costs': st.session_state.get('cmo_innovation_costs', 0)
        })
    
    @staticmethod
    def sync_production_to_shared():
        """Sync Production session state to shared_outputs.json"""
        if not export_dashboard_data:
            return
        
        export_dashboard_data('Production', {
            'production_plan': {zone: st.session_state.get(f'prod_units_{zone}', 0) for zone in ZONES},
            'capacity_utilization': st.session_state.get('prod_capacity_util', 0.85),
            'overtime_hours': st.session_state.get('prod_overtime', 0),
            'unit_costs': {zone: st.session_state.get(f'prod_unit_cost_{zone}', 40) for zone in ZONES}
        })
    
    @staticmethod
    def sync_logistics_to_shared():
        """Sync Logistics session state to shared_outputs.json"""
        if not export_dashboard_data:
            return
        
        inventory_df = st.session_state.get('logistics_inventory')
        inv_by_zone = {}
        
        if inventory_df is not None:
            for _, row in inventory_df.iterrows():
                inv_by_zone[row['Zone']] = row.get('Initial_Inv', 0)
        else:
            inv_by_zone = {zone: 0 for zone in ZONES}
        
        export_dashboard_data('CLO', {
            'shipping_schedule': {},
            'logistics_costs': st.session_state.get('logistics_cost', 0),
            'inventory_by_zone': inv_by_zone
        })
    
    @staticmethod
    def sync_cpo_to_shared():
        """Sync CPO session state to shared_outputs.json"""
        if not export_dashboard_data:
            return
        
        export_dashboard_data('CPO', {
            'workforce_headcount': {zone: st.session_state.get(f'cpo_headcount_{zone}', 0) for zone in ZONES},
            'payroll_forecast': st.session_state.get('cpo_payroll', 0),
            'hiring_costs': st.session_state.get('cpo_hiring_costs', 0)
        })
    
    @staticmethod
    def sync_purchasing_to_shared():
        """Sync Purchasing session state to shared_outputs.json"""
        if not export_dashboard_data:
            return
        
        export_dashboard_data('Purchasing', {
            'material_orders': {},
            'supplier_spend': st.session_state.get('purchasing_spend', 0),
            'lead_time_schedule': {}
        })
    
    @staticmethod
    def sync_esg_to_shared():
        """Sync ESG session state to shared_outputs.json"""
        if not export_dashboard_data:
            return
        
        export_dashboard_data('ESG', {
            'co2_emissions': st.session_state.get('esg_emissions', 0),
            'abatement_investment': st.session_state.get('esg_abatement', 0),
            'tax_liability': st.session_state.get('esg_tax', 0)
        })
    
    # =========================================================================
    # EXPORT METHODS: Generate Excel files from session state
    # =========================================================================
    
    @staticmethod
    def export_finance_dashboard() -> io.BytesIO:
        """Export CFO Finance Dashboard using live Streamlit data."""
        from generate_finance_dashboard_final import create_finance_dashboard
        
        # Sync to shared outputs first
        ReportBridge.sync_cfo_to_shared()
        
        # Extract operational cash flow data
        cash_flow_df = st.session_state.get('cfo_cash_flow')
        
        # Build cash_data dict from session state
        cash_data = {
            'final_cash': st.session_state.get('cfo_cash_at_end_last_period', 0),
            'tax_payments': st.session_state.get('cfo_tax_payments', 0)
        }
        
        # Build balance_data dict from session state
        balance_data = {
            'net_sales': st.session_state.get('cfo_net_sales', 0),
            'cogs': st.session_state.get('cfo_cogs', 0),
            'gross_income': st.session_state.get('cfo_gross_margin', 0),
            'net_profit': st.session_state.get('cfo_net_profit', 0),
            'total_assets': st.session_state.get('cfo_total_assets', 0),
            'total_liabilities': st.session_state.get('cfo_total_liabilities', 0),
            'equity': (st.session_state.get('cfo_total_assets', 0) 
                      - st.session_state.get('cfo_total_liabilities', 0)),
            'retained_earnings': st.session_state.get('cfo_retained_earnings', 0),
            'depreciation': 0,
            'gross_margin_pct': st.session_state.get('cfo_gross_margin_pct', 0.4),
            'net_margin_pct': st.session_state.get('cfo_net_margin_pct', 0.1)
        }
        
        sa_data = {'total_sa_expenses': 200000}
        ar_ap_data = {'receivables': [0]*8, 'payables': [0]*8}
        
        if cash_flow_df is not None:
            for fn in range(1, 9):
                fn_col = f'FN{fn}'
                if fn_col in cash_flow_df.columns:
                    ar_ap_data['receivables'][fn-1] = float(cash_flow_df.at[3, fn_col]) if 3 in cash_flow_df.index else 0
                    ar_ap_data['payables'][fn-1] = abs(float(cash_flow_df.at[4, fn_col])) if 4 in cash_flow_df.index else 0
        
        starting_cash = (cash_data['final_cash'] 
                        - cash_data['tax_payments'] 
                        - st.session_state.get('cfo_dividend_payments', 0)
                        - st.session_state.get('cfo_asset_purchases', 0))
        
        hard_data = {
            'depreciation': 0,
            'starting_cash': starting_cash,
            'schedule': {fn: {'receivables': ar_ap_data['receivables'][fn-1], 
                              'payables': ar_ap_data['payables'][fn-1]} for fn in range(1, 9)},
            'retained_earnings': balance_data['retained_earnings']
        }
        
        output = io.BytesIO()
        create_finance_dashboard(
            cash_data=cash_data,
            balance_data=balance_data,
            sa_data=sa_data,
            ar_ap_data=ar_ap_data,
            template_data=None,
            hard_data=hard_data,
            output_buffer=output
        )
        return output
    
    @staticmethod
    def export_logistics_dashboard() -> io.BytesIO:
        """Export CLO Logistics Dashboard using live Streamlit data."""
        from generate_logistics_dashboard import create_logistics_dashboard
        
        # Sync to shared outputs first
        ReportBridge.sync_logistics_to_shared()
        
        inventory_df = st.session_state.get('logistics_inventory')
        warehouses_df = st.session_state.get('logistics_warehouses')
        
        inventory_data = {}
        
        if inventory_df is not None and warehouses_df is not None:
            for zone in ZONES:
                zone_inv_rows = inventory_df[inventory_df['Zone'] == zone]
                zone_wh_rows = warehouses_df[warehouses_df['Zone'] == zone]
                
                initial_inv = zone_inv_rows.iloc[0].get('Initial_Inv', 0) if not zone_inv_rows.empty else 0
                capacity = zone_wh_rows.iloc[0].get('Total_Capacity', 1000) if not zone_wh_rows.empty else 1000
                
                inventory_data[zone] = {'inventory': initial_inv, 'capacity': capacity}
        else:
            for zone in ZONES:
                inventory_data[zone] = {'inventory': 500, 'capacity': 1000}
        
        template_data = {'df': None, 'exists': False}
        cost_data = {'total_shipping_cost': 0}
        
        output = io.BytesIO()
        create_logistics_dashboard(
            inventory_data=inventory_data,
            template_data=template_data,
            cost_data=cost_data,
            intelligence_data=None,
            output_buffer=output
        )
        return output
    
    @staticmethod
    def export_cmo_dashboard() -> io.BytesIO:
        """Export CMO Marketing Dashboard (placeholder - returns empty buffer if generator not available)."""
        ReportBridge.sync_cmo_to_shared()
        
        try:
            from generate_cmo_dashboard_complete import (
                create_complete_dashboard,
                load_market_report,
                load_innovation_features,
                load_marketing_template,
                load_sales_admin_expenses,
                load_finished_goods_inventory,
                load_marketing_intelligence,
                get_data_path, REQUIRED_FILES
            )
            
            # Load Base Data
            market = load_market_report(get_data_path(REQUIRED_FILES[2]))
            innov_features = load_innovation_features(get_data_path(REQUIRED_FILES[1]))
            template = load_marketing_template(get_data_path(REQUIRED_FILES[0]))
            sales = load_sales_admin_expenses(get_data_path(REQUIRED_FILES[3]))
            inv = load_finished_goods_inventory(get_data_path(REQUIRED_FILES[4]))
            intel = load_marketing_intelligence(get_data_path(REQUIRED_FILES[3]), get_data_path(REQUIRED_FILES[2]))
            
            # Build Overrides
            overrides = {
                'innovation': {},
                'zones': {}
            }
            
            # 1. Innovation
            innov_decisions = st.session_state.get('cmo_innovation_decisions', {})
            for feature, data in innov_decisions.items():
                if data['decision'] == 1:
                    overrides['innovation'][feature] = 1
            
            # 2. Global
            overrides['tv_spots'] = st.session_state.get('cmo_tv_spots', 10)
            overrides['brand_focus'] = st.session_state.get('cmo_brand_focus', 50)
            
            # 3. Zones
            strategy_df = st.session_state.get('cmo_strategy_inputs')
            if strategy_df is not None:
                for _, row in strategy_df.iterrows():
                    zone = row['Zone']
                    overrides['zones'][zone] = {
                        'target_demand': row['Target_Demand'],
                        'radio': row['Radio_Spots'],
                        'salespeople': row['Headcount'],
                        'price': row['Price']
                    }
                    
            output = io.BytesIO()
            create_complete_dashboard(
                market, innov_features, template, sales, inv, intel,
                output_buffer=output,
                decision_overrides=overrides
            )
            return output
        except ImportError as e:
            st.error(f"Generate Error: {e}")
            return io.BytesIO()
        except Exception as e:
            st.error(f"Export Error: {e}")
            return io.BytesIO()
    
    @staticmethod
    def export_production_dashboard() -> io.BytesIO:
        """Export Production Dashboard (placeholder)."""
        ReportBridge.sync_production_to_shared()
        
        try:
            from generate_production_dashboard_zones import (
                create_zones_dashboard, 
                load_raw_materials_by_zone, 
                load_finished_goods_by_zone,
                load_workers_by_zone,
                load_machines_by_zone,
                load_production_template,
                get_data_path, REQUIRED_FILES
            )
            
            # Load Base Data
            materials = load_raw_materials_by_zone(get_data_path(REQUIRED_FILES[0]))
            fg = load_finished_goods_by_zone(get_data_path(REQUIRED_FILES[1]))
            workers = load_workers_by_zone(get_data_path(REQUIRED_FILES[2]))
            machines = load_machines_by_zone(get_data_path(REQUIRED_FILES[3]))
            template = load_production_template(get_data_path(REQUIRED_FILES[4]))
            
            # Build Overrides
            # Production tab often uses 'prod_units_{zone}' for target
            # And 'prod_overtime_{zone}' for overtime boolean
            decisions_override = {
                'targets': {},
                'overtime': {}
            }
            
            for zone in ZONES:
                decisions_override['targets'][zone] = st.session_state.get(f'prod_units_{zone}', 0)
                is_ot = st.session_state.get(f'prod_overtime_{zone}', False)
                decisions_override['overtime'][zone] = 'Y' if is_ot else 'N'
            
            output = io.BytesIO()
            create_zones_dashboard(
                materials, fg, workers, machines, template, 
                output_buffer=output,
                decision_overrides=decisions_override
            )
            return output
        except ImportError as e:
            st.error(f"Generate Error: {e}")
            return io.BytesIO()
        except Exception as e:
            st.error(f"Export Error: {e}")
            return io.BytesIO()
    
    @staticmethod
    def export_purchasing_dashboard() -> io.BytesIO:
        """Export Purchasing Dashboard."""
        ReportBridge.sync_purchasing_to_shared()
        
        try:
            from generate_purchasing_dashboard_v2 import (
                create_purchasing_dashboard,
                load_raw_materials,
                load_production_costs,
                load_procurement_template,
                get_data_path, REQUIRED_FILES
            )
            
            # Load Base Data
            materials = load_raw_materials(get_data_path(REQUIRED_FILES[0]))
            costs = load_production_costs(get_data_path(REQUIRED_FILES[1]))
            # Template not strictly needed if we generate from scratch, but function requires it
            template = load_procurement_template(get_data_path(REQUIRED_FILES[2]))
            
            # Build Overrides
            # Map App Suppliers (A1, A2...) to Dashboard Suppliers (Supplier A, Supplier B...)
            # App: Part A -> A1, A2. Dashboard: Part A -> Supplier A, Supplier B.
            # App: Part B -> B1, B2. Dashboard: Part B -> Supplier A, Supplier B.
            
            supplier_map = {
                'A1': 'Supplier A', 'A2': 'Supplier B',
                'B1': 'Supplier A', 'B2': 'Supplier B'
            }
            
            # Orders DF: [Supplier, FN1...FN8]
            orders_df = st.session_state.get('purchasing_orders')
            
            overrides = {} 
            # Structure: {'Part A': {'Supplier A': {1: 500...}}}
            
            if orders_df is not None:
                for _, row in orders_df.iterrows():
                    app_sup = row['Supplier']
                    dash_sup = supplier_map.get(app_sup)
                    
                    if not dash_sup: continue
                    
                    part = "Part A" if "A" in app_sup else "Part B"
                    
                    if part not in overrides: overrides[part] = {}
                    if dash_sup not in overrides[part]: overrides[part][dash_sup] = {}
                    
                    for fn in FORTNIGHTS:
                        overrides[part][dash_sup][fn] = float(row.get(f'FN{fn}', 0))
            
            output = io.BytesIO()
            create_purchasing_dashboard(
                materials, costs, template,
                output_buffer=output,
                decision_overrides=overrides
            )
            return output
        except ImportError as e:
            st.error(f"Generate Error: {e}")
            return io.BytesIO()
        except Exception as e:
            st.error(f"Export Error: {e}")
            return io.BytesIO()
    
    @staticmethod
    def export_cpo_dashboard() -> io.BytesIO:
        """Export CPO People Dashboard (placeholder)."""
        ReportBridge.sync_cpo_to_shared()
        
        try:
            from generate_cpo_dashboard import (
                create_cpo_dashboard,
                load_workers_balance,
                load_sales_admin,
                load_labor_costs,
                load_absenteeism_data,
                get_data_path, REQUIRED_FILES
            )
            
            # Load Base Data
            workers = load_workers_balance(get_data_path(REQUIRED_FILES[0]))
            sales = load_sales_admin(get_data_path(REQUIRED_FILES[2]))
            labor = load_labor_costs(get_data_path(REQUIRED_FILES[1]))
            absenteeism = load_absenteeism_data(get_data_path(REQUIRED_FILES[0]))
            
            # Build Overrides from Session State (DataFrame)
            wf_df = st.session_state.get('cpo_workforce')
            
            decisions_override = {
                'workforce': {},
                'salary': {}
            }
            
            if wf_df is not None:
                for _, row in wf_df.iterrows():
                    zone = row['Zone']
                    decisions_override['workforce'][zone] = {
                        'required': row['Required_Workers'],
                        'turnover': row['Turnover_Rate'] / 100.0  # App uses %, Excel expects decimal
                    }
                    decisions_override['salary'][zone] = row['New_Salary']
                    
            output = io.BytesIO()
            create_cpo_dashboard(
                workers, sales, labor, absenteeism,
                output_buffer=output,
                decision_overrides=decisions_override
            )
            return output
        except ImportError as e:
            st.error(f"Generate Error: {e}")
            return io.BytesIO()
        except Exception as e:
            st.error(f"Export Error: {e}")
            return io.BytesIO()
    
        ReportBridge.sync_esg_to_shared()
        
        try:
            from generate_esg_dashboard import (
                create_esg_dashboard,
                load_esg_report,
                load_production_data,
                get_data_path, REQUIRED_FILES
            )
            
            # Load Base Data
            esg_data = load_esg_report(get_data_path(REQUIRED_FILES[0]))
            prod_data = load_production_data(get_data_path(REQUIRED_FILES[1]))
            
            # Build Overrides
            decisions_override = {
                "Solar PV Panels": st.session_state.get('esg_solar_panels', 0),
                "Trees Planted": st.session_state.get('esg_trees', 0),
                "Green Electricity": st.session_state.get('esg_green_electricity_pct', 0.0),
                "CO2 Credits": st.session_state.get('esg_co2_credits', 0)
            }
            
            output = io.BytesIO()
            create_esg_dashboard(
                esg_data, prod_data,
                output_buffer=output,
                decision_overrides=decisions_override
            )
            return output
        except ImportError as e:
            st.error(f"Generate Error: {e}")
            return io.BytesIO()
        except Exception as e:
            st.error(f"Export Error: {e}")
            return io.BytesIO()


def create_download_button(dashboard_name: str, tab_label: str):
    """
    Create a standardized download button for any dashboard.
    
    Args:
        dashboard_name: Internal name (CFO, CLO, CMO, Production, Purchasing, CPO, ESG)
        tab_label: Display label for the button
    """
    export_methods = {
        'CFO': ReportBridge.export_finance_dashboard,
        'CLO': ReportBridge.export_logistics_dashboard,
        'CMO': ReportBridge.export_cmo_dashboard,
        'Production': ReportBridge.export_production_dashboard,
        'Purchasing': ReportBridge.export_purchasing_dashboard,
        'CPO': ReportBridge.export_cpo_dashboard,
        'ESG': ReportBridge.export_esg_dashboard,
    }
    
    file_names = {
        'CFO': 'Finance_Dashboard_Live.xlsx',
        'CLO': 'Logistics_Dashboard_Live.xlsx',
        'CMO': 'Marketing_Dashboard_Live.xlsx',
        'Production': 'Production_Dashboard_Live.xlsx',
        'Purchasing': 'Purchasing_Dashboard_Live.xlsx',
        'CPO': 'People_Dashboard_Live.xlsx',
        'ESG': 'ESG_Dashboard_Live.xlsx',
    }
    
    if dashboard_name not in export_methods:
        st.error(f"Unknown dashboard: {dashboard_name}")
        return
    
    try:
        excel_buffer = export_methods[dashboard_name]()
        
        # Only show download if buffer has content
        if excel_buffer.getvalue():
            st.download_button(
                label="📥 Download Live",
                data=excel_buffer,
                file_name=file_names[dashboard_name],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        else:
            st.info("Export not yet available")
    except Exception as e:
        st.error(f"Export: {e}")
