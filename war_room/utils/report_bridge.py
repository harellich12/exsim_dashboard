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
sys.path.insert(0, str(PROJECT_ROOT / "Purchasing Role"))
sys.path.insert(0, str(PROJECT_ROOT / "CPO Dashboard"))
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
    from config import get_data_path as config_get_data_path
    ZONES = COMMON.get('ZONES', ['Center', 'West', 'North', 'East', 'South'])
    FORTNIGHTS = COMMON.get('FORTNIGHTS', list(range(1, 9)))
except ImportError:
    ZONES = ['Center', 'West', 'North', 'East', 'South']
    FORTNIGHTS = list(range(1, 9))
    config_get_data_path = None


def safe_get_path(filename):
    """
    Safely get file path, returning None if not found.
    Uses config.get_data_path with required=False if available.
    """
    if config_get_data_path:
        try:
            return config_get_data_path(filename, required=False)
        except TypeError:
            # Old version without required parameter
            try:
                return config_get_data_path(filename)
            except FileNotFoundError:
                return None
    return None


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
            'innovation_costs': st.session_state.get('cmo_innovation_costs', 0),
            'payment_terms': st.session_state.get('cmo_payment_terms', {zone: 'D' for zone in ZONES})
        })
    
    @staticmethod
    def sync_production_to_shared():
        """Sync Production session state to shared_outputs.json"""
        if not export_dashboard_data:
            return
        
        export_dashboard_data('Production', {
            'production_plan': {zone: {'Target': st.session_state.get(f'prod_units_{zone}', 0)} for zone in ZONES},
            'capacity_utilization': {'mean': st.session_state.get('prod_capacity_util', 0.85)},
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
        
        # Build balance_data dict from session state (prefer live/uploaded data)
        # Try to use uploaded balance_data directly if available for missing keys
        balance_upload = st.session_state.get('balance_data', {})
        
        balance_data = {
            'net_sales': float(st.session_state.get('cfo_net_sales', balance_upload.get('net_sales', 0))),
            'cogs': float(st.session_state.get('cfo_cogs', balance_upload.get('cogs', 0))),
            'gross_income': float(st.session_state.get('cfo_gross_margin', balance_upload.get('gross_income', 0))),
            'net_profit': float(st.session_state.get('cfo_net_profit', balance_upload.get('net_profit', 0))),
            'total_assets': float(st.session_state.get('cfo_total_assets', balance_upload.get('total_assets', 0))),
            'total_liabilities': float(st.session_state.get('cfo_total_liabilities', balance_upload.get('total_liabilities', 0))),
            'equity': (float(st.session_state.get('cfo_total_assets', balance_upload.get('total_assets', 0))) 
                      - float(st.session_state.get('cfo_total_liabilities', balance_upload.get('total_liabilities', 0)))),
            'retained_earnings': float(st.session_state.get('cfo_retained_earnings', balance_upload.get('retained_earnings', 0))),
            'depreciation': float(balance_upload.get('depreciation', 0)),
            'gross_margin_pct': float(st.session_state.get('cfo_gross_margin_pct', balance_upload.get('gross_margin_pct', 0.4))),
            'net_margin_pct': float(st.session_state.get('cfo_net_margin_pct', balance_upload.get('net_margin_pct', 0.1)))
        }
        
        # Sales Admin Data (from 'sales_admin_expenses.xlsx')
        sa_upload = st.session_state.get('sales_admin_data', {})
        sa_data = {'total_sa_expenses': sa_upload.get('total_sa_expenses', 200000)}
        
        # AR/AP Data (from 'accounts_receivable_payable.xlsx')
        ar_ap_source = st.session_state.get('ar_ap_data', {})
        
        # Build schedule: Prefer cfo_cash_flow DF if available, else spread total
        ar_ap_data = {'receivables': [], 'payables': []}
        cash_flow_df = st.session_state.get('cfo_cash_flow')
        
        if cash_flow_df is not None and not cash_flow_df.empty:
             for _, row in cash_flow_df.iterrows():
                  ar_ap_data['receivables'].append(row.get('Receipts', 0))
                  ar_ap_data['payables'].append(row.get('Payments', 0))
        
        if not ar_ap_data['receivables']:
            rec_total = ar_ap_source.get('receivables', 0)
            rec_list = [0]*8
            if isinstance(rec_total, (int, float)):
                 rec_list = [rec_total/8]*8
            elif isinstance(rec_total, list):
                 rec_list = (rec_total + [0]*8)[:8]
            ar_ap_data['receivables'] = rec_list
            
            pay_total = ar_ap_source.get('payables', 0)
            pay_list = [0]*8
            if isinstance(pay_total, (int, float)):
                 pay_list = [pay_total/8]*8
            elif isinstance(pay_total, list):
                 pay_list = (pay_total + [0]*8)[:8]
            ar_ap_data['payables'] = pay_list
        else:
             # Ensure lists are exactly length 8 even if populated from DF
             ar_ap_data['receivables'] = (ar_ap_data['receivables'] + [0]*8)[:8]
             ar_ap_data['payables'] = (ar_ap_data['payables'] + [0]*8)[:8]
        
        # If cash_flow_df exists (user edited), overwrite with that logic?
        # Actually ReportBridge logic uses cash_flow_df to Populate the output if present,
        # but the `create_finance_dashboard` function separates Hard Data (from AR/AP) vs User Data (from df).
        # We should prioritize the Uploaded Data for the "Hard" slots if available.
        # But ReportBridge original logic tried to scrape from `cash_flow_df`. 
        # Better to rely on `ar_ap_upload` which is the source of truth for the Hard Data.
        
        starting_cash = (float(cash_data['final_cash']) 
                        - float(cash_data['tax_payments']) 
                        - float(st.session_state.get('cfo_dividend_payments', 0))
                        - float(st.session_state.get('cfo_asset_purchases', 0)))
        
        hard_data = {
            'depreciation': balance_data['depreciation'],
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
        
        # Fallback to upload data if session state not initialized
        fg_data = st.session_state.get('finished_goods_data', {})
        log_data = st.session_state.get('logistics_data', {})
        
        inventory_data = {}
        
        if inventory_df is not None and warehouses_df is not None:
             # Use Edited State
            for zone in ZONES:
                zone_inv_rows = inventory_df[inventory_df['Zone'] == zone]
                zone_wh_rows = warehouses_df[warehouses_df['Zone'] == zone]
                
                initial_inv = zone_inv_rows.iloc[0].get('Initial_Inv', 0) if not zone_inv_rows.empty else 0
                # Use Total_Capacity which considers buy/rent
                capacity = zone_wh_rows.iloc[0].get('Total_Capacity', 1000) if not zone_wh_rows.empty else 1000
                
                inventory_data[zone] = {'inventory': initial_inv, 'capacity': capacity}
        else:
             # Use Raw Upload Data
             zones_data = fg_data.get('zones', {}) if fg_data else {}
             for zone in ZONES:
                 z_info = zones_data.get(zone, {})
                 inventory_data[zone] = {
                     'inventory': z_info.get('inventory', 500), 
                     'capacity': z_info.get('capacity', 1000)
                 }
        
        template_data = {'df': None, 'exists': False}
        cost_data = {'total_shipping_cost': st.session_state.get('logistics_cost', 0)}
        
        # Pass intelligence if available
        # logistics_intelligence is benchmarks/penalties from 'logistics.xlsx'
        intelligence_data = None
        if log_data:
             intelligence_data = {
                 'benchmarks': log_data.get('benchmarks', {}),
                 'penalties': log_data.get('penalties', {})
             }
        
        output = io.BytesIO()
        create_logistics_dashboard(
            inventory_data=inventory_data,
            template_data=template_data,
            cost_data=cost_data,
            intelligence_data=intelligence_data,
            output_buffer=output
        )
        return output
    
    @staticmethod
    def export_cmo_dashboard() -> io.BytesIO:
        """Export CMO Marketing Dashboard."""
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
                get_data_path
            )

            # Define files locally for fallback
            REQUIRED_FILES = [
                'Marketing Decisions.xlsx',       # 0
                'Marketing Innovation Decisions.xlsx', # 1
                'market-report.xlsx',             # 2
                'sales_admin_expenses.xlsx',      # 3
                'finished_goods_inventory.xlsx'   # 4
            ]
            
            # 1. Market Data (Preferred from Session)
            market = st.session_state.get('market_data')
            if not market:
                market = load_market_report(safe_get_path(REQUIRED_FILES[2]))
            
            # 2. Innovation Features (Static? Or from Disk)
            # Not in bulk upload, so load from disk.
            innov_features = load_innovation_features(safe_get_path(REQUIRED_FILES[1]))
            
            # 3. Template (Static? Or from Disk)
            template = load_marketing_template(safe_get_path(REQUIRED_FILES[0]))
            
            # 4. Sales Admin (Preferred from Session)
            sales = st.session_state.get('sales_admin_data') # Keys map?
            # load_sales_admin_expenses returns dict {'last_sales': ...}
            # 'sales_admin_data' from bulk upload uses load_sales_admin_expenses too.
            if not sales:
                sales = load_sales_admin_expenses(safe_get_path(REQUIRED_FILES[3]))
            
            # 5. Inventory (Preferred from Session)
            inv = st.session_state.get('finished_goods_data')
            if not inv:
                inv = load_finished_goods_inventory(safe_get_path(REQUIRED_FILES[4]))
                
            # 6. Intelligence - BUILD FROM SESSION DATA to match UI
            # Instead of loading from disk (which may differ from Test Data),
            # build from session state to ensure consistency
            intel = {
                'economics': {
                    'TV_Cost_Spot': 3000,
                    'Radio_Cost_Spot': 300,
                    'Salary_Per_Person': 1500,
                    'Hiring_Cost': 1100
                },
                'pricing': {}
            }
            
            # Build competitor pricing from session market_data
            if market and market.get('zones'):
                for zone in ['Center', 'West', 'North', 'East', 'South']:
                    zone_data = market.get('zones', {}).get(zone, {})
                    intel['pricing'][zone] = zone_data.get('comp_avg_price', 68.0)
            
            # Build Overrides from UI state
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
            
            # 3. Zones - from cmo_strategy_inputs, but detect placeholder values
            strategy_df = st.session_state.get('cmo_strategy_inputs')
            placeholder_demands = [0, 1000, 900, 800, 850, 950]  # Default placeholders from init_cmo_state
            placeholder_prices = [0, 150, 145, 148, 152, 147]      # Default placeholders from init_cmo_state
            
            for zone in ['Center', 'West', 'North', 'East', 'South']:
                zone_market = market.get('zones', {}).get(zone, {}) if market else {}
                zone_sales = sales.get('by_zone', {}).get(zone, {}) if sales else {}
                
                # Get values from strategy_df if it exists
                df_target_demand = None
                df_price = None
                if strategy_df is not None:
                    zone_row = strategy_df[strategy_df['Zone'] == zone]
                    if not zone_row.empty:
                        df_target_demand = zone_row['Target_Demand'].values[0]
                        df_price = zone_row['Price'].values[0]
                
                # Use session data if strategy_df value is a placeholder
                if df_target_demand is not None and df_target_demand not in placeholder_demands:
                    target_demand = df_target_demand
                else:
                    # Use session data - this is the actual uploaded value
                    target_demand = zone_sales.get('units', 1000)
                
                if df_price is not None and df_price not in placeholder_prices:
                    price = df_price
                else:
                    # Use session data - actual market price
                    price = zone_market.get('my_price', 100)
                
                # Get other values from strategy_df if available, else defaults
                radio = 10
                headcount = 5
                if strategy_df is not None:
                    zone_row = strategy_df[strategy_df['Zone'] == zone]
                    if not zone_row.empty:
                        radio = zone_row['Radio_Spots'].values[0]
                        headcount = zone_row['Headcount'].values[0]
                
                overrides['zones'][zone] = {
                    'target_demand': target_demand,
                    'radio': radio,
                    'salespeople': headcount,
                    'price': price
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
        """Export Production Dashboard."""
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
            
            # Load Base Data (Prefer Session)
            # 1. Materials
            materials = st.session_state.get('materials_data')
            if not materials:
                materials = load_raw_materials_by_zone(safe_get_path(REQUIRED_FILES[0]))
                
            # 2. Finished Goods
            fg = st.session_state.get('finished_goods_data')
            if not fg:
                fg = load_finished_goods_by_zone(safe_get_path(REQUIRED_FILES[1]))
            
            # 3. Workers
            workers = st.session_state.get('workers_data')
            if not workers:
                workers = load_workers_by_zone(safe_get_path(REQUIRED_FILES[2]))
                
            # 4. Machines
            machines = st.session_state.get('machine_spaces_data')
            if not machines:
                machines = load_machines_by_zone(safe_get_path(REQUIRED_FILES[3]))
                
            # 5. Template (Static)
            template = load_production_template(safe_get_path(REQUIRED_FILES[4]))
            
            # Build Overrides
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
            
            # Load Base Data (Prefer Session)
            # 1. Materials
            materials = st.session_state.get('materials_data')
            if not materials:
                materials = load_raw_materials(safe_get_path(REQUIRED_FILES[0]))
            
            # 2. Production Costs
            # Check if production_data serves this purpose
            costs = st.session_state.get('production_data') 
            if not costs:
                costs = load_production_costs(safe_get_path(REQUIRED_FILES[1]))
                
            # 3. Template
            template = load_procurement_template(safe_get_path(REQUIRED_FILES[2]))
            
            # Build Overrides
            supplier_map = {
                'A1': 'Supplier A', 'A2': 'Supplier B',
                'B1': 'Supplier A', 'B2': 'Supplier B'
            }
            
            orders_df = st.session_state.get('purchasing_orders')
            
            overrides = {} 
            
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
        """Export CPO People Dashboard."""
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
            
            # Load Base Data (Prefer Session)
            # ================================================================
            # TRANSFORM: data_loader.py structures -> generator expectations
            # ================================================================
            
            # 1. Workers Balance
            # data_loader returns: {'zones': {zone: {'workers': N, 'salary': N}}, 'raw_df': df}
            # generator expects:   {zone: {'workers': N, 'absenteeism': N}}
            workers_source = st.session_state.get('workers_data')
            workers = {}
            absenteeism = 0.02
            
            if workers_source:
                zones_data = workers_source.get('zones', workers_source)
                ZONES = ['Center', 'West', 'North', 'East', 'South']
                for zone in ZONES:
                    zone_info = zones_data.get(zone, {})
                    workers[zone] = {
                        'workers': int(zone_info.get('workers', 0)),
                        'absenteeism': float(zone_info.get('absenteeism', 0.02))
                    }
                    # Extract first non-zero absenteeism for global rate
                    abs_val = zone_info.get('absenteeism')
                    if abs_val and float(abs_val) > 0 and absenteeism == 0.02:
                        absenteeism = float(abs_val)
            else:
                # Fallback to disk load if no session data
                workers = load_workers_balance(get_data_path(REQUIRED_FILES[0]))
                absenteeism = load_absenteeism_data(get_data_path(REQUIRED_FILES[0]))
            
            # 2. Sales Admin
            # data_loader returns: {'total_expenses': N, 'categories': {...}, 'by_zone': {...}, 'totals': {...}}
            # generator expects:   {'headcount': N, 'avg_salary': N, 'total_salary': N, 'hiring_cost': N}
            sales_source = st.session_state.get('sales_admin_data')
            
            if sales_source:
                # Transform the structure
                sales = {
                    'headcount': int(sales_source.get('headcount', 0)),
                    'avg_salary': float(sales_source.get('avg_salary', 750.0)),
                    'total_salary': float(sales_source.get('total_salary', 0)),
                    'hiring_cost': float(sales_source.get('hiring_cost', 1100.0)),
                }
                
                # Fallback: try to derive from categories if available
                categories = sales_source.get('categories', {})
                if sales['total_salary'] == 0:
                    for key, val in categories.items():
                        if 'salespeople' in key.lower() and 'salar' in key.lower():
                            sales['total_salary'] = float(val)
                            break
            else:
                # Fallback to disk load
                sales = load_sales_admin(get_data_path(REQUIRED_FILES[2]))
            
            # 3. Labor Costs
            # Not in bulk, load from disk
            labor = load_labor_costs(get_data_path(REQUIRED_FILES[1]))
            
            # Build Overrides from user inputs
            wf_df = st.session_state.get('cpo_workforce')
            
            decisions_override = {
                'workforce': {},
                'salary': {}
            }
            
            if wf_df is not None:
                for _, row in wf_df.iterrows():
                    try:
                        zone = row['Zone']
                        req = float(row['Required_Workers']) if pd.notna(row['Required_Workers']) else 0
                        turnover = float(row['Turnover_Rate']) if pd.notna(row['Turnover_Rate']) else 0
                        salary = float(row['New_Salary']) if pd.notna(row['New_Salary']) else 0
                        
                        decisions_override['workforce'][zone] = {
                            'required': req,
                            'turnover': turnover / 100.0
                        }
                        decisions_override['salary'][zone] = salary
                    except Exception:
                        continue
            
            # Final Validation: Ensure no None values
            if workers is None:
                workers = {}
            if sales is None:
                sales = {'hiring_cost': 1100.0, 'headcount': 0, 'avg_salary': 750, 'total_salary': 0}
            if labor is None:
                labor = {'total_labor': 0}
            if absenteeism is None:
                absenteeism = 0.02
            
            # Ensure sales dict has required keys with numeric values
            if not isinstance(sales.get('hiring_cost'), (int, float)):
                sales['hiring_cost'] = 1100.0
            if not isinstance(sales.get('headcount'), (int, float)):
                sales['headcount'] = 0
            if not isinstance(sales.get('avg_salary'), (int, float)):
                sales['avg_salary'] = 750.0
                    
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
    
    @staticmethod
    def export_esg_dashboard() -> io.BytesIO:
        """Export ESG Dashboard using live Streamlit data."""
        ReportBridge.sync_esg_to_shared()
        
        try:
            from generate_esg_dashboard import (
                create_esg_dashboard,
                load_esg_report,
                load_production_data,
                get_data_path, REQUIRED_FILES
            )
            
            # Load Base Data (Prefer Session)
            esg_data = st.session_state.get('esg_data')
            if not esg_data:
                esg_data = load_esg_report(get_data_path(REQUIRED_FILES[0]))
            
            # Patch for missing key if old data loaded
            if 'energy_consumption' not in esg_data and 'energy' in esg_data:
                esg_data['energy_consumption'] = esg_data['energy']
                
            prod_data = st.session_state.get('production_data')
            if not prod_data:
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
