"""
Report Bridge - Connects Streamlit live data to Excel dashboard generators.

This module bridges the session state data from the Streamlit app to the 
openpyxl-based dashboard generators, enabling "Live Report Export" functionality.
"""

import io
import streamlit as st
import sys
from pathlib import Path

# Add dashboard directories to path for imports
PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "CFO Dashboard"))
sys.path.insert(0, str(PROJECT_ROOT / "CLO Dashboard"))


class ReportBridge:
    """Bridge between Streamlit session state and Excel dashboard generators."""
    
    @staticmethod
    def export_finance_dashboard() -> io.BytesIO:
        """
        Export CFO Finance Dashboard using live Streamlit data.
        
        Gathers data from session state variables set by tab_cfo.py:
        - cfo_cash_flow: DataFrame with operational cash flow
        - cfo_financing: DataFrame with financing decisions
        - cfo_mortgages: DataFrame with mortgage data
        - Various cfo_* scalar values (net_sales, total_assets, etc.)
        
        Returns:
            BytesIO buffer containing the Excel file
        """
        from generate_finance_dashboard_final import create_finance_dashboard
        
        # Extract operatinonal cash flow data
        cash_flow_df = st.session_state.get('cfo_cash_flow')
        financing_df = st.session_state.get('cfo_financing')
        
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
        
        # S&A data
        sa_data = {'total_sa_expenses': 200000}  # Default or from session
        
        # Build AR/AP data from cash_flow grid rows (rows 3=receivables, 4=payables)
        ar_ap_data = {'receivables': [0]*8, 'payables': [0]*8}
        
        if cash_flow_df is not None:
            for fn in range(1, 9):
                fn_col = f'FN{fn}'
                if fn_col in cash_flow_df.columns:
                    ar_ap_data['receivables'][fn-1] = float(cash_flow_df.at[3, fn_col])
                    ar_ap_data['payables'][fn-1] = abs(float(cash_flow_df.at[4, fn_col]))
        
        # Calculate starting cash
        starting_cash = (cash_data['final_cash'] 
                        - cash_data['tax_payments'] 
                        - st.session_state.get('cfo_dividend_payments', 0)
                        - st.session_state.get('cfo_asset_purchases', 0))
        
        # Build hard_data dict
        hard_data = {
            'depreciation': 0,
            'starting_cash': starting_cash,
            'schedule': {fn: {'receivables': ar_ap_data['receivables'][fn-1], 
                              'payables': ar_ap_data['payables'][fn-1]} for fn in range(1, 9)},
            'retained_earnings': balance_data['retained_earnings']
        }
        
        # Generate to BytesIO buffer
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
        """
        Export CLO Logistics Dashboard using live Streamlit data.
        
        Gathers data from session state variables set by tab_logistics.py:
        - logistics_shipments: DataFrame with shipment data
        - logistics_inventory: DataFrame with inventory by zone
        - logistics_warehouses: DataFrame with warehouse configuration
        
        Returns:
            BytesIO buffer containing the Excel file
        """
        from generate_logistics_dashboard import create_logistics_dashboard
        
        # Extract live data from session state
        shipments_df = st.session_state.get('logistics_shipments')
        inventory_df = st.session_state.get('logistics_inventory')
        warehouses_df = st.session_state.get('logistics_warehouses')
        
        # Build inventory_data dict by zone
        zones = ['Center', 'West', 'North', 'East', 'South']
        inventory_data = {}
        
        if inventory_df is not None and warehouses_df is not None:
            for zone in zones:
                # Get inventory row for this zone
                zone_inv_rows = inventory_df[inventory_df['Zone'] == zone]
                zone_wh_rows = warehouses_df[warehouses_df['Zone'] == zone]
                
                initial_inv = 0
                capacity = 1000  # Default capacity
                
                if not zone_inv_rows.empty:
                    initial_inv = zone_inv_rows.iloc[0].get('Initial_Inv', 0)
                
                if not zone_wh_rows.empty:
                    capacity = zone_wh_rows.iloc[0].get('Total_Capacity', 1000)
                
                inventory_data[zone] = {
                    'inventory': initial_inv,
                    'capacity': capacity
                }
        else:
            # Default values if no data loaded
            for zone in zones:
                inventory_data[zone] = {'inventory': 500, 'capacity': 1000}
        
        # Template data - include shipments if available
        template_data = {'df': None, 'exists': False}
        
        # Cost data defaults
        cost_data = {'total_shipping_cost': 0}
        
        # Generate to BytesIO buffer
        output = io.BytesIO()
        create_logistics_dashboard(
            inventory_data=inventory_data,
            template_data=template_data,
            cost_data=cost_data,
            intelligence_data=None,
            output_buffer=output
        )
        return output
