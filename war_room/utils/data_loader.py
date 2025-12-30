"""
ExSim War Room - Data Loader
Parses Excel files using the exact formats from ExSim /data folders.
"""

import pandas as pd
import streamlit as st
from typing import Dict, Any, Optional


def parse_numeric(value) -> float:
    """Parse formatted number strings."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = str(value).replace('$', '').replace(',', '').replace('%', '').replace(' ', '').strip()
    if cleaned.startswith('(') and cleaned.endswith(')'):
        cleaned = '-' + cleaned[1:-1]
    try:
        return float(cleaned)
    except:
        return 0.0


def load_market_report(file) -> Dict[str, Any]:
    """Load market-report.xlsx - CMO input."""
    try:
        df = pd.read_excel(file, header=None)
        data = {
            'zones': {},
            'segments': ['High', 'Low'],
            'raw_df': df
        }
        
        current_zone = None
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            
            # Detect zone headers
            if first_val.upper() in ['CENTER', 'WEST', 'NORTH', 'EAST', 'SOUTH']:
                current_zone = first_val.capitalize()
                data['zones'][current_zone] = {'High': {}, 'Low': {}}
            
            # Parse segment data
            if current_zone and 'market share' in first_val.lower():
                for seg_idx, seg in enumerate(['High', 'Low']):
                    if seg_idx + 1 < len(row):
                        data['zones'][current_zone][seg]['market_share'] = parse_numeric(row.iloc[seg_idx + 1])
        
        return data
    except Exception as e:
        st.warning(f"Error loading market report: {e}")
        return {'zones': {}, 'segments': ['High', 'Low'], 'raw_df': None}


def load_workers_balance(file) -> Dict[str, Any]:
    """Load workers_balance_overtime.xlsx - CPO input."""
    try:
        df = pd.read_excel(file, header=None)
        data = {'zones': {}, 'raw_df': df}
        
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
            
            if 'workers assigned' in first_val:
                zones = ['Center', 'West', 'North', 'East', 'South']
                for z_idx, zone in enumerate(zones):
                    if z_idx + 1 < len(row):
                        if zone not in data['zones']:
                            data['zones'][zone] = {}
                        data['zones'][zone]['workers'] = parse_numeric(row.iloc[z_idx + 1])
            
            if 'salary' in first_val:
                zones = ['Center', 'West', 'North', 'East', 'South']
                for z_idx, zone in enumerate(zones):
                    if z_idx + 1 < len(row):
                        if zone not in data['zones']:
                            data['zones'][zone] = {}
                        data['zones'][zone]['salary'] = parse_numeric(row.iloc[z_idx + 1])
        
        return data
    except Exception as e:
        st.warning(f"Error loading workers balance: {e}")
        return {'zones': {}, 'raw_df': None}


def load_raw_materials(file) -> Dict[str, Any]:
    """Load raw_materials.xlsx - Purchasing input."""
    try:
        df = pd.read_excel(file, header=None)
        data = {'parts': {}, 'raw_df': df}
        
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            
            if 'part' in first_val.lower() and len(first_val) < 10:
                part_name = first_val
                data['parts'][part_name] = {
                    'stock': parse_numeric(row.iloc[1]) if len(row) > 1 else 0,
                    'cost': parse_numeric(row.iloc[2]) if len(row) > 2 else 0
                }
        
        return data
    except Exception as e:
        st.warning(f"Error loading raw materials: {e}")
        return {'parts': {}, 'raw_df': None}


def load_finished_goods(file) -> Dict[str, Any]:
    """Load finished_goods_inventory.xlsx - Logistics input."""
    try:
        df = pd.read_excel(file, header=None)
        data = {'zones': {}, 'raw_df': df}
        
        zones = ['Center', 'West', 'North', 'East', 'South']
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
            
            if 'inventory' in first_val or 'stock' in first_val:
                for z_idx, zone in enumerate(zones):
                    if z_idx + 1 < len(row):
                        if zone not in data['zones']:
                            data['zones'][zone] = {}
                        data['zones'][zone]['inventory'] = parse_numeric(row.iloc[z_idx + 1])
            
            if 'capacity' in first_val:
                for z_idx, zone in enumerate(zones):
                    if z_idx + 1 < len(row):
                        if zone not in data['zones']:
                            data['zones'][zone] = {}
                        data['zones'][zone]['capacity'] = parse_numeric(row.iloc[z_idx + 1])
        
        return data
    except Exception as e:
        st.warning(f"Error loading finished goods: {e}")
        return {'zones': {}, 'raw_df': None}


def load_balance_statements(file) -> Dict[str, Any]:
    """Load results_and_balance_statements.xlsx - CFO input."""
    try:
        df = pd.read_excel(file, header=None)
        data = {
            'net_sales': 0, 'cogs': 0, 'net_profit': 0,
            'total_assets': 0, 'total_liabilities': 0, 'equity': 0,
            'raw_df': df
        }
        
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
            val = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            
            if 'net sales' in first_val or 'revenue' in first_val:
                data['net_sales'] = val
            elif 'cost of goods sold' in first_val or 'cogs' in first_val:
                data['cogs'] = abs(val)
            elif 'net profit' in first_val or 'net income' in first_val:
                data['net_profit'] = val
            elif 'total assets' in first_val:
                data['total_assets'] = val
            elif 'total liabilities' in first_val:
                data['total_liabilities'] = abs(val)
            elif first_val == 'equity' or 'total equity' in first_val:
                data['equity'] = val
        
        return data
    except Exception as e:
        st.warning(f"Error loading balance statements: {e}")
        return {'net_sales': 0, 'cogs': 0, 'net_profit': 0, 'total_assets': 0, 'total_liabilities': 0, 'equity': 0, 'raw_df': None}


def load_esg_report(file) -> Dict[str, Any]:
    """Load esg_report.xlsx or ESG.xlsx - ESG input."""
    try:
        df = pd.read_excel(file, header=None)
        data = {'emissions': 0, 'energy': 0, 'tax_rate': 30, 'raw_df': df}
        
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
            val = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            
            if 'emission' in first_val and 'total' in first_val:
                data['emissions'] = val
            elif 'energy' in first_val:
                data['energy'] = val
            elif 'tax' in first_val and 'rate' in first_val:
                data['tax_rate'] = val
        
        return data
    except Exception as e:
        st.warning(f"Error loading ESG report: {e}")
        return {'emissions': 0, 'energy': 0, 'tax_rate': 30, 'raw_df': None}


def load_production_data(file) -> Dict[str, Any]:
    """Load production.xlsx - Production input."""
    try:
        df = pd.read_excel(file, header=None)
        data = {'zones': {}, 'machine_capacity': 0, 'raw_df': df}
        
        zones = ['Center', 'West', 'North', 'East', 'South']
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
            
            if 'machine' in first_val and 'capacity' in first_val:
                data['machine_capacity'] = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            
            if 'production' in first_val or 'output' in first_val:
                for z_idx, zone in enumerate(zones):
                    if z_idx + 1 < len(row):
                        if zone not in data['zones']:
                            data['zones'][zone] = {}
                        data['zones'][zone]['production'] = parse_numeric(row.iloc[z_idx + 1])
        
        return data
    except Exception as e:
        st.warning(f"Error loading production data: {e}")
        return {'zones': {}, 'machine_capacity': 0, 'raw_df': None}
