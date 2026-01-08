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


def _load_market_report_xml(content: bytes, data: Dict, zones: list, segments: list, my_company: str) -> Dict[str, Any]:
    """
    Parse SpreadsheetML XML format for market-report.xls files.
    Extracts data for the specified company (default A3) and populates the CMO data structure.
    """
    import xml.etree.ElementTree as ET
    import io
    
    ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}
    tree = ET.parse(io.BytesIO(content))
    root = tree.getroot()
    
    sheet = root.find('.//ss:Worksheet', ns)
    if sheet is None:
        return data
    table = sheet.find('ss:Table', ns)
    if table is None:
        return data
    rows = table.findall('ss:Row', ns)
    
    current_section = None
    COMPANIES = ['A1', 'A2', 'A3', 'A4']
    OTHER_COMPANIES = [c for c in COMPANIES if c != my_company]
    
    # Temporary storage
    store = {
        'market_share_region': {},
        'market_share_segment': {},
        'price': {},
        'awareness': {},
        'attractiveness': {},
        'salesforce': {}
    }
    
    def clean_num(val):
        try:
            if isinstance(val, str):
                val = val.strip()
                if not val:
                    return 0.0
            return float(val)
        except:
            return 0.0
    
    def parse_zone_table(row_data, key):
        if not row_data:
            return
        if row_data[0].strip() in zones:
            zone = row_data[0].strip()
            if len(row_data) >= 5:
                for i, comp in enumerate(COMPANIES):
                    store[key].setdefault(comp, {})[zone] = clean_num(row_data[i + 1])
    
    def parse_segment_table(row_data, key):
        if not row_data or len(row_data) < 3:
            return
        zone_attr = f'_current_{key}_zone'
        zone_candidate = row_data[0].strip()
        segment_candidate = row_data[1].strip() if len(row_data) > 1 else ""
        
        if zone_candidate in zones:
            _load_market_report_xml.__dict__[zone_attr] = zone_candidate
            zone = zone_candidate
        elif zone_attr in _load_market_report_xml.__dict__:
            zone = _load_market_report_xml.__dict__[zone_attr]
        else:
            return

        if segment_candidate in segments:
            if len(row_data) >= 6:
                for i, comp in enumerate(COMPANIES):
                    store[key].setdefault(comp, {}).setdefault(zone, {})[segment_candidate] = clean_num(row_data[i + 2])
    
    for row in rows:
        cells = row.findall('ss:Cell', ns)
        row_data = []
        for cell in cells:
            cell_data = cell.find('ss:Data', ns)
            text = cell_data.text if cell_data is not None and cell_data.text is not None else ""
            row_data.append(text)
        
        full_row_text = " ".join([str(x) for x in row_data]).strip()
        
        # Section detection
        if "Market Share Per Region (%)" in full_row_text and "Segment" not in full_row_text:
            current_section = "market_share_region"
            continue
        elif "Market Share Per Region Per Segment (%)" in full_row_text:
            current_section = "market_share_segment"
            continue
        elif row_data and row_data[0].strip().lower() == 'price' or (row_data and 'Price' in row_data[0]):
            current_section = "price"
            continue
        elif "Product Awareness Percentage Per Segment" in full_row_text:
            current_section = "awareness"
            continue
        elif "Product attractiveness (Perceived)" in full_row_text:
            current_section = "attractiveness"
            continue
        elif "Evaluation of the Promotional Impact of Salesforce" in full_row_text:
            current_section = "salesforce"
            continue
        
        # Data extraction
        if current_section == "market_share_region":
            parse_zone_table(row_data, 'market_share_region')
        elif current_section == "market_share_segment":
            parse_segment_table(row_data, 'market_share_segment')
        elif current_section == "price":
            parse_zone_table(row_data, 'price')
        elif current_section == "awareness":
            parse_segment_table(row_data, 'awareness')
        elif current_section == "attractiveness":
            parse_segment_table(row_data, 'attractiveness')
        elif current_section == "salesforce":
            parse_zone_table(row_data, 'salesforce')
    
    # Convert store to CMO data structure
    for zone in zones:
        # Zone-level data (from region tables)
        my_share = store['market_share_region'].get(my_company, {}).get(zone, 0)
        my_price = store['price'].get(my_company, {}).get(zone, 0)
        
        # Competitor averages
        comp_prices = [store['price'].get(c, {}).get(zone, 0) for c in OTHER_COMPANIES]
        comp_prices = [p for p in comp_prices if p > 0]
        comp_avg_price = sum(comp_prices) / len(comp_prices) if comp_prices else 0
        
        data['zones'][zone]['my_market_share'] = my_share
        data['zones'][zone]['my_price'] = my_price
        data['zones'][zone]['comp_avg_price'] = comp_avg_price
        
        # Segment-level data
        for segment in segments:
            seg_share = store['market_share_segment'].get(my_company, {}).get(zone, {}).get(segment, 0)
            seg_awareness = store['awareness'].get(my_company, {}).get(zone, {}).get(segment, 0)
            seg_attract = store['attractiveness'].get(my_company, {}).get(zone, {}).get(segment, 0)
            
            # Competitor awareness average
            comp_awareness = [store['awareness'].get(c, {}).get(zone, {}).get(segment, 0) for c in OTHER_COMPANIES]
            comp_awareness = [a for a in comp_awareness if a > 0]
            comp_avg_aware = sum(comp_awareness) / len(comp_awareness) if comp_awareness else 0
            
            data['by_segment'][segment][zone]['my_market_share'] = seg_share
            data['by_segment'][segment][zone]['my_awareness'] = seg_awareness
            data['by_segment'][segment][zone]['my_attractiveness'] = seg_attract
            data['by_segment'][segment][zone]['my_price'] = my_price
            data['by_segment'][segment][zone]['comp_avg_awareness'] = comp_avg_aware
            data['by_segment'][segment][zone]['comp_avg_price'] = comp_avg_price
            
            # Update zone-level awareness (use first segment found)
            if data['zones'][zone]['my_awareness'] == 0:
                data['zones'][zone]['my_awareness'] = seg_awareness
            if data['zones'][zone]['my_attractiveness'] == 0:
                data['zones'][zone]['my_attractiveness'] = seg_attract
            if data['zones'][zone]['comp_avg_awareness'] == 0:
                data['zones'][zone]['comp_avg_awareness'] = comp_avg_aware
    
    return data

def load_market_report(file) -> Dict[str, Any]:
    """
    Load market-report.xlsx or market-report.xls (SpreadsheetML XML) - CMO input.
    Parses segment-level data including market share, awareness, price, attractiveness.
    Must extract: my_market_share, my_awareness, my_price, my_attractiveness, 
                  comp_avg_awareness, comp_avg_price for each zone.
    """
    ZONES = ['Center', 'West', 'North', 'East', 'South']
    SEGMENTS = ['High', 'Low']
    MY_COMPANY = 'Company 3'  # Default company identifier
    MY_COMPANY_ID = 'A3'  # For XML format
    
    # Initialize data structure matching CMO generator
    data = {
        'by_segment': {seg: {zone: {
            'my_market_share': 0,
            'my_awareness': 0,
            'my_attractiveness': 0,
            'my_price': 0,
            'comp_avg_awareness': 0,
            'comp_avg_price': 0
        } for zone in ZONES} for seg in SEGMENTS},
        'zones': {zone: {
            'my_price': 0,
            'comp_avg_price': 0,
            'my_awareness': 0,
            'my_attractiveness': 0,
            'my_market_share': 0,
            'comp_avg_awareness': 0
        } for zone in ZONES},
        'segments': SEGMENTS,
        'raw_df': None
    }
    
    try:
        # Read file content to detect format
        if hasattr(file, 'read'):
            content = file.read()
            if hasattr(file, 'seek'):
                file.seek(0)
        else:
            with open(file, 'rb') as f:
                content = f.read()
        
        # Check if it's SpreadsheetML XML format
        is_xml = content[:100].decode('utf-8', errors='ignore').strip().startswith('<?xml')
        
        if is_xml:
            # Use custom XML parser for SpreadsheetML format
            return _load_market_report_xml(content, data, ZONES, SEGMENTS, MY_COMPANY_ID)
        else:
            # Standard Excel format - use pandas
            if hasattr(file, 'seek'):
                file.seek(0)
            df = pd.read_excel(file, header=None)
            data['raw_df'] = df
        
        current_section = None
        my_company_col = None
        comp_cols = []
        last_zone = None
        
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            second_val = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ''
            
            # Detect section headers
            if 'market share' in first_val.lower() and 'segment' in first_val.lower():
                current_section = 'segment_share'
                my_company_col = None
            elif 'market share' in first_val.lower() and 'region' in first_val.lower():
                current_section = 'region_share'
                my_company_col = None
            elif 'awareness' in first_val.lower() and 'segment' in first_val.lower():
                current_section = 'segment_awareness'
                my_company_col = None
            elif 'attractiveness' in first_val.lower():
                current_section = 'attractiveness'
                my_company_col = None
            elif 'price' in first_val.lower() and 'zone' not in first_val.lower():
                current_section = 'price'
                my_company_col = None
            
            # Detect column headers with company names
            # Handle both formats: "Company 3", "A3 - CompanyName", or just "A3"
            if first_val.lower() == 'zone' or first_val.lower() == 'region':
                comp_cols = []
                for col_idx in range(len(row)):
                    col_val = str(row.iloc[col_idx]).strip() if pd.notna(row.iloc[col_idx]) else ''
                    # Match A3 or "A3 - ..." for my company
                    if col_val.startswith(MY_COMPANY_ID) or MY_COMPANY in col_val:
                        my_company_col = col_idx
                    # Match other company patterns: A1, A2, A4 or "Company N"
                    elif any(col_val.startswith(c) for c in ['A1', 'A2', 'A4']) or \
                         ('Company' in col_val and MY_COMPANY not in col_val):
                        comp_cols.append(col_idx)
            
            # Parse zone data rows
            current_zone = None
            current_segment = None
            
            for zone in ZONES:
                if first_val.lower() == zone.lower():
                    current_zone = zone
                    last_zone = zone
                    break
            
            # Check for segment in second column
            if second_val.lower() in ['high', 'low']:
                current_segment = second_val.capitalize()
            
            # Also check continuation rows
            if first_val == '' and second_val.lower() in ['high', 'low']:
                current_segment = second_val.capitalize()
                current_zone = last_zone  # Use last known zone
            
            # Parse data if we have zone and company column
            if current_zone and my_company_col:
                try:
                    my_val = parse_numeric(row.iloc[my_company_col])
                    
                    # Get competitor average
                    comp_vals = [parse_numeric(row.iloc[c]) for c in comp_cols if c < len(row)]
                    comp_vals = [v for v in comp_vals if v > 0]
                    comp_avg = sum(comp_vals) / len(comp_vals) if comp_vals else 0
                    
                    # Store data based on section
                    if current_section == 'region_share' and my_val > 0:
                        data['zones'][current_zone]['my_market_share'] = my_val
                        
                    elif current_section == 'segment_share' and current_segment and my_val > 0:
                        data['by_segment'][current_segment][current_zone]['my_market_share'] = my_val
                        # Also update zones aggregate
                        if data['zones'][current_zone]['my_market_share'] == 0:
                            data['zones'][current_zone]['my_market_share'] = my_val
                        
                    elif current_section == 'price' and my_val > 0:
                        data['zones'][current_zone]['my_price'] = my_val
                        if comp_avg > 0:
                            data['zones'][current_zone]['comp_avg_price'] = comp_avg
                        for seg in SEGMENTS:
                            data['by_segment'][seg][current_zone]['my_price'] = my_val
                            if comp_avg > 0:
                                data['by_segment'][seg][current_zone]['comp_avg_price'] = comp_avg
                                
                    elif current_section == 'segment_awareness' and current_segment and my_val > 0:
                        data['zones'][current_zone]['my_awareness'] = my_val
                        if comp_avg > 0:
                            data['zones'][current_zone]['comp_avg_awareness'] = comp_avg
                        data['by_segment'][current_segment][current_zone]['my_awareness'] = my_val
                        if comp_avg > 0:
                            data['by_segment'][current_segment][current_zone]['comp_avg_awareness'] = comp_avg
                            
                    elif current_section == 'attractiveness' and my_val > 0:
                        data['zones'][current_zone]['my_attractiveness'] = my_val
                        if current_segment:
                            data['by_segment'][current_segment][current_zone]['my_attractiveness'] = my_val
                        
                except Exception:
                    continue
        
        return data
        
    except Exception as e:
        st.warning(f"Error loading market report: {e}")
        return {
            'zones': {zone: {
                'my_price': 0, 'comp_avg_price': 0, 'my_awareness': 0,
                'my_attractiveness': 0, 'my_market_share': 0, 'comp_avg_awareness': 0
            } for zone in ['Center', 'West', 'North', 'East', 'South']},
            'by_segment': {seg: {zone: {
                'my_market_share': 0, 'my_awareness': 0, 'my_attractiveness': 0,
                'my_price': 0, 'comp_avg_awareness': 0, 'comp_avg_price': 0
            } for zone in ['Center', 'West', 'North', 'East', 'South']} for seg in ['High', 'Low']},
            'segments': ['High', 'Low'],
            'raw_df': None
        }


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
    """
    Load finished_goods_inventory.xlsx - Logistics/CMO input.
    Detects stockouts: if final inventory <= 0 for any zone, is_stockout = True.
    """
    ZONES = ['Center', 'West', 'North', 'East', 'South']
    try:
        df = pd.read_excel(file, header=None)
        data = {
            'zones': {zone: {'inventory': 0, 'capacity': 0, 'final': 0} for zone in ZONES},
            'is_stockout': False,
            'total_final_inventory': 0,
            'raw_df': df
        }
        
        current_zone_idx = 0
        
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
            
            # Parse capacity header
            if 'capacity' in first_val:
                for z_idx, zone in enumerate(ZONES):
                    if z_idx + 1 < len(row):
                        data['zones'][zone]['capacity'] = parse_numeric(row.iloc[z_idx + 1])
            
            # Parse inventory/stock rows
            if 'inventory' in first_val or 'stock' in first_val:
                for z_idx, zone in enumerate(ZONES):
                    if z_idx + 1 < len(row):
                        data['zones'][zone]['inventory'] = parse_numeric(row.iloc[z_idx + 1])
            
            # Parse final inventory row (key for stockout detection)
            if 'final' in first_val and 'inventory' in first_val:
                # Get fortnight 8 value if available, otherwise last column
                final_col = min(9, len(row) - 1)  # Column 9 (0-indexed) = FN8
                for z_idx, zone in enumerate(ZONES):
                    col_idx = z_idx + 1
                    if col_idx < len(row):
                        final_val = parse_numeric(row.iloc[col_idx])
                        data['zones'][zone]['final'] = final_val
                        data['total_final_inventory'] += final_val
                        
                        # Stockout if any zone has final inventory <= 0
                        if final_val <= 0:
                            data['is_stockout'] = True
        
        return data
    except Exception as e:
        st.warning(f"Error loading finished goods: {e}")
        return {
            'zones': {zone: {'inventory': 0, 'capacity': 0, 'final': 0} for zone in ['Center', 'West', 'North', 'East', 'South']},
            'is_stockout': False,
            'total_final_inventory': 0,
            'raw_df': None
        }


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
        data = {'emissions': 0, 'energy': 0, 'energy_consumption': 0, 'tax_rate': 30, 'raw_df': df}
        
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
            val = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            
            if 'emission' in first_val and 'total' in first_val:
                data['emissions'] = val
            elif 'energy' in first_val:
                data['energy'] = val
                data['energy_consumption'] = val
            elif 'tax' in first_val and 'rate' in first_val:
                data['tax_rate'] = val
        
        return data
    except Exception as e:
        st.warning(f"Error loading ESG report: {e}")
        return {'emissions': 0, 'energy': 0, 'energy_consumption': 0, 'tax_rate': 30, 'raw_df': None}


def load_production_data(file) -> Dict[str, Any]:
    """
    Load production.xlsx - Production input.
    Robust parser that greedily looks for:
    - Machine counts ("Machines", "Machine Capacity")
    - Module counts ("Modules", "Slots")
    - Historic Production ("Production", "Output")
    """
    try:
        df = pd.read_excel(file, header=None)
        
        # Initialize structure
        zones = ['Center', 'West', 'North', 'East', 'South']
        data = {
            'zones': {z: {'machines': 0, 'modules': 0, 'capacity': 0, 'production': 0} for z in zones},
            'machine_capacity': 0,
            'raw_df': df
        }
        
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
            
            # Helper to extract zone values from a row
            def extract_zone_values(row_data, key):
                found = False
                for z_idx, zone in enumerate(zones):
                    # Data usually starts at column 1 (index 1) or 2 depending on format
                    # matches legacy format: Label | Center | West ...
                    # or: Label | ... | Center ...
                    
                    # Try simple offset first
                    if z_idx + 1 < len(row_data):
                        val = parse_numeric(row_data.iloc[z_idx + 1])
                        if val > 0:
                            data['zones'][zone][key] = val
                            found = True
                return found

            # 1. Machines Count
            if 'machines' in first_val and 'capacity' not in first_val:
                extract_zone_values(row, 'machines')

            # 2. Modules / Slots
            elif 'module' in first_val or 'slot' in first_val or 'space' in first_val:
                extract_zone_values(row, 'modules')

            # 3. Machine Capacity (Units)
            elif 'machine' in first_val and 'capacity' in first_val:
                extract_zone_values(row, 'capacity')

            # 4. Historic Production
            elif 'production' in first_val or 'output' in first_val:
                extract_zone_values(row, 'production')

        # Backfill: If we hav Capacity but no Machines, derive Machines (Capacity / 100)
        for zone in zones:
            z_data = data['zones'][zone]
            if z_data['machines'] == 0 and z_data['capacity'] > 0:
                z_data['machines'] = int(z_data['capacity'] / 100)
        
        return data
    except Exception as e:
        st.warning(f"Error loading production data: {e}")
        return {'zones': {}, 'machine_capacity': 0, 'raw_df': None}


def load_sales_admin_expenses(file) -> Dict[str, Any]:
    """Load sales_admin_expenses.xlsx - CFO input for S&A expenses."""
    ZONES = ['Center', 'West', 'North', 'East', 'South']
    try:
        df = pd.read_excel(file, header=None)
        data = {
            'total_expenses': 0, 
            'categories': {}, 
            'raw_df': df,
            # Match CMO structure requirements:
            'by_zone': {zone: {'units': 0, 'price': 0} for zone in ZONES},
            'totals': {'units': 0, 'tv_spend': 0, 'radio_spend': 0, 'salespeople_cost': 0}
        }
        
        in_sales_section = False
        in_expense_section = False
        
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
            val = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            
            # Detect sections
            if 'sales' in first_val and 'expense' not in first_val and 'admin' not in first_val:
                in_sales_section = True
                in_expense_section = False
            elif 'expense' in first_val or 'admin' in first_val:
                in_sales_section = False
                in_expense_section = True
                
            if in_sales_section:
                for zone in ZONES:
                    if first_val == zone.lower():
                        # Format: Region, Brand, Units, Local Price, ...
                        units = parse_numeric(row.iloc[2]) if len(row) > 2 else 0
                        price = parse_numeric(row.iloc[3]) if len(row) > 3 else 0
                        if units > 0:
                            data['by_zone'][zone]['units'] = units
                            data['totals']['units'] += units
                        if price > 0:
                            data['by_zone'][zone]['price'] = price
                            
            if 'total' in first_val and ('expense' in first_val or 's&a' in first_val):
                data['total_expenses'] = val
            elif in_expense_section and first_val and val != 0:
                data['categories'][first_val] = val
                
                # Update totals for CMO
                if 'tv' in first_val and 'advert' in first_val:
                    data['totals']['tv_spend'] = val
                elif 'radio' in first_val and 'advert' in first_val:
                    data['totals']['radio_spend'] = val
                elif 'salespeople' in first_val and 'salar' in first_val:
                    data['totals']['salespeople_cost'] = val
        
        return data
    except Exception as e:
        st.warning(f"Error loading sales admin expenses: {e}")
        return {
            'total_expenses': 0, 'categories': {}, 'raw_df': None,
            'by_zone': {zone: {'units': 0, 'price': 0} for zone in ZONES},
            'totals': {'units': 0, 'tv_spend': 0, 'radio_spend': 0, 'salespeople_cost': 0}
        }


def load_subperiod_cash_flow(file) -> Dict[str, Any]:
    """Load subperiod_cash_flow.xlsx - CFO input for cash flow by fortnight."""
    try:
        df = pd.read_excel(file, header=None)
        data = {'fortnights': {}, 'total_inflow': 0, 'total_outflow': 0, 'raw_df': df}
        
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
            
            if 'inflow' in first_val or 'receipts' in first_val:
                for fn in range(1, 9):
                    if fn < len(row):
                        if f'fn{fn}' not in data['fortnights']:
                            data['fortnights'][f'fn{fn}'] = {}
                        data['fortnights'][f'fn{fn}']['inflow'] = parse_numeric(row.iloc[fn])
            
            if 'outflow' in first_val or 'payments' in first_val:
                for fn in range(1, 9):
                    if fn < len(row):
                        if f'fn{fn}' not in data['fortnights']:
                            data['fortnights'][f'fn{fn}'] = {}
                        data['fortnights'][f'fn{fn}']['outflow'] = parse_numeric(row.iloc[fn])
        
        return data
    except Exception as e:
        st.warning(f"Error loading subperiod cash flow: {e}")
        return {'fortnights': {}, 'total_inflow': 0, 'total_outflow': 0, 'raw_df': None}


def load_accounts_receivable_payable(file) -> Dict[str, Any]:
    """Load accounts_receivable_payable.xlsx - CFO input for AR/AP."""
    try:
        df = pd.read_excel(file, header=None)
        data = {'receivables': 0, 'payables': 0, 'net_position': 0, 'raw_df': df}
        
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
            val = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            
            if 'receivable' in first_val and 'total' in first_val:
                data['receivables'] = val
            elif 'payable' in first_val and 'total' in first_val:
                data['payables'] = abs(val)
        
        data['net_position'] = data['receivables'] - data['payables']
        return data
    except Exception as e:
        st.warning(f"Error loading accounts receivable/payable: {e}")
        return {'receivables': 0, 'payables': 0, 'net_position': 0, 'raw_df': None}


def load_financial_statements_summary(file) -> Dict[str, Any]:
    """Load financial_statements_summary.xlsx - CFO input for summary financials."""
    try:
        df = pd.read_excel(file, header=None)
        data = {'revenue': 0, 'gross_profit': 0, 'operating_income': 0, 'net_income': 0, 'raw_df': df}
        
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
            val = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            
            if 'revenue' in first_val or 'net sales' in first_val:
                data['revenue'] = val
            elif 'gross profit' in first_val:
                data['gross_profit'] = val
            elif 'operating' in first_val and ('income' in first_val or 'profit' in first_val):
                data['operating_income'] = val
            elif 'net' in first_val and ('income' in first_val or 'profit' in first_val):
                data['net_income'] = val
        
        return data
    except Exception as e:
        st.warning(f"Error loading financial statements summary: {e}")
        return {'revenue': 0, 'gross_profit': 0, 'operating_income': 0, 'net_income': 0, 'raw_df': None}


def load_initial_cash_flow(file) -> Dict[str, Any]:
    """Load initial_cash_flow.xlsx - CFO input for opening cash positions."""
    try:
        df = pd.read_excel(file, header=None)
        data = {'opening_cash': 0, 'available_credit': 0, 'net_liquidity': 0, 'raw_df': df}
        
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
            val = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            
            if 'opening' in first_val and 'cash' in first_val:
                data['opening_cash'] = val
            elif 'available' in first_val and 'credit' in first_val:
                data['available_credit'] = val
            elif 'liquidity' in first_val or 'total cash' in first_val:
                data['net_liquidity'] = val
        
        return data
    except Exception as e:
        st.warning(f"Error loading initial cash flow: {e}")
        return {'opening_cash': 0, 'available_credit': 0, 'net_liquidity': 0, 'raw_df': None}


def load_logistics_data(file) -> Dict[str, Any]:
    """
    Load logistics.xlsx - Logistics input for shipping costs and warehouse penalties.
    Extracts:
    1. 'benchmarks': Route costs (Transport Costs table)
    2. 'penalties': Warehouse rent costs (Incoming/Outcoming table)
    3. 'shipping_costs': Shipping costs per zone (legacy)
    """
    try:
        df = pd.read_excel(file, header=None)
        data = {
            'zones': {}, 
            'shipping_costs': {}, 
            'benchmarks': {},  # NEW: Route costs (e.g. Center-North Train)
            'penalties': {},   # NEW: Zone warehouse costs
            'raw_df': df
        }
        
        zones = ['Center', 'West', 'North', 'East', 'South']
        
        # 1. Parse Transportation Costs (Benchmarks)
        in_transport_section = False
        for idx, row in df.iterrows():
            label = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            
            if "Transportation Costs" in label:
                in_transport_section = True
                continue
            
            if in_transport_section:
                if "Type" in label or "Subtotal" in label or label == "Total":
                    if label == "Total": in_transport_section = False
                    continue
                
                # Route row (e.g. "Train Center-North")
                if label:
                    route = label
                    units = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
                    total = parse_numeric(row.iloc[3]) if len(row) > 3 else 0
                    
                    if units > 0:
                        cost_per_unit = total / units
                        # Aggregate average if duplicate routes appear
                        if route in data['benchmarks']:
                            data['benchmarks'][route] = (data['benchmarks'][route] + cost_per_unit) / 2
                        else:
                            data['benchmarks'][route] = cost_per_unit
        
        # 2. Parse Warehouse Penalties & Shipping Costs
        for idx, row in df.iterrows():
            label = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            
            # Shipping Costs (Legacy / Direct map)
            if 'shipping' in label.lower() or 'transport' in label.lower():
                for z_idx, zone in enumerate(zones):
                    if z_idx + 1 < len(row):
                        data['shipping_costs'][zone] = parse_numeric(row.iloc[z_idx + 1])

            # Warehouse Penalties from "Incoming and Outcoming by Zone" table
            if "Incoming and Outcoming by Zone" in label:
                # Data starts 2 rows down
                for offset in range(2, 10):
                    if idx + offset < len(df):
                        data_row = df.iloc[idx + offset]
                        zone_label = str(data_row.iloc[0]).strip() if pd.notna(data_row.iloc[0]) else ""
                        
                        if zone_label in zones:
                            # Warehouse costs usually in column 5 (index 5)
                            warehouse_cost = parse_numeric(data_row.iloc[5]) if len(data_row) > 5 else 0
                            if warehouse_cost > 0:
                                data['penalties'][zone_label] = warehouse_cost
                                if zone_label not in data['zones']:
                                    data['zones'][zone_label] = {}
                                data['zones'][zone_label]['warehouse_cost'] = warehouse_cost
        
        return data
    except Exception as e:
        st.warning(f"Error loading logistics data: {e}")
        return {'zones': {}, 'shipping_costs': {}, 'benchmarks': {}, 'penalties': {}, 'raw_df': None}


def load_machine_spaces(file) -> Dict[str, Any]:
    """Load machine_spaces.xlsx - Production input for machine capacity by zone."""
    try:
        df = pd.read_excel(file, header=None)
        data = {'zones': {}, 'total_capacity': 0, 'raw_df': df}
        
        zones = ['Center', 'West', 'North', 'East', 'South']
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
            
            if 'capacity' in first_val or 'machine' in first_val:
                for z_idx, zone in enumerate(zones):
                    if z_idx + 1 < len(row):
                        if zone not in data['zones']:
                            data['zones'][zone] = {}
                        data['zones'][zone]['machine_capacity'] = parse_numeric(row.iloc[z_idx + 1])
            
            if 'available' in first_val or 'spaces' in first_val:
                for z_idx, zone in enumerate(zones):
                    if z_idx + 1 < len(row):
                        if zone not in data['zones']:
                            data['zones'][zone] = {}
                        data['zones'][zone]['available_spaces'] = parse_numeric(row.iloc[z_idx + 1])
        
        # Calculate total capacity
        data['total_capacity'] = sum(
            zone_data.get('machine_capacity', 0) 
            for zone_data in data['zones'].values()
        )
        
        return data
    except Exception as e:
        st.warning(f"Error loading machine spaces: {e}")
        return {'zones': {}, 'total_capacity': 0, 'raw_df': None}


def load_sales_data(file) -> Dict[str, Any]:
    """
    Load sales_admin_expenses.xlsx - CMO input for Last Sales per zone.
    Extracts: units sold by zone, price by zone, and marketing spend totals.
    """
    ZONES = ['Center', 'West', 'North', 'East', 'South']
    
    try:
        df = pd.read_excel(file, header=None)
        data = {
            'by_zone': {zone: {'units': 0, 'price': 0} for zone in ZONES},
            'totals': {'units': 0, 'tv_spend': 0, 'radio_spend': 0, 'salespeople_cost': 0},
            'raw_df': df
        }
        
        in_sales_section = False
        in_expense_section = False
        
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            
            # Detect sections
            if 'sales' in first_val.lower() and 'expense' not in first_val.lower() and 'admin' not in first_val.lower():
                in_sales_section = True
                in_expense_section = False
                continue
            elif 'expense' in first_val.lower() or 'admin' in first_val.lower():
                in_sales_section = False
                in_expense_section = True
                continue
            
            # Parse sales data rows (Region, Brand, Units, Local Price, ...)
            if in_sales_section:
                for zone in ZONES:
                    if first_val.lower() == zone.lower():
                        # Format: Region, Brand, Units, Local Price, Gross Sales, ...
                        units = parse_numeric(row.iloc[2]) if len(row) > 2 else 0
                        price = parse_numeric(row.iloc[3]) if len(row) > 3 else 0
                        
                        if units > 0:
                            data['by_zone'][zone]['units'] = units
                            data['totals']['units'] += units
                        if price > 0:
                            data['by_zone'][zone]['price'] = price
                        break
            
            # Parse expense data rows
            if in_expense_section:
                expense = parse_numeric(row.iloc[2]) if len(row) > 2 else 0
                if 'tv' in first_val.lower() and 'advert' in first_val.lower():
                    data['totals']['tv_spend'] = expense
                elif 'radio' in first_val.lower() and 'advert' in first_val.lower():
                    data['totals']['radio_spend'] = expense
                elif 'salespeople' in first_val.lower() and 'salar' in first_val.lower():
                    data['totals']['salespeople_cost'] = expense
        
        return data
    except Exception as e:
        st.warning(f"Error loading sales data: {e}")
        return {
            'by_zone': {zone: {'units': 0, 'price': 0} for zone in ['Center', 'West', 'North', 'East', 'South']},
            'totals': {'units': 0, 'tv_spend': 0, 'radio_spend': 0, 'salespeople_cost': 0},
            'raw_df': None
        }
