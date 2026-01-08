"""
Market Data Mapper Utility
Transforms raw market-report.xls data into Demand Planner format.
"""

import pandas as pd
import openpyxl
import xml.etree.ElementTree as ET
import re
import io
from openpyxl.utils.dataframe import dataframe_to_rows


# Known target schema
TARGET_SCHEMA = ['Period', 'Company', 'Region', 'Segment', 'Run Type', 
                'Market Share Region %', 'Market Share Segment %', 'Price', 
                'Awareness %', 'Attractiveness %', 'Salesforce Effectiveness %']


class SpreadsheetMLParser:
    """Parse SpreadsheetML XML format used by market-report.xls"""
    
    def __init__(self, file_content):
        """Initialize with file content (bytes or file-like object)."""
        self.ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}
        
        if hasattr(file_content, 'read'):
            content = file_content.read()
            if hasattr(file_content, 'seek'):
                file_content.seek(0)
        else:
            content = file_content
            
        self.tree = ET.parse(io.BytesIO(content) if isinstance(content, bytes) else io.StringIO(content))
        self.root = self.tree.getroot()
        self.data_store = {
            'market_share_region': {},
            'market_share_segment': {},
            'price': {},
            'awareness': {},
            'attractiveness': {},
            'salesforce': {}
        }
    
    def parse(self):
        """Parse the XML and extract all data tables."""
        sheet = self.root.find('.//ss:Worksheet', self.ns)
        table = sheet.find('ss:Table', self.ns)
        rows = table.findall('ss:Row', self.ns)
        
        current_section = None
        
        for row_idx, row in enumerate(rows):
            cells = row.findall('ss:Cell', self.ns)
            row_data = []
            for cell in cells:
                data = cell.find('ss:Data', self.ns)
                text = data.text if data is not None and data.text is not None else ""
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
                self._parse_zone_table(row_data, 'market_share_region')
            elif current_section == "market_share_segment":
                self._parse_segment_table(row_data, 'market_share_segment')
            elif current_section == "price":
                self._parse_zone_table(row_data, 'price')
            elif current_section == "awareness":
                self._parse_segment_table(row_data, 'awareness')
            elif current_section == "attractiveness":
                self._parse_segment_table(row_data, 'attractiveness')
            elif current_section == "salesforce":
                self._parse_zone_table(row_data, 'salesforce')
                
    def _parse_zone_table(self, row_data, key):
        """Parse simple Zone | A1 | A2 | A3 | A4 tables."""
        if not row_data:
            return
        if row_data[0].strip() in ['Center', 'West', 'North', 'East', 'South']:
            zone = row_data[0].strip()
            if len(row_data) >= 5:
                self.data_store[key].setdefault('A1', {})[zone] = self._clean_num(row_data[1])
                self.data_store[key].setdefault('A2', {})[zone] = self._clean_num(row_data[2])
                self.data_store[key].setdefault('A3', {})[zone] = self._clean_num(row_data[3])
                self.data_store[key].setdefault('A4', {})[zone] = self._clean_num(row_data[4])

    def _parse_segment_table(self, row_data, key):
        """Parse Zone | Segment | A1 | A2 | A3 | A4 tables."""
        if not row_data or len(row_data) < 3:
            return
        
        zone_attr = f'_current_{key}_zone'
        zone_candidate = row_data[0].strip()
        segment_candidate = row_data[1].strip() if len(row_data) > 1 else ""
        
        if zone_candidate in ['Center', 'West', 'North', 'East', 'South']:
            setattr(self, zone_attr, zone_candidate)
            zone = zone_candidate
        elif hasattr(self, zone_attr):
            zone = getattr(self, zone_attr)
        else:
            return

        if segment_candidate in ['High', 'Low']:
            if len(row_data) >= 6:
                self.data_store[key].setdefault('A1', {}).setdefault(zone, {})[segment_candidate] = self._clean_num(row_data[2])
                self.data_store[key].setdefault('A2', {}).setdefault(zone, {})[segment_candidate] = self._clean_num(row_data[3])
                self.data_store[key].setdefault('A3', {}).setdefault(zone, {})[segment_candidate] = self._clean_num(row_data[4])
                self.data_store[key].setdefault('A4', {}).setdefault(zone, {})[segment_candidate] = self._clean_num(row_data[5])

    def _clean_num(self, val):
        try:
            if isinstance(val, str):
                val = val.strip()
                if not val:
                    return 0.0
            return float(val)
        except:
            return 0.0


def generate_formatted_market_data(source_file) -> bytes:
    """
    Generate formatted market data Excel from source market report.
    Handles both SpreadsheetML XML (.xls) and standard Excel (.xlsx) formats.
    
    Args:
        source_file: File-like object or bytes containing the market report
        
    Returns:
        bytes: Excel file content ready for download
    """
    # Read source content
    if hasattr(source_file, 'read'):
        content = source_file.read()
        if hasattr(source_file, 'seek'):
            source_file.seek(0)
    else:
        content = source_file
    
    # Detect format
    is_xml = content[:100].decode('utf-8', errors='ignore').strip().startswith('<?xml')
    
    columns = TARGET_SCHEMA
    output_rows = []
    
    COMPANIES = ['A1', 'A2', 'A3', 'A4']
    ZONES = ['Center', 'West', 'North', 'East', 'South']
    SEGMENTS = ['High', 'Low']
    
    if is_xml:
        # Use XML parser for SpreadsheetML format
        parser = SpreadsheetMLParser(content)
        parser.parse()
        
        for zone in ZONES:
            for segment in SEGMENTS:
                for company in COMPANIES:
                    def get_data(store_key, use_segment=False):
                        try:
                            c_data = parser.data_store[store_key].get(company, {})
                            z_data = c_data.get(zone, {})
                            if use_segment:
                                if isinstance(z_data, dict):
                                    return z_data.get(segment, 0.0)
                                return 0.0
                            else:
                                if isinstance(z_data, dict):
                                    return 0.0
                                return z_data
                        except:
                            return 0.0

                    row = {
                        'Period': 7,
                        'Company': company,
                        'Region': zone,
                        'Segment': segment,
                        'Run Type': 'Real',
                        'Market Share Region %': get_data('market_share_region', use_segment=False),
                        'Market Share Segment %': get_data('market_share_segment', use_segment=True),
                        'Price': get_data('price', use_segment=False),
                        'Awareness %': get_data('awareness', use_segment=True),
                        'Attractiveness %': get_data('attractiveness', use_segment=True),
                        'Salesforce Effectiveness %': get_data('salesforce', use_segment=False),
                    }
                    
                    ordered_row = [row.get(col, 0) for col in columns]
                    output_rows.append(ordered_row)
    else:
        # For standard Excel, parse directly to extract ALL companies
        if hasattr(source_file, 'seek'):
            source_file.seek(0)
        df = pd.read_excel(source_file if hasattr(source_file, 'read') else io.BytesIO(content), header=None)
        
        # Storage for all companies
        store = {
            'market_share_region': {},
            'market_share_segment': {},
            'price': {},
            'awareness': {},
            'attractiveness': {},
            'salesforce': {},
        }
        
        current_section = None
        company_cols = {}  # col_idx -> company_id (A1, A2, A3, A4)
        last_zone = None
        
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            second_val = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ''
            
            # Detect section headers
            if 'market share' in first_val.lower() and 'segment' in first_val.lower():
                current_section = 'market_share_segment'
                company_cols = {}
            elif 'market share' in first_val.lower() and 'region' in first_val.lower():
                current_section = 'market_share_region'
                company_cols = {}
            elif 'awareness' in first_val.lower():
                current_section = 'awareness'
                company_cols = {}
            elif 'attractiveness' in first_val.lower():
                current_section = 'attractiveness'
                company_cols = {}
            elif 'price' in first_val.lower() and current_section != 'price':
                current_section = 'price'
                company_cols = {}
            elif 'promotional' in first_val.lower() or 'salesforce' in first_val.lower():
                current_section = 'salesforce'
                company_cols = {}
            
            # Detect column headers with company names
            if first_val.lower() == 'zone' or first_val.lower() == 'region':
                company_cols = {}
                for col_idx in range(len(row)):
                    col_val = str(row.iloc[col_idx]).strip() if pd.notna(row.iloc[col_idx]) else ''
                    for comp in COMPANIES:
                        if col_val.startswith(comp):
                            company_cols[col_idx] = comp
                            break
            
            # Parse zone data rows
            if first_val in ZONES:
                last_zone = first_val
            
            current_zone = first_val if first_val in ZONES else (last_zone if first_val == '' else None)
            current_segment = None
            
            # Check for segment
            if second_val in SEGMENTS:
                current_segment = second_val
            
            # Parse data if we have zone and company columns
            if current_zone and company_cols and current_section:
                for col_idx, company in company_cols.items():
                    if col_idx < len(row):
                        try:
                            val = float(row.iloc[col_idx]) if pd.notna(row.iloc[col_idx]) else 0.0
                        except:
                            val = 0.0
                        
                        if current_section == 'market_share_region' and val > 0:
                            store['market_share_region'].setdefault(company, {})[current_zone] = val
                        elif current_section == 'market_share_segment' and current_segment and val > 0:
                            store['market_share_segment'].setdefault(company, {}).setdefault(current_zone, {})[current_segment] = val
                        elif current_section == 'price' and val > 0:
                            store['price'].setdefault(company, {})[current_zone] = val
                        elif current_section == 'awareness' and current_segment and val > 0:
                            store['awareness'].setdefault(company, {}).setdefault(current_zone, {})[current_segment] = val
                        elif current_section == 'attractiveness' and current_segment and val > 0:
                            store['attractiveness'].setdefault(company, {}).setdefault(current_zone, {})[current_segment] = val
                        elif current_section == 'salesforce' and val > 0:
                            store['salesforce'].setdefault(company, {})[current_zone] = val
        
        # Generate output rows for all companies
        for zone in ZONES:
            for segment in SEGMENTS:
                for company in COMPANIES:
                    row = {
                        'Period': 7,
                        'Company': company,
                        'Region': zone,
                        'Segment': segment,
                        'Run Type': 'Real',
                        'Market Share Region %': store['market_share_region'].get(company, {}).get(zone, 0),
                        'Market Share Segment %': store['market_share_segment'].get(company, {}).get(zone, {}).get(segment, 0),
                        'Price': store['price'].get(company, {}).get(zone, 0),
                        'Awareness %': store['awareness'].get(company, {}).get(zone, {}).get(segment, 0),
                        'Attractiveness %': store['attractiveness'].get(company, {}).get(zone, {}).get(segment, 0),
                        'Salesforce Effectiveness %': store['salesforce'].get(company, {}).get(zone, 0),
                    }
                    
                    ordered_row = [row.get(col, 0) for col in columns]
                    output_rows.append(ordered_row)

    df_out = pd.DataFrame(output_rows, columns=columns)
    
    # Create Excel workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'MARKET_DATA'
    
    # Write header
    ws.append(columns)
    
    # Write data
    for r in dataframe_to_rows(df_out, index=False, header=False):
        ws.append(r)
    
    # Save to bytes
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output.getvalue()

