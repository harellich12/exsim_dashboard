import pandas as pd
import openpyxl
import xml.etree.ElementTree as ET
import re
import os
from openpyxl.utils.dataframe import dataframe_to_rows

# --- CONFIGURATION ---
SOURCE_FILE = 'Reports/market-report.xls'
TARGET_TEMPLATE = 'Reports/Demand Planner vs02.xlsx'
OUTPUT_FILE = 'dashboards_v2/Demand_Planner_Filled.xlsx'
TEAM_ID = "A3" # Usage: To identify "My Share" if needed, though we map all companies

# --- 1. XML PARSER FOR SPREADSHEETML ---
class SpreadsheetMLParser:
    def __init__(self, filepath):
        self.filepath = filepath
        self.ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}
        self.tree = ET.parse(filepath)
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
        sheet = self.root.find('.//ss:Worksheet', self.ns)
        table = sheet.find('ss:Table', self.ns)
        
        rows = table.findall('ss:Row', self.ns)
        
        current_section = None
        header_map = {} # Col Index -> Company/Zone/Segment
        
        # Iteration state helpers
        current_zone = None
        
        print("Scannning source file structure...")
        
        for row_idx, row in enumerate(rows):
            cells = row.findall('ss:Cell', self.ns)
            row_data = []
            for cell in cells:
                data = cell.find('ss:Data', self.ns)
                text = data.text if data is not None and data.text is not None else ""
                
                # Handle MergeAcross (implies the cell spans multiple columns)
                # For this specific file, simple text extraction is mostly enough if we track indices
                # But we need to be careful about strict column alignment.
                # Given the XML structure analysis, we can rely on text signatures.
                row_data.append(text)
            
            # Combine all text to identify section headers
            full_row_text = " ".join([str(x) for x in row_data]).strip()
            
            # --- SECTION DETECTION ---
            if "Market Share Per Region (%)" in full_row_text and "Segment" not in full_row_text:
                current_section = "market_share_region"
                continue
            elif "Market Share Per Region Per Segment (%)" in full_row_text:
                current_section = "market_share_segment"
                continue
            elif row_data and row_data[0].strip().lower() == 'price' or 'Price' in row_data[0]:
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
            
            # --- DATA EXTRACTION ---
            if current_section == "market_share_region":
                self._parse_market_share_region(row_data)
            elif current_section == "market_share_segment":
                self._parse_market_share_segment(row_data)
            elif current_section == "price":
                self._parse_price(row_data)
            elif current_section == "awareness":
                self._parse_awareness(row_data)
            elif current_section == "attractiveness":
                self._parse_attractiveness(row_data)
            elif current_section == "salesforce":
                self._parse_salesforce(row_data)
                
    def _parse_market_share_region(self, row_data):
        # Header: Zone, A1..., A2..., A3..., A4...
        # Data: Center, 38.2, 7.9, 37.8, 16.1
        if not row_data: return
        
        first_cell = row_data[0].strip()
        if first_cell in ['Center', 'West', 'North', 'East', 'South']:
            # Assuming fixed columns 1=A1, 2=A2, 3=A3, 4=A4 based on file analysis
            # XML index is 0-based in list, but col index 0 is Zone.
            # Col 1: A1, Col 2: A2, Col 3: A3, Col 4: A4
            if len(row_data) >= 5:
                zone = first_cell
                self.data_store['market_share_region'].setdefault('A1', {})[zone] = self._clean_num(row_data[1])
                self.data_store['market_share_region'].setdefault('A2', {})[zone] = self._clean_num(row_data[2])
                self.data_store['market_share_region'].setdefault('A3', {})[zone] = self._clean_num(row_data[3])
                self.data_store['market_share_region'].setdefault('A4', {})[zone] = self._clean_num(row_data[4])

    def _parse_market_share_segment(self, row_data):
        # Format is tricky based on analysis:
        # Row 189: "Center", "High", 33.1, 24.3, 22.0, 20.7 ...
        # But wait, lines 189+ in view_file showed:
        # Cell 0: Center, Cell 1: High, Cell 2: A1, Cell 3: A2...
        if not row_data or len(row_data) < 3: return
        
        if row_data[0].strip() in ['Center', 'West', 'North', 'East', 'South']:
            self._current_ms_zone = row_data[0].strip() # Remember Zone for subsequent lines if needed? 
            # Looking at file, Zone is only present on the first row of the block? 
            # Row 190 Center High
            # Row 203 (Empty) Low
            pass
        
        # Detect Segment
        segment = None
        vals_start_idx = 0
        
        # Check col 1 for segment
        if len(row_data) > 1 and row_data[1].strip() in ['High', 'Low']:
            segment = row_data[1].strip()
            # If Zone is empty at col 0, use stored
            if not row_data[0].strip() and hasattr(self, '_current_ms_zone'):
                zone = self._current_ms_zone
            else:
                zone = row_data[0].strip()
                self._current_ms_zone = zone
            
            vals_start_idx = 2
            
            if zone and segment:
                 # Col 2: A1, 3: A2, 4: A3, 5: A4
                 if len(row_data) >= 6:
                     self._store_segment_data('market_share_segment', zone, segment, row_data, vals_start_idx)

    def _parse_price(self, row_data):
        # Simple table: Zone (Col 0) | A1(1) | A2(2) | A3(3) | A4(4)
        if not row_data: return
        if row_data[0].strip() in ['Center', 'West', 'North', 'East', 'South']:
            zone = row_data[0].strip()
            if len(row_data) >= 5:
                # Store per company
                self.data_store['price'].setdefault('A1', {})[zone] = self._clean_num(row_data[1])
                self.data_store['price'].setdefault('A2', {})[zone] = self._clean_num(row_data[2])
                self.data_store['price'].setdefault('A3', {})[zone] = self._clean_num(row_data[3])
                self.data_store['price'].setdefault('A4', {})[zone] = self._clean_num(row_data[4])

    def _parse_awareness(self, row_data):
        # Structure seems identical to market_share_segment based on analysis
        # Row 504: Center | High | 60.71 | 72.62 ...
        self._parse_generic_segment_table(row_data, 'awareness')

    def _parse_attractiveness(self, row_data):
        self._parse_generic_segment_table(row_data, 'attractiveness')
        
    def _parse_salesforce(self, row_data):
        # Structure seems identical to Price (Zone | A1..A4)
        self._parse_generic_zone_table(row_data, 'salesforce')

    def _parse_generic_segment_table(self, row_data, key):
        if not row_data or len(row_data) < 3: return
        
        # Track current zone via attribute name specific to key to avoid collision
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
                self._store_segment_data(key, zone, segment_candidate, row_data, 2)
    
    def _parse_generic_zone_table(self, row_data, key):
         if not row_data: return
         if row_data[0].strip() in ['Center', 'West', 'North', 'East', 'South']:
            zone = row_data[0].strip()
            if len(row_data) >= 5:
                self.data_store[key].setdefault('A1', {})[zone] = self._clean_num(row_data[1])
                self.data_store[key].setdefault('A2', {})[zone] = self._clean_num(row_data[2])
                self.data_store[key].setdefault('A3', {})[zone] = self._clean_num(row_data[3])
                self.data_store[key].setdefault('A4', {})[zone] = self._clean_num(row_data[4])

    def _store_segment_data(self, key, zone, segment, row_data, start_idx):
        self.data_store[key].setdefault('A1', {}).setdefault(zone, {})[segment] = self._clean_num(row_data[start_idx])
        self.data_store[key].setdefault('A2', {}).setdefault(zone, {})[segment] = self._clean_num(row_data[start_idx+1])
        self.data_store[key].setdefault('A3', {}).setdefault(zone, {})[segment] = self._clean_num(row_data[start_idx+2])
        self.data_store[key].setdefault('A4', {}).setdefault(zone, {})[segment] = self._clean_num(row_data[start_idx+3])

    def _clean_num(self, val):
        try:
            if isinstance(val, str):
                val = val.strip()
                if not val: return 0.0
            return float(val)
        except:
            return 0.0

# --- 2. MAIN ETL PROCESS ---

def run_mapping():
    print(f"Loading source: {SOURCE_FILE}")
    parser = SpreadsheetMLParser(SOURCE_FILE)
    parser.parse()
    
    # Verify we extracted something
    if not parser.data_store['market_share_region']:
        print("ERROR: Failed to extract market share data. Check parser logic.")
        return

    print("Data extracted successfully. Constructing output DataFrame...")
    
    # Load Target Schema (or use known schema as fallback)
    KNOWN_SCHEMA = ['Period', 'Company', 'Region', 'Segment', 'Run Type', 
                    'Market Share Region %', 'Market Share Segment %', 'Price', 
                    'Awareness %', 'Attractiveness %', 'Salesforce Effectiveness %']
    
    try:
        print(f"Loading target schema from: {TARGET_TEMPLATE}")
        df_template = pd.read_excel(TARGET_TEMPLATE, sheet_name='MARKET_DATA')
        columns = [c for c in df_template.columns if not str(c).startswith('Unnamed')]
    except Exception as e:
        print(f"Could not read template ({e}). Using known schema.")
        columns = KNOWN_SCHEMA
    
    print(f"Target Schema: {columns}")
    
    # --- 3. MAPPING ENGINE ---
    # Generate rows for output
    output_rows = []
    
    COMPANIES = ['A1', 'A2', 'A3', 'A4']
    ZONES = ['Center', 'West', 'North', 'East', 'South']
    SEGMENTS = ['High', 'Low']
    
    for zone in ZONES:
        for segment in SEGMENTS:
            for company in COMPANIES:
                row = {}
                
                # Retrieve Metric Helper
                def get_data(store_key, use_segment=False):
                    try:
                        c_data = parser.data_store[store_key].get(company, {})
                        z_data = c_data.get(zone, {})
                        if use_segment:
                            if isinstance(z_data, dict):
                                return z_data.get(segment, 0.0)
                            return 0.0 # Should be dict
                        else:
                            # If z_data is a value, return it. If it's a dict (shouldn't be for non-segment keys), issue
                            if isinstance(z_data, dict):
                                return 0.0 # Warning?
                            return z_data
                    except:
                        return 0.0

                # --- MAPPING LOGIC ---
                row['Period'] = 7 # Hardcoded extraction from analysis
                row['Company'] = company
                row['Region'] = zone
                row['Segment'] = segment
                row['Run Type'] = 'Real'
                
                row['Market Share Region %'] = get_data('market_share_region', use_segment=False)
                row['Market Share Segment %'] = get_data('market_share_segment', use_segment=True)
                row['Price'] = get_data('price', use_segment=False) # Broadcast zone price to segment
                row['Awareness %'] = get_data('awareness', use_segment=True)
                row['Attractiveness %'] = get_data('attractiveness', use_segment=True)
                row['Salesforce Effectiveness %'] = get_data('salesforce', use_segment=False) # Broadcast zone val
                
                # Order row according to schema
                ordered_row = [row.get(col, 0) for col in columns]
                output_rows.append(ordered_row)

    df_out = pd.DataFrame(output_rows, columns=columns)
    
    print("\n--- Validation Report ---")
    print(df_out.head())
    print(f"Total Rows Generated: {len(df_out)}")
    
    if (df_out['Salesforce Effectiveness %'] == 0).all():
        print("WARNING: All 'Salesforce Effectiveness %' values are 0. Check mapping.")
    else:
        print("SUCCESS: 'Salesforce Effectiveness %' mapped successfully.")

    # --- 4. EXPORT ---
    print(f"Writing to {OUTPUT_FILE}...")
    
    # Use openpyxl to output cleanly
    if not os.path.exists(os.path.dirname(OUTPUT_FILE)):
        os.makedirs(os.path.dirname(OUTPUT_FILE))
    
    # Create a fresh workbook with only MARKET_DATA sheet
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'MARKET_DATA'
    
    # Write header
    ws.append(columns)
    
    # Write data
    for r in dataframe_to_rows(df_out, index=False, header=False):
        ws.append(r)
        
    wb.save(OUTPUT_FILE)
    print("Done.")

if __name__ == "__main__":
    run_mapping()
