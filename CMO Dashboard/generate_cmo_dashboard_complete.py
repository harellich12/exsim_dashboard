"""
ExSim CMO Complete Dashboard - Market Allocation & Strategy

Integrates Marketing Decisions, Innovation Decisions, Inventory Checks,
and Segment Analysis into a single cohesive decision-support tool.

Required libraries: pandas, openpyxl
"""

import pandas as pd
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import DataBarRule, FormulaRule, IconSetRule
from openpyxl.chart import ScatterChart, BarChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.chart.label import DataLabelList
import warnings

warnings.filterwarnings('ignore')

# =============================================================================
# CONFIGURATION
# =============================================================================
DATA_FOLDER = Path("data")
OUTPUT_FILE = "CMO_Dashboard_Complete.xlsx"
MY_COMPANY = "Company 3"

ZONES = ["Center", "West", "North", "East", "South"]
SEGMENTS = ["High", "Low"]

# Defaults
DEFAULT_PRICE = 100
DEFAULT_AWARENESS = 50
DEFAULT_ATTRACTIVENESS = 50
DEFAULT_COGS = 40
DEFAULT_SALESPEOPLE_SALARY = 5000


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def parse_numeric(value):
    """Parse formatted number strings."""
    if pd.isna(value):
        return 0
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = str(value).replace('$', '').replace(',', '').replace('%', '').replace(' ', '').strip()
    try:
        return float(cleaned)
    except:
        return 0


def load_excel_file(filepath, sheet_name=None):
    """Load Excel file, optionally from specific sheet."""
    try:
        if sheet_name:
            return pd.read_excel(filepath, sheet_name=sheet_name, header=None)
        return pd.read_excel(filepath, header=None)
    except Exception as e:
        print(f"Warning: Could not load {filepath}: {e}")
        return None


# =============================================================================
# DATA LOADING FUNCTIONS
# =============================================================================

def load_market_report(filepath):
    """Load market report with segment-level data."""
    df = load_excel_file(filepath)
    
    data = {
        'by_segment': {seg: {zone: {
            'my_market_share': 25,
            'my_awareness': DEFAULT_AWARENESS,
            'my_attractiveness': DEFAULT_ATTRACTIVENESS,
            'my_price': DEFAULT_PRICE,
            'comp_avg_awareness': DEFAULT_AWARENESS,
            'comp_avg_price': DEFAULT_PRICE
        } for zone in ZONES} for seg in SEGMENTS},
        'zones': {zone: {
            'my_price': DEFAULT_PRICE,
            'comp_avg_price': DEFAULT_PRICE,
            'my_awareness': DEFAULT_AWARENESS,
            'my_attractiveness': DEFAULT_ATTRACTIVENESS,
            'my_market_share': 25
        } for zone in ZONES}
    }
    
    if df is None:
        return data
    
    current_section = None
    current_zone = None
    current_segment = None
    
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        second_val = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ''
        
        # Detect sections
        if 'market share' in first_val.lower() and 'segment' in first_val.lower():
            current_section = 'segment_share'
        elif 'market share' in first_val.lower() and 'region' in first_val.lower():
            current_section = 'region_share'
        elif 'awareness' in first_val.lower() and 'segment' in first_val.lower():
            current_section = 'segment_awareness'
        elif 'awareness' in first_val.lower():
            current_section = 'awareness'
        elif 'attractiveness' in first_val.lower():
            current_section = 'attractiveness'
        elif 'price' in first_val.lower() and 'zone' not in first_val.lower():
            current_section = 'price'
        
        # Zone detection
        for zone in ZONES:
            if first_val.lower() == zone.lower():
                current_zone = zone
                
                # Check for segment in second column
                if second_val.lower() in ['high', 'low']:
                    current_segment = second_val.capitalize()
                else:
                    current_segment = None
                
                # Extract data based on section
                if current_section == 'segment_share' and current_segment:
                    for col_idx in range(2, min(6, len(row))):
                        val = parse_numeric(row.iloc[col_idx])
                        if val > 0:
                            data['by_segment'][current_segment][zone]['my_market_share'] = val
                            break
                            
                elif current_section == 'region_share':
                    for col_idx in range(1, min(6, len(row))):
                        val = parse_numeric(row.iloc[col_idx])
                        if val > 0:
                            data['zones'][zone]['my_market_share'] = val
                            break
                            
                elif current_section == 'price':
                    prices = []
                    for col_idx in range(1, min(6, len(row))):
                        val = parse_numeric(row.iloc[col_idx])
                        if val > 0:
                            prices.append(val)
                    if prices:
                        data['zones'][zone]['my_price'] = prices[0]
                        if len(prices) > 1:
                            data['zones'][zone]['comp_avg_price'] = sum(prices[1:]) / len(prices[1:])
                        # Copy to segment data
                        for seg in SEGMENTS:
                            data['by_segment'][seg][zone]['my_price'] = prices[0]
                            if len(prices) > 1:
                                data['by_segment'][seg][zone]['comp_avg_price'] = sum(prices[1:]) / len(prices[1:])
                break
        
        # Segment rows (continuation)
        if second_val.lower() in ['high', 'low'] and current_zone:
            current_segment = second_val.capitalize()
            if current_section == 'segment_share':
                for col_idx in range(2, min(6, len(row))):
                    val = parse_numeric(row.iloc[col_idx])
                    if val > 0:
                        data['by_segment'][current_segment][current_zone]['my_market_share'] = val
                        break
    
    return data


def load_innovation_template(filepath):
    """Load innovation features dynamically."""
    df = load_excel_file(filepath, sheet_name='Innovation')
    
    features = []
    
    if df is None:
        # Default features
        features = [
            "STAINLESS MATERIAL", "RECYCLABLE MATERIALS", "ENERGY EFFICIENCY",
            "LIGHTER AND MORE COMPACT", "IMPACT RESISTANCE", "NOISE REDUCTION",
            "IMPROVED BATTERY CAPACITY", "SELF-CLEANING", "SPEED SETTINGS",
            "DIGITAL CONTROLS", "VOICE ASSISTANCE INTEGRATION",
            "AUTOMATION AND PROGRAMMABILITY", "MULTIFUNCTIONAL ACCESSORIES",
            "MAPPING TECHNOLOGY"
        ]
    else:
        for idx, row in df.iterrows():
            # Look for Improvement column data
            if len(row) > 1:
                improvement = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
                if improvement and improvement.upper() not in ['IMPROVEMENT', 'NAN', '']:
                    features.append(improvement)
    
    return features


def load_marketing_template(filepath):
    """Load marketing template structure."""
    df = load_excel_file(filepath, sheet_name='Marketing')
    
    template = {
        'df': df,
        'tv_budget': 35,
        'brand_focus': 50,
        'radio_budgets': {zone: 100 for zone in ZONES},
        'demand': {zone: 0 for zone in ZONES},
        'prices': {zone: 68 for zone in ZONES},
        'payment_terms': {zone: 'B' for zone in ZONES},
        'salespeople': {zone: 10 for zone in ZONES}
    }
    
    if df is not None:
        for idx, row in df.iterrows():
            first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            
            # TV budget
            if first_val.upper() == 'A' and len(row) > 2:
                channel = str(row.iloc[2]).strip().lower() if pd.notna(row.iloc[2]) else ''
                if channel == 'tv':
                    template['tv_budget'] = parse_numeric(row.iloc[3])
                    template['brand_focus'] = parse_numeric(row.iloc[4])
                elif channel == 'radio':
                    zone = str(row.iloc[1]).strip()
                    if zone in ZONES:
                        template['radio_budgets'][zone] = parse_numeric(row.iloc[3])
            
            # Demand (column 7)
            for zone in ZONES:
                zone_val = str(row.iloc[7]).strip() if len(row) > 7 and pd.notna(row.iloc[7]) else ''
                if zone_val.lower() == zone.lower():
                    template['demand'][zone] = parse_numeric(row.iloc[8]) if len(row) > 8 else 0
    
    return template


def load_sales_data(filepath):
    """Load sales and expenses data."""
    df = load_excel_file(filepath)
    
    data = {
        'by_zone': {zone: {'units': 1000, 'price': DEFAULT_PRICE} for zone in ZONES},
        'totals': {'units': 0, 'tv_spend': 0, 'radio_spend': 0, 'salespeople_cost': 0}
    }
    
    if df is None:
        return data
    
    current_zone = None
    
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        
        for zone in ZONES:
            if zone.lower() == first_val.lower():
                current_zone = zone
                break
        
        if 'units' in first_val.lower() and current_zone:
            for col_idx in range(1, min(12, len(row))):
                val = parse_numeric(row.iloc[col_idx])
                if val > 0:
                    data['by_zone'][current_zone]['units'] = val
                    data['totals']['units'] += val
                    break
    
    return data


def load_inventory_data(filepath):
    """Load inventory to detect stockouts."""
    df = load_excel_file(filepath)
    
    data = {'final_inventory': 500, 'is_stockout': False}
    
    if df is None:
        return data
    
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        
        if 'final' in first_val.lower() and 'inventory' in first_val.lower():
            # Get fortnight 8 value
            final_val = parse_numeric(row.iloc[8]) if len(row) > 8 else 0
            data['final_inventory'] = final_val
            data['is_stockout'] = final_val <= 0
            break
    
    return data


# =============================================================================
# EXCEL GENERATION
# =============================================================================

def create_complete_dashboard(market_data, innovation_features, marketing_template, 
                              sales_data, inventory_data):
    """Create the complete 5-tab CMO Dashboard."""
    
    wb = Workbook()
    
    # Styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    section_font = Font(bold=True, size=12, color="2F5496")
    title_font = Font(bold=True, size=14, color="2F5496")
    input_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    calc_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    output_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    ref_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    orange_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # =========================================================================
    # TAB 1: SEGMENT_PULSE
    # =========================================================================
    ws1 = wb.active
    ws1.title = "SEGMENT_PULSE"
    
    ws1['A1'] = "SEGMENT PULSE - Market Allocation Drivers"
    ws1['A1'].font = title_font
    
    row = 3
    for segment in SEGMENTS:
        ws1.cell(row=row, column=1, value=f"{segment.upper()} SEGMENT ANALYSIS").font = section_font
        row += 1
        
        # Headers
        seg_headers = ['Zone', 'My Market Share', 'Awareness Gap', 'Price Gap', 
                       'Attractiveness', 'Allocation Flag']
        for col, header in enumerate(seg_headers, start=1):
            cell = ws1.cell(row=row, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
        row += 1
        
        data_start_row = row
        
        for zone in ZONES:
            zone_seg = market_data['by_segment'][segment].get(zone, {})
            zone_data = market_data['zones'].get(zone, {})
            
            market_share = zone_seg.get('my_market_share', 25)
            my_awareness = zone_data.get('my_awareness', DEFAULT_AWARENESS)
            comp_awareness = zone_seg.get('comp_avg_awareness', DEFAULT_AWARENESS)
            awareness_gap = my_awareness - comp_awareness
            
            my_price = zone_data.get('my_price', DEFAULT_PRICE)
            comp_price = zone_data.get('comp_avg_price', DEFAULT_PRICE)
            price_gap = ((my_price - comp_price) / comp_price * 100) if comp_price > 0 else 0
            
            attractiveness = zone_data.get('my_attractiveness', DEFAULT_ATTRACTIVENESS)
            
            # Allocation flag logic
            if segment == "High":
                if my_awareness < 30:
                    flag = "CRITICAL: Boost TV for Allocation"
                    flag_fill = red_fill
                else:
                    flag = "OK"
                    flag_fill = output_fill
            else:  # Low segment
                if price_gap > 5:
                    flag = "RISK: Losing Volume to Price"
                    flag_fill = orange_fill
                else:
                    flag = "OK"
                    flag_fill = output_fill
            
            ws1.cell(row=row, column=1, value=zone).border = thin_border
            
            cell = ws1.cell(row=row, column=2, value=market_share)
            cell.border = thin_border
            cell.number_format = '0.0%' if market_share <= 1 else '0.0'
            
            cell = ws1.cell(row=row, column=3, value=awareness_gap)
            cell.border = thin_border
            if awareness_gap < 0:
                cell.fill = red_fill
            
            cell = ws1.cell(row=row, column=4, value=price_gap / 100)
            cell.border = thin_border
            cell.number_format = '0.0%'
            
            ws1.cell(row=row, column=5, value=attractiveness).border = thin_border
            
            cell = ws1.cell(row=row, column=6, value=flag)
            cell.border = thin_border
            cell.fill = flag_fill
            cell.font = Font(bold=True)
            
            row += 1
        
        # Add data bars for market share
        ws1.conditional_formatting.add(
            f'B{data_start_row}:B{row-1}',
            DataBarRule(start_type='num', start_value=0, end_type='num', end_value=50,
                       color="63C384", showValue=True, minLength=None, maxLength=None)
        )
        
        row += 2
    
    # Column widths
    ws1.column_dimensions['A'].width = 12
    ws1.column_dimensions['B'].width = 16
    ws1.column_dimensions['C'].width = 14
    ws1.column_dimensions['D'].width = 12
    ws1.column_dimensions['E'].width = 14
    ws1.column_dimensions['F'].width = 32
    
    # =========================================================================
    # CHART DATA SECTION (Right side of sheet, starting column H)
    # =========================================================================
    
    # Calculate averages for charts
    high_awareness_avg = sum(market_data['zones'][z].get('my_awareness', DEFAULT_AWARENESS) for z in ZONES) / len(ZONES)
    low_awareness_avg = high_awareness_avg  # Same default for now
    
    my_price_avg = sum(market_data['zones'][z].get('my_price', DEFAULT_PRICE) for z in ZONES) / len(ZONES)
    comp_price_avg = sum(market_data['zones'][z].get('comp_avg_price', DEFAULT_PRICE) for z in ZONES) / len(ZONES)
    my_attract_avg = sum(market_data['zones'][z].get('my_attractiveness', DEFAULT_ATTRACTIVENESS) for z in ZONES) / len(ZONES)
    comp_attract_avg = DEFAULT_ATTRACTIVENESS  # Competitor default
    
    # High segment averages
    high_avg_awareness = sum(market_data['by_segment']['High'][z].get('my_awareness', DEFAULT_AWARENESS) 
                             if 'my_awareness' in market_data['by_segment']['High'][z] 
                             else DEFAULT_AWARENESS for z in ZONES) / len(ZONES)
    high_avg_price_gap = sum(((market_data['zones'][z].get('my_price', DEFAULT_PRICE) - 
                               market_data['zones'][z].get('comp_avg_price', DEFAULT_PRICE)) / 
                              max(1, market_data['zones'][z].get('comp_avg_price', DEFAULT_PRICE)) * 100) 
                             for z in ZONES) / len(ZONES)
    high_avg_attract = my_attract_avg
    
    # Low segment (using same zone data)
    low_avg_awareness = high_avg_awareness
    low_avg_price_gap = high_avg_price_gap
    low_avg_attract = my_attract_avg
    
    # ----- Chart 1: Competitive Positioning Matrix Data -----
    ws1['H1'] = "COMPETITIVE POSITIONING"
    ws1['H1'].font = section_font
    
    # Data table for scatter chart
    ws1['H3'] = "Entity"
    ws1['I3'] = "Price"
    ws1['J3'] = "Attractiveness"
    for c in ['H', 'I', 'J']:
        ws1[f'{c}3'].font = header_font
        ws1[f'{c}3'].fill = header_fill
        ws1[f'{c}3'].border = thin_border
    
    ws1['H4'] = "My Product"
    ws1['I4'] = my_price_avg
    ws1['J4'] = my_attract_avg
    ws1['I4'].number_format = '$#,##0'
    for c in ['H', 'I', 'J']:
        ws1[f'{c}4'].border = thin_border
    
    ws1['H5'] = "Competitors"
    ws1['I5'] = comp_price_avg
    ws1['J5'] = comp_attract_avg
    ws1['I5'].number_format = '$#,##0'
    for c in ['H', 'I', 'J']:
        ws1[f'{c}5'].border = thin_border
    
    # Create Scatter Chart
    chart1 = ScatterChart()
    chart1.title = "Competitive Positioning Matrix"
    chart1.x_axis.title = "Price ($)"
    chart1.y_axis.title = "Attractiveness Score"
    chart1.style = 13
    chart1.height = 10
    chart1.width = 12
    
    # My Product series
    x_values1 = Reference(ws1, min_col=9, min_row=4, max_row=4)
    y_values1 = Reference(ws1, min_col=10, min_row=4, max_row=4)
    series1 = Series(y_values1, x_values1, title="My Product")
    series1.marker = Marker(symbol='circle', size=12)
    series1.graphicalProperties.solidFill = "4472C4"  # Blue
    chart1.series.append(series1)
    
    # Competitor series
    x_values2 = Reference(ws1, min_col=9, min_row=5, max_row=5)
    y_values2 = Reference(ws1, min_col=10, min_row=5, max_row=5)
    series2 = Series(y_values2, x_values2, title="Competitors")
    series2.marker = Marker(symbol='diamond', size=12)
    series2.graphicalProperties.solidFill = "ED7D31"  # Orange
    chart1.series.append(series2)
    
    ws1.add_chart(chart1, "H7")
    
    # ----- Chart 2: High vs Low Segment Gap Data -----
    ws1['H24'] = "HIGH vs LOW SEGMENT GAP"
    ws1['H24'].font = section_font
    
    # Data table for bar chart
    ws1['H26'] = "Metric"
    ws1['I26'] = "High Segment"
    ws1['J26'] = "Low Segment"
    for c in ['H', 'I', 'J']:
        ws1[f'{c}26'].font = header_font
        ws1[f'{c}26'].fill = header_fill
        ws1[f'{c}26'].border = thin_border
    
    ws1['H27'] = "Awareness"
    ws1['I27'] = high_avg_awareness
    ws1['J27'] = low_avg_awareness
    for c in ['H', 'I', 'J']:
        ws1[f'{c}27'].border = thin_border
    
    ws1['H28'] = "Price Competitiveness"
    ws1['I28'] = 100 - abs(high_avg_price_gap)  # Higher = better
    ws1['J28'] = 100 - abs(low_avg_price_gap)
    for c in ['H', 'I', 'J']:
        ws1[f'{c}28'].border = thin_border
    
    ws1['H29'] = "Attractiveness"
    ws1['I29'] = high_avg_attract
    ws1['J29'] = low_avg_attract
    for c in ['H', 'I', 'J']:
        ws1[f'{c}29'].border = thin_border
    
    # Create Clustered Bar Chart
    chart2 = BarChart()
    chart2.type = "col"
    chart2.grouping = "clustered"
    chart2.title = "High vs Low Segment Comparison"
    chart2.style = 13
    chart2.height = 10
    chart2.width = 12
    
    # Data references
    categories = Reference(ws1, min_col=8, min_row=27, max_row=29)
    data = Reference(ws1, min_col=9, min_row=26, max_col=10, max_row=29)
    
    chart2.add_data(data, titles_from_data=True)
    chart2.set_categories(categories)
    chart2.shape = 4
    
    ws1.add_chart(chart2, "H31")
    
    # ----- Traffic Light Conditional Formatting -----
    # Apply to awareness data in both High and Low segment sections
    # High segment awareness is in column C, rows 5-9 (approx)
    # Low segment awareness is in column C, rows 13-17 (approx)
    
    # Create icon set rule for awareness columns
    icon_rule = IconSetRule(
        icon_style='3TrafficLights1',
        type='num',
        values=[0, 40, 70],  # Red < 40, Yellow 40-70, Green > 70
        showValue=True,
        reverse=False
    )
    
    # Apply to High segment (rows 5-9, column C - Awareness Gap) 
    ws1.conditional_formatting.add('C5:C9', icon_rule)
    # Apply to Low segment (rows 13-17, column C)
    ws1.conditional_formatting.add('C13:C17', icon_rule)
    
    # Price Gap red text formatting (column D if > 10%)
    price_gap_rule = FormulaRule(
        formula=['D5>0.1'],
        font=Font(bold=True, color="9C0006")
    )
    ws1.conditional_formatting.add('D5:D9', price_gap_rule)
    ws1.conditional_formatting.add('D13:D17', price_gap_rule)
    
    # Additional column widths for chart data
    ws1.column_dimensions['H'].width = 18
    ws1.column_dimensions['I'].width = 14
    ws1.column_dimensions['J'].width = 14

    
    # =========================================================================
    # TAB 2: INNOVATION_LAB
    # =========================================================================
    ws2 = wb.create_sheet("INNOVATION_LAB")
    
    ws2['A1'] = "INNOVATION LAB - Feature Selection"
    ws2['A1'].font = title_font
    
    ws2['A2'] = "Note: Innovations increase Attractiveness. Required for High Segment Allocation."
    ws2['A2'].font = Font(italic=True, color="666666")
    
    # Headers
    innov_headers = ['Feature Name', 'Decision (1=Yes)', 'Est. Cost ($)']
    for col, header in enumerate(innov_headers, start=1):
        cell = ws2.cell(row=4, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    
    # Dynamic feature list
    row = 5
    for feature in innovation_features:
        ws2.cell(row=row, column=1, value=feature).border = thin_border
        
        cell = ws2.cell(row=row, column=2, value=0)
        cell.border = thin_border
        cell.fill = input_fill
        cell.alignment = Alignment(horizontal='center')
        
        cell = ws2.cell(row=row, column=3, value=10000)
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '$#,##0'
        
        row += 1
    
    # Total innovation cost
    row += 1
    ws2.cell(row=row, column=1, value="TOTAL INNOVATION COST").font = Font(bold=True)
    cell = ws2.cell(row=row, column=3, value=f'=SUMPRODUCT(B5:B{row-2},C5:C{row-2})')
    cell.fill = calc_fill
    cell.font = Font(bold=True)
    cell.number_format = '$#,##0'
    
    ws2.column_dimensions['A'].width = 35
    ws2.column_dimensions['B'].width = 18
    ws2.column_dimensions['C'].width = 15
    
    innov_cost_cell = f'C{row}'
    
    # =========================================================================
    # TAB 3: STRATEGY_COCKPIT
    # =========================================================================
    ws3 = wb.create_sheet("STRATEGY_COCKPIT")
    
    ws3['A1'] = "HOW TO USE: Adjust Yellow cells. Check Profit Projection. Go to UPLOAD_READY tabs to copy decisions."
    ws3['A1'].font = Font(italic=True, color="666666")
    
    # Section A: Global Allocations
    ws3['A3'] = "SECTION A: GLOBAL ALLOCATIONS"
    ws3['A3'].font = section_font
    
    ws3.cell(row=5, column=1, value="TV Budget ($)").border = thin_border
    cell = ws3.cell(row=5, column=2, value=marketing_template['tv_budget'])
    cell.border = thin_border
    cell.fill = input_fill
    cell.number_format = '$#,##0'
    ws3['C5'] = "Primary Driver: High Segment Awareness"
    ws3['C5'].font = Font(italic=True, color="666666")
    
    ws3.cell(row=6, column=1, value="Brand Focus (0-100)").border = thin_border
    cell = ws3.cell(row=6, column=2, value=marketing_template['brand_focus'])
    cell.border = thin_border
    cell.fill = input_fill
    ws3['C6'] = "0=Awareness focus, 100=Attributes focus"
    ws3['C6'].font = Font(italic=True, color="666666")
    
    # Section B: Zonal Allocations
    ws3['A9'] = "SECTION B: ZONAL ALLOCATIONS"
    ws3['A9'].font = section_font
    
    zonal_headers = ['Zone', 'Last Sales', 'Stockout?', 'Target Demand', 'Radio Budget',
                     'Salespeople', 'Price', 'Payment', 'Est. Revenue', 'Mkt Cost', 'Contribution']
    
    for col, header in enumerate(zonal_headers, start=1):
        cell = ws3.cell(row=11, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
    
    row = 12
    for zone in ZONES:
        zone_sales = sales_data['by_zone'].get(zone, {})
        last_sales = zone_sales.get('units', 1000)
        is_stockout = inventory_data['is_stockout']
        
        ws3.cell(row=row, column=1, value=zone).border = thin_border
        
        # Reference data (gray)
        cell = ws3.cell(row=row, column=2, value=last_sales)
        cell.border = thin_border
        cell.fill = ref_fill
        cell.number_format = '#,##0'
        
        cell = ws3.cell(row=row, column=3, value="TRUE DEMAND HIGHER" if is_stockout else "OK")
        cell.border = thin_border
        if is_stockout:
            cell.fill = red_fill
            cell.font = Font(bold=True, color="9C0006")
        else:
            cell.fill = ref_fill
        
        # Inputs (yellow)
        target_demand = int(last_sales * 1.1) if is_stockout else last_sales
        cell = ws3.cell(row=row, column=4, value=target_demand)
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '#,##0'
        
        cell = ws3.cell(row=row, column=5, value=marketing_template['radio_budgets'].get(zone, 100))
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '$#,##0'
        
        cell = ws3.cell(row=row, column=6, value=marketing_template['salespeople'].get(zone, 10))
        cell.border = thin_border
        cell.fill = input_fill
        
        cell = ws3.cell(row=row, column=7, value=marketing_template['prices'].get(zone, 68))
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '$#,##0'
        
        cell = ws3.cell(row=row, column=8, value=marketing_template['payment_terms'].get(zone, 'B'))
        cell.border = thin_border
        cell.fill = input_fill
        cell.alignment = Alignment(horizontal='center')
        
        # Calculated (formulas)
        # Est. Revenue = Demand * Price
        cell = ws3.cell(row=row, column=9, value=f'=D{row}*G{row}')
        cell.border = thin_border
        cell.fill = calc_fill
        cell.number_format = '$#,##0'
        
        # Marketing Cost = TV/5 + Radio + Salespeople*Salary + Innovation/5
        cell = ws3.cell(row=row, column=10, value=f'=($B$5/5)+E{row}+(F{row}*{DEFAULT_SALESPEOPLE_SALARY})+(INNOVATION_LAB!{innov_cost_cell}/5)')
        cell.border = thin_border
        cell.fill = calc_fill
        cell.number_format = '$#,##0'
        
        # Contribution = Revenue - MktCost - (Demand * COGS)
        cell = ws3.cell(row=row, column=11, value=f'=I{row}-J{row}-(D{row}*{DEFAULT_COGS})')
        cell.border = thin_border
        cell.fill = output_fill
        cell.font = Font(bold=True)
        cell.number_format = '$#,##0'
        
        row += 1
    
    # Totals
    ws3.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
    ws3.cell(row=row, column=4, value=f'=SUM(D12:D{row-1})').fill = input_fill
    ws3.cell(row=row, column=9, value=f'=SUM(I12:I{row-1})').fill = calc_fill
    ws3.cell(row=row, column=10, value=f'=SUM(J12:J{row-1})').fill = calc_fill
    cell = ws3.cell(row=row, column=11, value=f'=SUM(K12:K{row-1})')
    cell.fill = output_fill
    cell.font = Font(bold=True)
    
    # Column widths
    ws3.column_dimensions['A'].width = 12
    for col in range(2, 12):
        ws3.column_dimensions[get_column_letter(col)].width = 14
    
    # =========================================================================
    # TAB 4: UPLOAD_READY_MARKETING
    # =========================================================================
    ws4 = wb.create_sheet("UPLOAD_READY_MARKETING")
    
    ws4['A1'] = "MARKETING DECISIONS - ExSim Upload Format"
    ws4['A1'].font = title_font
    ws4['A2'] = "Copy these values to ExSim Marketing upload"
    ws4['A2'].font = Font(italic=True, color="666666")
    
    # Recreate the side-by-side layout
    # Marketing Campaigns (cols A-E)
    ws4['A4'] = "Marketing Campaigns"
    ws4['A4'].font = section_font
    
    camp_headers = ['Brand', 'Zone', 'Channel', 'Amount', 'Brand Focus']
    for col, h in enumerate(camp_headers, start=1):
        cell = ws4.cell(row=5, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    
    # TV row
    ws4.cell(row=6, column=1, value='A').border = thin_border
    ws4.cell(row=6, column=2, value='All').border = thin_border
    ws4.cell(row=6, column=3, value='TV').border = thin_border
    ws4.cell(row=6, column=4, value='=STRATEGY_COCKPIT!B5').border = thin_border
    ws4.cell(row=6, column=5, value='=STRATEGY_COCKPIT!B6').border = thin_border
    
    # Radio rows
    row = 7
    for zone_idx, zone in enumerate(ZONES):
        ws4.cell(row=row, column=1, value='A').border = thin_border
        ws4.cell(row=row, column=2, value=zone).border = thin_border
        ws4.cell(row=row, column=3, value='Radio').border = thin_border
        ws4.cell(row=row, column=4, value=f'=STRATEGY_COCKPIT!E{12+zone_idx}').border = thin_border
        ws4.cell(row=row, column=5, value='=STRATEGY_COCKPIT!B6').border = thin_border
        row += 1
    
    # Demand section (cols G-H)
    ws4['G4'] = "Demand"
    ws4['G4'].font = section_font
    
    ws4.cell(row=5, column=7, value='Zone').font = header_font
    ws4.cell(row=5, column=7).fill = header_fill
    ws4.cell(row=5, column=8, value='Demand').font = header_font
    ws4.cell(row=5, column=8).fill = header_fill
    
    for zone_idx, zone in enumerate(ZONES):
        ws4.cell(row=6+zone_idx, column=7, value=zone).border = thin_border
        ws4.cell(row=6+zone_idx, column=8, value=f'=STRATEGY_COCKPIT!D{12+zone_idx}').border = thin_border
    
    # Pricing section (cols J-L)
    ws4['J4'] = "Pricing Strategy"
    ws4['J4'].font = section_font
    
    ws4.cell(row=5, column=10, value='Zone').font = header_font
    ws4.cell(row=5, column=10).fill = header_fill
    ws4.cell(row=5, column=11, value='Brand').font = header_font
    ws4.cell(row=5, column=11).fill = header_fill
    ws4.cell(row=5, column=12, value='Price').font = header_font
    ws4.cell(row=5, column=12).fill = header_fill
    
    for zone_idx, zone in enumerate(ZONES):
        ws4.cell(row=6+zone_idx, column=10, value=zone).border = thin_border
        ws4.cell(row=6+zone_idx, column=11, value='A').border = thin_border
        ws4.cell(row=6+zone_idx, column=12, value=f'=STRATEGY_COCKPIT!G{12+zone_idx}').border = thin_border
    
    # Channels section (cols N-P)
    ws4['N4'] = "Channels"
    ws4['N4'].font = section_font
    
    ws4.cell(row=5, column=14, value='Zone').font = header_font
    ws4.cell(row=5, column=14).fill = header_fill
    ws4.cell(row=5, column=15, value='Payment').font = header_font
    ws4.cell(row=5, column=15).fill = header_fill
    ws4.cell(row=5, column=16, value='Salespeople').font = header_font
    ws4.cell(row=5, column=16).fill = header_fill
    
    for zone_idx, zone in enumerate(ZONES):
        ws4.cell(row=6+zone_idx, column=14, value=zone).border = thin_border
        ws4.cell(row=6+zone_idx, column=15, value=f'=STRATEGY_COCKPIT!H{12+zone_idx}').border = thin_border
        ws4.cell(row=6+zone_idx, column=16, value=f'=STRATEGY_COCKPIT!F{12+zone_idx}').border = thin_border
    
    # =========================================================================
    # TAB 5: UPLOAD_READY_INNOVATION
    # =========================================================================
    ws5 = wb.create_sheet("UPLOAD_READY_INNOVATION")
    
    ws5['A1'] = "INNOVATION DECISIONS - ExSim Upload Format"
    ws5['A1'].font = title_font
    ws5['A2'] = "Copy these values to ExSim Innovation upload"
    ws5['A2'].font = Font(italic=True, color="666666")
    
    # Headers
    ws5.cell(row=4, column=1, value='Brand').font = header_font
    ws5.cell(row=4, column=1).fill = header_fill
    ws5.cell(row=4, column=2, value='Improvement').font = header_font
    ws5.cell(row=4, column=2).fill = header_fill
    ws5.cell(row=4, column=3, value='Value').font = header_font
    ws5.cell(row=4, column=3).fill = header_fill
    
    for i, feature in enumerate(innovation_features):
        ws5.cell(row=5+i, column=1, value='A').border = thin_border
        ws5.cell(row=5+i, column=2, value=feature).border = thin_border
        # Link to INNOVATION_LAB decision
        ws5.cell(row=5+i, column=3, value=f'=INNOVATION_LAB!B{5+i}').border = thin_border
    
    ws5.column_dimensions['A'].width = 10
    ws5.column_dimensions['B'].width = 35
    ws5.column_dimensions['C'].width = 10
    
    # Save
    wb.save(OUTPUT_FILE)
    print(f"[SUCCESS] Created '{OUTPUT_FILE}'")


def main():
    """Main function."""
    print("ExSim CMO Complete Dashboard Generator")
    print("=" * 50)
    
    print("\n[*] Loading data files...")
    
    # Market Report
    market_path = DATA_FOLDER / "market-report.xlsx"
    if market_path.exists():
        market_data = load_market_report(market_path)
        print(f"  [OK] Loaded market report with segment data")
    else:
        market_data = load_market_report(None)
        print("  [!] Using default market data")
    
    # Innovation Template
    innov_path = DATA_FOLDER / "Marketing Innovation Decisions.xlsx"
    if innov_path.exists():
        innovation_features = load_innovation_template(innov_path)
        print(f"  [OK] Loaded {len(innovation_features)} innovation features")
    else:
        innovation_features = load_innovation_template(None)
        print("  [!] Using default innovation features")
    
    # Marketing Template
    mkt_path = DATA_FOLDER / "Marketing Decisions.xlsx"
    if mkt_path.exists():
        marketing_template = load_marketing_template(mkt_path)
        print(f"  [OK] Loaded marketing template")
    else:
        marketing_template = load_marketing_template(None)
        print("  [!] Using default marketing template")
    
    # Sales Data
    sales_path = DATA_FOLDER / "sales_admin_expenses.xlsx"
    if sales_path.exists():
        sales_data = load_sales_data(sales_path)
        print(f"  [OK] Loaded sales data")
    else:
        sales_data = load_sales_data(None)
        print("  [!] Using default sales data")
    
    # Inventory
    inv_path = DATA_FOLDER / "finished_goods_inventory.xlsx"
    if inv_path.exists():
        inventory_data = load_inventory_data(inv_path)
        stockout_status = "STOCKOUT DETECTED" if inventory_data['is_stockout'] else "OK"
        print(f"  [OK] Loaded inventory: {stockout_status}")
    else:
        inventory_data = load_inventory_data(None)
        print("  [!] Using default inventory data")
    
    print("\n[*] Generating CMO Dashboard...")
    
    create_complete_dashboard(market_data, innovation_features, marketing_template,
                              sales_data, inventory_data)
    
    print("\nSheets created:")
    print("  * SEGMENT_PULSE (High/Low Segment Analysis)")
    print("  * INNOVATION_LAB (Feature Selection)")
    print("  * STRATEGY_COCKPIT (4 Ps Decisions + ROI)")
    print("  * UPLOAD_READY_MARKETING (ExSim Format)")
    print("  * UPLOAD_READY_INNOVATION (ExSim Format)")


if __name__ == "__main__":
    main()
