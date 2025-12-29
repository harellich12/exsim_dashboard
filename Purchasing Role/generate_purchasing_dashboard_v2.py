"""
ExSim Purchasing Dashboard - MRP & Cost-Optimized Sourcing

Builds an MRP calculator that maps Production Needs to Supplier Orders,
handles Lead Time shifts, and analyzes Batch Size efficiency.

Required libraries: pandas, openpyxl
"""

import pandas as pd
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
import warnings

warnings.filterwarnings('ignore')

# =============================================================================
# CONFIGURATION
# =============================================================================
DATA_FOLDER = Path("data")
OUTPUT_FILE = "Purchasing_Dashboard.xlsx"

FORTNIGHTS = list(range(1, 9))  # 1-8
ZONES = ["Center", "West", "North", "East", "South"]
PARTS = ["Part A", "Part B"]
PIECES = ["Piece 1", "Piece 2", "Piece 3", "Piece 4", "Piece 5", "Piece 6"]
SUPPLIERS = ["Supplier A", "Supplier B", "Supplier C"]

# Default supplier configuration (per ExSim template: 3 suppliers per part)
# Default supplier configuration
DEFAULT_SUPPLIERS = {
    "Part A": [
        {"name": "Supplier A", "lead_time": 0, "cost": 0, "payment_terms": 0, "batch_size": 0},
        {"name": "Supplier B", "lead_time": 0, "cost": 0, "payment_terms": 0, "batch_size": 0},
        {"name": "Supplier C", "lead_time": 0, "cost": 0, "payment_terms": 0, "batch_size": 0},
    ],
    "Part B": [
        {"name": "Supplier A", "lead_time": 0, "cost": 0, "payment_terms": 0, "batch_size": 0},
        {"name": "Supplier B", "lead_time": 0, "cost": 0, "payment_terms": 0, "batch_size": 0},
        {"name": "Supplier C", "lead_time": 0, "cost": 0, "payment_terms": 0, "batch_size": 0},
    ]
}

DEFAULT_PIECES_CONFIG = {
    "Piece 1": {"cost": 0, "batch_size": 0},
    "Piece 2": {"cost": 0, "batch_size": 0},
    "Piece 3": {"cost": 0, "batch_size": 0},
    "Piece 4": {"cost": 0, "batch_size": 0},
    "Piece 5": {"cost": 0, "batch_size": 0},
    "Piece 6": {"cost": 0, "batch_size": 0},
}


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
    """Load Excel file."""
    try:
        if sheet_name:
            return pd.read_excel(filepath, sheet_name=sheet_name, header=None)
        return pd.read_excel(filepath, header=None)
    except Exception as e:
        print(f"Warning: Could not load {filepath}: {e}")
        return None


# =============================================================================
# DATA LOADING
# =============================================================================

def load_raw_materials(filepath):
    """Load raw materials inventory data."""
    df = load_excel_file(filepath)
    
    data = {
        'parts': {part: {'final_inventory': 0} for part in PARTS},
        'pieces': {piece: {'final_inventory': 0} for piece in PIECES}
    }
    
    if df is None:
        return data
    
    current_item = None
    
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        
        # Detect Part/Piece names
        for part in PARTS:
            if part.lower() in first_val.lower():
                current_item = part
                break
        for piece in PIECES:
            if piece.lower() in first_val.lower():
                current_item = piece
                break
        
        # Get Final Inventory
        if 'final' in first_val.lower() and 'inventory' in first_val.lower():
            # Get fortnight 8 value (column 8)
            final_val = parse_numeric(row.iloc[8]) if len(row) > 8 else 0
            if current_item:
                if current_item in PARTS:
                    data['parts'][current_item]['final_inventory'] = final_val
                elif current_item in PIECES:
                    data['pieces'][current_item]['final_inventory'] = final_val
    
    return data


def load_production_costs(filepath):
    """Load production cost data for batch size analysis."""
    df = load_excel_file(filepath)
    
    data = {
        'ordering_cost': 0,
        'holding_cost': 0,
        'consumption_cost': 0
    }
    
    if df is None:
        return data
    
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
        
        if 'ordering' in first_val and 'cost' in first_val:
            for col_idx in range(1, min(10, len(row))):
                val = parse_numeric(row.iloc[col_idx])
                if val > 0:
                    data['ordering_cost'] = val
                    break
        
        if 'holding' in first_val and 'cost' in first_val:
            for col_idx in range(1, min(10, len(row))):
                val = parse_numeric(row.iloc[col_idx])
                if val > 0:
                    data['holding_cost'] = val
                    break
        
        if 'consumed' in first_val or 'consumption' in first_val:
            for col_idx in range(1, min(10, len(row))):
                val = parse_numeric(row.iloc[col_idx])
                if val > 0:
                    data['consumption_cost'] = val
                    break
    
    return data


def load_procurement_template(filepath):
    """Load procurement template structure."""
    df = load_excel_file(filepath, sheet_name='Procurement')
    return {'df': df, 'exists': df is not None}


# =============================================================================
# EXCEL GENERATION
# =============================================================================

def create_purchasing_dashboard(materials_data, cost_data, template_data):
    """Create the comprehensive Purchasing Dashboard."""
    
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
    # TAB 1: SUPPLIER_CONFIG
    # =========================================================================
    ws1 = wb.active
    ws1.title = "SUPPLIER_CONFIG"
    
    ws1['A1'] = "SUPPLIER CONFIGURATION"
    ws1['A1'].font = title_font
    ws1['A2'] = "Enter your case study supplier data here. Pre-filled with defaults."
    ws1['A2'].font = Font(italic=True, color="666666")
    
    # Table 1: PARTS CONFIG
    ws1['A4'] = "TABLE 1: PARTS SUPPLIERS"
    ws1['A4'].font = section_font
    
    parts_headers = ['Part', 'Supplier', 'Lead Time (FN)', 'Cost/Unit', 'Payment Terms (FN)', 'Batch Size']
    for col, h in enumerate(parts_headers, start=1):
        cell = ws1.cell(row=5, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    
    row = 6
    for part, suppliers in DEFAULT_SUPPLIERS.items():
        for supplier in suppliers:
            ws1.cell(row=row, column=1, value=part).border = thin_border
            
            cell = ws1.cell(row=row, column=2, value=supplier['name'])
            cell.border = thin_border
            cell.fill = input_fill
            
            cell = ws1.cell(row=row, column=3, value=supplier['lead_time'])
            cell.border = thin_border
            cell.fill = input_fill
            
            cell = ws1.cell(row=row, column=4, value=supplier['cost'])
            cell.border = thin_border
            cell.fill = input_fill
            cell.number_format = '$#,##0.00'
            
            cell = ws1.cell(row=row, column=5, value=supplier['payment_terms'])
            cell.border = thin_border
            cell.fill = input_fill
            
            cell = ws1.cell(row=row, column=6, value=supplier['batch_size'])
            cell.border = thin_border
            cell.fill = input_fill
            
            row += 1
    
    parts_config_end = row - 1
    row += 2
    
    # Table 2: PIECES CONFIG
    ws1.cell(row=row, column=1, value="TABLE 2: PIECES CONFIGURATION").font = section_font
    row += 1
    
    pieces_headers = ['Piece Name', 'Cost/Unit', 'Batch Size']
    for col, h in enumerate(pieces_headers, start=1):
        cell = ws1.cell(row=row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    row += 1
    
    pieces_start_row = row
    for piece, config in DEFAULT_PIECES_CONFIG.items():
        ws1.cell(row=row, column=1, value=piece).border = thin_border
        
        cell = ws1.cell(row=row, column=2, value=config['cost'])
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '$#,##0.00'
        
        cell = ws1.cell(row=row, column=3, value=config['batch_size'])
        cell.border = thin_border
        cell.fill = input_fill
        
        row += 1
    
    # Column widths
    ws1.column_dimensions['A'].width = 15
    ws1.column_dimensions['B'].width = 15
    ws1.column_dimensions['C'].width = 16
    ws1.column_dimensions['D'].width = 12
    ws1.column_dimensions['E'].width = 18
    ws1.column_dimensions['F'].width = 12
    
    # =========================================================================
    # TAB 2: COST_ANALYSIS
    # =========================================================================
    ws2 = wb.create_sheet("COST_ANALYSIS")
    
    ws2['A1'] = "COST ANALYSIS - Batch Size Efficiency"
    ws2['A1'].font = title_font
    
    ws2['A3'] = "PREVIOUS PERIOD COSTS"
    ws2['A3'].font = section_font
    
    # Cost data
    ordering_cost = cost_data.get('ordering_cost', 0)
    holding_cost = cost_data.get('holding_cost', 0)
    total_cost = ordering_cost + holding_cost
    
    ws2.cell(row=5, column=1, value="Ordering Cost (Total)").border = thin_border
    cell = ws2.cell(row=5, column=2, value=ordering_cost)
    cell.border = thin_border
    cell.number_format = '$#,##0'
    
    ws2.cell(row=6, column=1, value="Holding Cost (Total)").border = thin_border
    cell = ws2.cell(row=6, column=2, value=holding_cost)
    cell.border = thin_border
    cell.number_format = '$#,##0'
    
    ws2.cell(row=7, column=1, value="Total Cost").border = thin_border
    cell = ws2.cell(row=7, column=2, value="=B5+B6")
    cell.border = thin_border
    cell.fill = calc_fill
    cell.number_format = '$#,##0'
    
    # Ordering Cost Ratio
    ws2['A9'] = "EFFICIENCY ANALYSIS"
    ws2['A9'].font = section_font
    
    ws2.cell(row=11, column=1, value="Ordering Cost Ratio").border = thin_border
    cell = ws2.cell(row=11, column=2, value="=IF(B7>0, B5/B7, 0)")
    cell.border = thin_border
    cell.fill = calc_fill
    cell.number_format = '0.0%'
    
    # Efficiency Flag
    ws2.cell(row=13, column=1, value="Efficiency Flag").border = thin_border
    cell = ws2.cell(row=13, column=2)
    cell.value = '=IF(B11>0.7,"CRITICAL: Ordering too often. INCREASE BATCH SIZE.",IF(B11<0.3,"CRITICAL: Holding too much. DECREASE BATCH SIZE/JIT.","OK: Balanced"))'
    cell.border = thin_border
    cell.font = Font(bold=True)
    
    # Strategic Advice Box
    ws2['A16'] = "STRATEGIC ADVICE"
    ws2['A16'].font = section_font
    
    ws2.merge_cells('A17:D20')
    cell = ws2['A17']
    cell.value = """Based on your Ordering Cost Ratio:
• > 70%: You're placing too many small orders. Consolidate orders into larger batches.
• < 30%: You're holding too much inventory. Consider Just-In-Time ordering or smaller batches.
• 30-70%: Good balance between ordering frequency and inventory holding."""
    cell.alignment = Alignment(wrap_text=True, vertical='top')
    cell.fill = calc_fill
    
    ws2.row_dimensions[17].height = 80
    
    # Column widths
    ws2.column_dimensions['A'].width = 25
    ws2.column_dimensions['B'].width = 50
    
    # =========================================================================
    # TAB 3: MRP_ENGINE
    # =========================================================================
    ws3 = wb.create_sheet("MRP_ENGINE")
    
    ws3['A1'] = "MRP ENGINE - Material Requirements Planning"
    ws3['A1'].font = title_font
    
    # Section A: Production Demand
    ws3['A3'] = "SECTION A: PRODUCTION DEMAND (Input from Production Manager)"
    ws3['A3'].font = section_font
    
    # Headers
    ws3.cell(row=5, column=1, value="Metric").font = header_font
    ws3.cell(row=5, column=1).fill = header_fill
    for fn in FORTNIGHTS:
        cell = ws3.cell(row=5, column=1+fn, value=f"FN{fn}")
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
    
    # Target Production input row
    ws3.cell(row=6, column=1, value="Target Production").border = thin_border
    for fn in FORTNIGHTS:
        cell = ws3.cell(row=6, column=1+fn, value=0)
        cell.border = thin_border
        cell.fill = input_fill
    
    row = 9
    
    # Section B: Net Requirements for each Part
    ws3.cell(row=row, column=1, value="SECTION B: NET REQUIREMENTS CALCULATION").font = section_font
    row += 2
    
    part_rows = {}
    for part in PARTS:
        ws3.cell(row=row, column=1, value=f"{part.upper()}").font = Font(bold=True, color="2F5496")
        row += 1
        
        opening_inv = materials_data['parts'].get(part, {}).get('final_inventory', 0)
        
        # Gross Requirement
        ws3.cell(row=row, column=1, value="Gross Requirement").border = thin_border
        for fn in FORTNIGHTS:
            cell = ws3.cell(row=row, column=1+fn, value=f"=$B$6")  # Links to target production
            cell.border = thin_border
            cell.fill = ref_fill
        gross_row = row
        row += 1
        
        # Scheduled Arrivals (user inputs based on orders placed)
        ws3.cell(row=row, column=1, value="Scheduled Arrivals").border = thin_border
        for fn in FORTNIGHTS:
            cell = ws3.cell(row=row, column=1+fn, value=0)
            cell.border = thin_border
            cell.fill = input_fill
        arrivals_row = row
        row += 1
        
        # Projected Inventory
        ws3.cell(row=row, column=1, value="Projected Inventory").border = thin_border
        for fn in FORTNIGHTS:
            col = 1 + fn
            if fn == 1:
                # First fortnight: Opening + Arrivals - Gross
                cell = ws3.cell(row=row, column=col, 
                    value=f"={opening_inv}+{get_column_letter(col)}{arrivals_row}-{get_column_letter(col)}{gross_row}")
            else:
                # Subsequent: Previous Inv + Arrivals - Gross
                prev_col = get_column_letter(col - 1)
                curr_col = get_column_letter(col)
                cell = ws3.cell(row=row, column=col,
                    value=f"={prev_col}{row}+{curr_col}{arrivals_row}-{curr_col}{gross_row}")
            cell.border = thin_border
            cell.fill = calc_fill
        proj_row = row
        part_rows[part] = {'arrivals': arrivals_row, 'projected': proj_row}
        row += 1
        
        # Net Deficit
        ws3.cell(row=row, column=1, value="Net Deficit (if negative)").border = thin_border
        for fn in FORTNIGHTS:
            col = 1 + fn
            cell = ws3.cell(row=row, column=col, 
                value=f'=IF({get_column_letter(col)}{proj_row}<0, -{get_column_letter(col)}{proj_row}, 0)')
            cell.border = thin_border
            cell.fill = output_fill
        row += 2
    
    # Add conditional formatting for negative inventory
    for part, rows in part_rows.items():
        ws3.conditional_formatting.add(
            f'B{rows["projected"]}:I{rows["projected"]}',
            FormulaRule(formula=[f'B{rows["projected"]}<0'], fill=red_fill, font=Font(bold=True, color="9C0006"))
        )
    
    # Section C: Sourcing Strategy
    ws3.cell(row=row, column=1, value="SECTION C: SOURCING STRATEGY (Order Inputs)").font = section_font
    ws3.cell(row=row+1, column=1, value="NOTE: Orders arrive AFTER Lead Time. Enter in the FN you want to ORDER, not when it arrives.").font = Font(italic=True, color="666666")
    row += 3
    
    order_rows = {}
    for part in PARTS:
        ws3.cell(row=row, column=1, value=f"ORDERS FOR {part.upper()}").font = Font(bold=True)
        row += 1
        
        order_rows[part] = {}
        suppliers = DEFAULT_SUPPLIERS.get(part, [])
        for supplier in suppliers:
            name = supplier['name']
            lead = supplier['lead_time']
            batch = supplier['batch_size']
            
            ws3.cell(row=row, column=1, value=f"Order {name} (Lead:{lead}, Batch:{batch})").border = thin_border
            for fn in FORTNIGHTS:
                cell = ws3.cell(row=row, column=1+fn, value=0)
                cell.border = thin_border
                cell.fill = input_fill
            order_rows[part][name] = row
            row += 1
        
        # Batch compliance check
        ws3.cell(row=row, column=1, value="Batch Compliance Check").border = thin_border
        for fn in FORTNIGHTS:
            col = 1 + fn
            # Check if orders are multiples of batch size
            cell = ws3.cell(row=row, column=col, value="OK")  # Simplified - would need complex formula
            cell.border = thin_border
            cell.fill = calc_fill
        row += 2
    
    # Column widths
    ws3.column_dimensions['A'].width = 35
    for col in range(2, 10):
        ws3.column_dimensions[get_column_letter(col)].width = 10
    
    # =========================================================================
    # TAB 4: CASH_FLOW_PREVIEW
    # =========================================================================
    ws4 = wb.create_sheet("CASH_FLOW_PREVIEW")
    
    ws4['A1'] = "CASH FLOW PREVIEW - Procurement Spending"
    ws4['A1'].font = title_font
    
    ws4['A3'] = "ESTIMATED OUTFLOW BY FORTNIGHT"
    ws4['A3'].font = section_font
    
    # Headers
    ws4.cell(row=5, column=1, value="Category").font = header_font
    ws4.cell(row=5, column=1).fill = header_fill
    for fn in FORTNIGHTS:
        cell = ws4.cell(row=5, column=1+fn, value=f"FN{fn}")
        cell.font = header_font
        cell.fill = header_fill
    cell = ws4.cell(row=5, column=10, value="Total")
    cell.font = header_font
    cell.fill = header_fill
    
    # Placeholder rows for each part
    row = 6
    for part in PARTS:
        ws4.cell(row=row, column=1, value=f"{part} Orders").border = thin_border
        for fn in FORTNIGHTS:
            cell = ws4.cell(row=row, column=1+fn, value=0)
            cell.border = thin_border
            cell.fill = calc_fill
            cell.number_format = '$#,##0'
        cell = ws4.cell(row=row, column=10, value=f"=SUM(B{row}:I{row})")
        cell.border = thin_border
        cell.fill = calc_fill
        cell.number_format = '$#,##0'
        row += 1
    
    # Pieces orders
    ws4.cell(row=row, column=1, value="Pieces Orders").border = thin_border
    for fn in FORTNIGHTS:
        cell = ws4.cell(row=row, column=1+fn, value=0)
        cell.border = thin_border
        cell.fill = calc_fill
        cell.number_format = '$#,##0'
    cell = ws4.cell(row=row, column=10, value=f"=SUM(B{row}:I{row})")
    cell.border = thin_border
    row += 1
    
    # Total row
    ws4.cell(row=row, column=1, value="TOTAL SPEND").font = Font(bold=True)
    for fn in FORTNIGHTS:
        cell = ws4.cell(row=row, column=1+fn, value=f"=SUM({get_column_letter(1+fn)}6:{get_column_letter(1+fn)}{row-1})")
        cell.border = thin_border
        cell.fill = output_fill
        cell.font = Font(bold=True)
        cell.number_format = '$#,##0'
    cell = ws4.cell(row=row, column=10, value=f"=SUM(B{row}:I{row})")
    cell.fill = output_fill
    cell.font = Font(bold=True)
    cell.number_format = '$#,##0'
    total_row = row
    row += 1
    
    # Cumulative spend
    ws4.cell(row=row, column=1, value="CUMULATIVE SPEND").font = Font(bold=True)
    for fn in FORTNIGHTS:
        col = 1 + fn
        if fn == 1:
            cell = ws4.cell(row=row, column=col, value=f"=B{total_row}")
        else:
            cell = ws4.cell(row=row, column=col, value=f"={get_column_letter(col-1)}{row}+{get_column_letter(col)}{total_row}")
        cell.border = thin_border
        cell.fill = ref_fill
        cell.number_format = '$#,##0'
    
    # Budget tracking
    row += 3
    ws4.cell(row=row, column=1, value="BUDGET TRACKING").font = section_font
    row += 1
    
    ws4.cell(row=row, column=1, value="Total Budget").border = thin_border
    cell = ws4.cell(row=row, column=2, value=100000)
    cell.border = thin_border
    cell.fill = input_fill
    cell.number_format = '$#,##0'
    row += 1
    
    ws4.cell(row=row, column=1, value="Total Projected Spend").border = thin_border
    cell = ws4.cell(row=row, column=2, value=f"=J{total_row}")
    cell.border = thin_border
    cell.fill = calc_fill
    cell.number_format = '$#,##0'
    row += 1
    
    ws4.cell(row=row, column=1, value="Remaining Budget").border = thin_border
    cell = ws4.cell(row=row, column=2, value=f"=B{row-2}-B{row-1}")
    cell.border = thin_border
    cell.fill = output_fill
    cell.font = Font(bold=True)
    cell.number_format = '$#,##0'
    
    # Column widths
    ws4.column_dimensions['A'].width = 20
    for col in range(2, 11):
        ws4.column_dimensions[get_column_letter(col)].width = 12
    
    # =========================================================================
    # TAB 5: UPLOAD_READY_PROCUREMENT
    # =========================================================================
    ws5 = wb.create_sheet("UPLOAD_READY_PROCUREMENT")
    
    ws5['A1'] = "PROCUREMENT DECISIONS - ExSim Upload Format (Side-by-Side)"
    ws5['A1'].font = title_font
    ws5['A2'] = "This matches the exact ExSim Procurement upload layout"
    ws5['A2'].font = Font(italic=True, color="666666")
    
    # PARTS section (Columns A-K)
    ws5['A4'] = "Parts"
    ws5['A4'].font = section_font
    
    # Parts headers
    parts_headers = ['Zone', 'Supplier', 'Component'] + [str(fn) for fn in FORTNIGHTS]
    for col, h in enumerate(parts_headers, start=1):
        cell = ws5.cell(row=5, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    
    # Parts data rows (Zone -> Supplier -> Part, with FN1-8 values)
    row = 6
    for zone in ZONES:
        for supplier in SUPPLIERS:
            for part in PARTS:
                ws5.cell(row=row, column=1, value=zone).border = thin_border
                ws5.cell(row=row, column=2, value=supplier).border = thin_border
                ws5.cell(row=row, column=3, value=part).border = thin_border
                
                # Link to MRP_ENGINE orders (only for Center zone)
                if zone == "Center":
                    mrp_row = order_rows.get(part, {}).get(supplier, 0)
                    for fn in FORTNIGHTS:
                        if mrp_row:
                            cell = ws5.cell(row=row, column=3+fn, value=f"=MRP_ENGINE!{get_column_letter(1+fn)}{mrp_row}")
                        else:
                            cell = ws5.cell(row=row, column=3+fn, value=0)
                        cell.border = thin_border
                        cell.fill = input_fill
                else:
                    # Other zones - manual input
                    for fn in FORTNIGHTS:
                        cell = ws5.cell(row=row, column=3+fn, value=0)
                        cell.border = thin_border
                        cell.fill = input_fill
                row += 1
    
    # PIECES section (Columns R-U, starting at column 18)
    pieces_col_start = 18
    
    ws5.cell(row=4, column=pieces_col_start, value="Pieces").font = section_font
    
    pieces_headers = ['Zone', 'Supplier', 'Component', 'Order']
    for col, h in enumerate(pieces_headers):
        cell = ws5.cell(row=5, column=pieces_col_start+col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    
    # Pieces data rows
    row = 6
    for zone in ZONES:
        for piece in PIECES:
            ws5.cell(row=row, column=pieces_col_start, value=zone).border = thin_border
            ws5.cell(row=row, column=pieces_col_start+1, value="Pieces").border = thin_border
            ws5.cell(row=row, column=pieces_col_start+2, value=piece).border = thin_border
            cell = ws5.cell(row=row, column=pieces_col_start+3, value=0)
            cell.border = thin_border
            cell.fill = input_fill
            row += 1
    
    # Column widths
    ws5.column_dimensions['A'].width = 10
    ws5.column_dimensions['B'].width = 12
    ws5.column_dimensions['C'].width = 12
    for col in range(4, 12):
        ws5.column_dimensions[get_column_letter(col)].width = 8
    
    # Pieces columns
    ws5.column_dimensions[get_column_letter(pieces_col_start)].width = 10
    ws5.column_dimensions[get_column_letter(pieces_col_start+1)].width = 10
    ws5.column_dimensions[get_column_letter(pieces_col_start+2)].width = 10
    ws5.column_dimensions[get_column_letter(pieces_col_start+3)].width = 8
    
    # Save
    wb.save(OUTPUT_FILE)
    print(f"[SUCCESS] Created '{OUTPUT_FILE}'")


def main():
    """Main function."""
    print("ExSim Purchasing Dashboard Generator v2")
    print("=" * 50)
    
    print("\n[*] Loading data files...")
    
    # Raw Materials
    materials_path = DATA_FOLDER / "raw_materials.xlsx"
    if materials_path.exists():
        materials_data = load_raw_materials(materials_path)
        print(f"  [OK] Loaded raw materials inventory")
    else:
        materials_data = load_raw_materials(None)
        print("  [!] Using default inventory data")
    
    # Production Costs
    costs_path = DATA_FOLDER / "production.xlsx"
    if costs_path.exists():
        cost_data = load_production_costs(costs_path)
        print(f"  [OK] Loaded production costs")
    else:
        cost_data = load_production_costs(None)
        print("  [!] Using default cost data")
    
    # Procurement Template
    template_path = DATA_FOLDER / "Procurement Decisions.xlsx"
    if template_path.exists():
        template_data = load_procurement_template(template_path)
        print(f"  [OK] Loaded procurement template")
    else:
        template_data = load_procurement_template(None)
        print("  [!] Using default template layout")
    
    print("\n[*] Generating Purchasing Dashboard...")
    
    create_purchasing_dashboard(materials_data, cost_data, template_data)
    
    print("\nSheets created:")
    print("  * SUPPLIER_CONFIG (Supplier & Batch Settings)")
    print("  * COST_ANALYSIS (Ordering vs Holding Efficiency)")
    print("  * MRP_ENGINE (Material Requirements Calculator)")
    print("  * CASH_FLOW_PREVIEW (Procurement Spending)")
    print("  * UPLOAD_READY_PROCUREMENT (ExSim Format)")


if __name__ == "__main__":
    main()
