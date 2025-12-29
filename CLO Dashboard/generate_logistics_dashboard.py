"""
ExSim Logistics Dashboard - Supply Network Optimization

Balances Inventory levels across zones using Shipments.
Handles warehouse capacity, transport modes, and stockout prevention.

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
OUTPUT_FILE = "Logistics_Dashboard.xlsx"

FORTNIGHTS = list(range(1, 9))  # 1-8
ZONES = ["Center", "West", "North", "East", "South"]
TRANSPORT_MODES = ["Train", "Truck", "Plane"]
DEFAULT_MATERIAL = "Electroclean"

# Default transport configuration
# Default transport configuration
DEFAULT_TRANSPORT = {
    "Train": {"lead_time": 0, "cost": 0},
    "Truck": {"lead_time": 0, "cost": 0},
    "Plane": {"lead_time": 0, "cost": 0},
}

# Default warehouse configuration
# Default warehouse configuration
DEFAULT_WAREHOUSE = {
    "Center": {"capacity": 0, "cost_per_module": 0, "capacity_per_module": 0},
    "West": {"capacity": 0, "cost_per_module": 0, "capacity_per_module": 0},
    "North": {"capacity": 0, "cost_per_module": 0, "capacity_per_module": 0},
    "East": {"capacity": 0, "cost_per_module": 0, "capacity_per_module": 0},
    "South": {"capacity": 0, "cost_per_module": 0, "capacity_per_module": 0},
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

def load_finished_goods_by_zone(filepath):
    """Load finished goods inventory grouped by zone."""
    df = load_excel_file(filepath)
    
    # Default data
    data = {zone: {'capacity': DEFAULT_WAREHOUSE[zone]['capacity'], 
                   'inventory': 0} for zone in ZONES}
    
    if df is None:
        return data
    
    current_zone_idx = 0
    zone_order = ['Center', 'West', 'North', 'East', 'South']
    
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        
        # Detect capacity header (new zone section)
        if 'capacity:' in first_val.lower():
            import re
            match = re.search(r'(\d+)', first_val.replace(',', ''))
            if match and current_zone_idx < len(zone_order):
                zone = zone_order[current_zone_idx]
                data[zone]['capacity'] = int(match.group(1))
        
        # Get final inventory
        if 'final inventory' in first_val.lower():
            if current_zone_idx < len(zone_order):
                zone = zone_order[current_zone_idx]
                # Get last fortnight value (column 8)
                val = parse_numeric(row.iloc[8]) if len(row) > 8 else 0
                data[zone]['inventory'] = val
                current_zone_idx += 1
    
    return data


def load_logistics_template(filepath):
    """Load logistics decisions template."""
    try:
        df = pd.read_excel(filepath, sheet_name='Logistics', header=None)
        return {'df': df, 'exists': True}
    except:
        return {'df': None, 'exists': False}


def load_shipping_costs(filepath):
    """Load logistics shipping costs."""
    df = load_excel_file(filepath)
    
    data = {'total_shipping_cost': 0}
    
    if df is None:
        return data
    
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
        
        if 'shipping' in first_val and 'cost' in first_val:
            for col_idx in range(1, min(10, len(row))):
                val = parse_numeric(row.iloc[col_idx])
                if val > 0:
                    data['total_shipping_cost'] = val
                    break
    
    return data


# =============================================================================
# EXCEL GENERATION
# =============================================================================

def create_logistics_dashboard(inventory_data, template_data, cost_data):
    """Create the Logistics Dashboard."""
    
    wb = Workbook()
    
    # Styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    section_font = Font(bold=True, size=12, color="2F5496")
    title_font = Font(bold=True, size=14, color="2F5496")
    zone_font = Font(bold=True, size=11, color="FFFFFF")
    zone_fills = {
        'Center': PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid"),
        'West': PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid"),
        'North': PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid"),
        'East': PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid"),
        'South': PatternFill(start_color="9E480E", end_color="9E480E", fill_type="solid"),
    }
    input_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    calc_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    output_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    ref_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    purple_fill = PatternFill(start_color="E4DFEC", end_color="E4DFEC", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # Track zone data rows for formulas
    zone_data_rows = {}
    
    # =========================================================================
    # TAB 1: ROUTE_CONFIG
    # =========================================================================
    ws1 = wb.active
    ws1.title = "ROUTE_CONFIG"
    
    ws1['A1'] = "ROUTE CONFIGURATION - Transport Physics"
    ws1['A1'].font = title_font
    ws1['A2'] = "Define transport modes and warehouse costs. Yellow cells are editable."
    ws1['A2'].font = Font(italic=True, color="666666")
    
    # MODES CONFIG
    ws1['A4'] = "TABLE 1: TRANSPORT MODES"
    ws1['A4'].font = section_font
    
    mode_headers = ['Mode', 'Lead Time (Fortnights)', 'Cost Per Unit ($)']
    for col, h in enumerate(mode_headers, start=1):
        cell = ws1.cell(row=5, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    
    row = 6
    for mode, config in DEFAULT_TRANSPORT.items():
        ws1.cell(row=row, column=1, value=mode).border = thin_border
        
        cell = ws1.cell(row=row, column=2, value=config['lead_time'])
        cell.border = thin_border
        cell.fill = input_fill
        
        cell = ws1.cell(row=row, column=3, value=config['cost'])
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '$#,##0'
        
        row += 1
    
    row += 2
    
    # WAREHOUSE CONFIG
    ws1.cell(row=row, column=1, value="TABLE 2: WAREHOUSE CONFIGURATION").font = section_font
    row += 1
    
    wh_headers = ['Zone', 'Current Capacity', 'Cost Per Module', 'Capacity Per Module']
    for col, h in enumerate(wh_headers, start=1):
        cell = ws1.cell(row=row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    row += 1
    
    wh_config_start = row
    for zone in ZONES:
        zone_inv = inventory_data.get(zone, {})
        capacity = zone_inv.get('capacity', DEFAULT_WAREHOUSE[zone]['capacity'])
        
        cell = ws1.cell(row=row, column=1, value=zone)
        cell.border = thin_border
        cell.fill = zone_fills[zone]
        cell.font = Font(color="FFFFFF")
        
        cell = ws1.cell(row=row, column=2, value=capacity)
        cell.border = thin_border
        cell.fill = ref_fill
        
        cell = ws1.cell(row=row, column=3, value=DEFAULT_WAREHOUSE[zone]['cost_per_module'])
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '$#,##0'
        
        cell = ws1.cell(row=row, column=4, value=DEFAULT_WAREHOUSE[zone]['capacity_per_module'])
        cell.border = thin_border
        cell.fill = input_fill
        
        row += 1
    
    # Column widths
    ws1.column_dimensions['A'].width = 12
    ws1.column_dimensions['B'].width = 22
    ws1.column_dimensions['C'].width = 18
    ws1.column_dimensions['D'].width = 22
    
    # =========================================================================
    # TAB 2: INVENTORY_TETRIS
    # =========================================================================
    ws2 = wb.create_sheet("INVENTORY_TETRIS")
    
    ws2['A1'] = "INVENTORY TETRIS - Zone-by-Zone Balance"
    ws2['A1'].font = title_font
    ws2['A2'] = "Balance inventory using shipments. Watch for STOCKOUT (red) and OVERFLOW (purple) flags."
    ws2['A2'].font = Font(italic=True, color="666666")
    
    row = 4
    
    for zone in ZONES:
        zone_inv = inventory_data.get(zone, {})
        opening_inv = zone_inv.get('inventory', 0)
        capacity = zone_inv.get('capacity', DEFAULT_WAREHOUSE[zone]['capacity'])
        
        # Zone Header
        ws2.merge_cells(f'A{row}:H{row}')
        cell = ws2.cell(row=row, column=1, value=f"═══ {zone.upper()} ZONE (Capacity: {capacity:,}) ═══")
        cell.font = zone_font
        cell.fill = zone_fills[zone]
        cell.alignment = Alignment(horizontal='center')
        row += 1
        
        # Parameters
        ws2.cell(row=row, column=1, value="Opening Inventory").border = thin_border
        cell = ws2.cell(row=row, column=2, value=opening_inv)
        cell.border = thin_border
        cell.fill = ref_fill
        
        ws2.cell(row=row, column=4, value="Capacity").border = thin_border
        cell = ws2.cell(row=row, column=5, value=capacity)
        cell.border = thin_border
        cell.fill = ref_fill
        
        ws2.cell(row=row, column=7, value="Rent Modules?").border = thin_border
        cell = ws2.cell(row=row, column=8, value=0)
        cell.border = thin_border
        cell.fill = input_fill
        
        params_row = row
        row += 2
        
        # Headers
        inv_headers = ['Fortnight', 'Production', 'Sales', 'Outgoing', 'Incoming', 
                       'Projected Inv', 'Flag']
        for col, h in enumerate(inv_headers, start=1):
            cell = ws2.cell(row=row, column=col, value=h)
            cell.font = header_font
            cell.fill = zone_fills[zone]
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center')
        row += 1
        
        data_start = row
        for fn in FORTNIGHTS:
            ws2.cell(row=row, column=1, value=f"FN{fn}").border = thin_border
            
            # Production (input from Production Dashboard)
            cell = ws2.cell(row=row, column=2, value=0)
            cell.border = thin_border
            cell.fill = input_fill
            
            # Sales Forecast (input from Marketing Dashboard)
            cell = ws2.cell(row=row, column=3, value=0)
            cell.border = thin_border
            cell.fill = input_fill
            
            # Outgoing Shipments (negative, manual input)
            cell = ws2.cell(row=row, column=4, value=0)
            cell.border = thin_border
            cell.fill = input_fill
            
            # Incoming Shipments (positive, manual input)
            cell = ws2.cell(row=row, column=5, value=0)
            cell.border = thin_border
            cell.fill = input_fill
            
            # Projected Inventory = Prev + Production + Incoming - Outgoing - Sales
            if fn == 1:
                formula = f"=$B${params_row}+B{row}+E{row}-D{row}-C{row}"
            else:
                formula = f"=F{row-1}+B{row}+E{row}-D{row}-C{row}"
            cell = ws2.cell(row=row, column=6, value=formula)
            cell.border = thin_border
            cell.fill = calc_fill
            cell.font = Font(bold=True)
            
            # Flag - reference ROUTE_CONFIG for capacity per module
            # Zone rows in ROUTE_CONFIG: Center=12, West=13, North=14, East=15, South=16
            zone_config_row = 12 + ZONES.index(zone)
            cell = ws2.cell(row=row, column=7, 
                value=f'=IF(F{row}<0,"STOCKOUT: SHIP HERE!",IF(F{row}>$E${params_row}+(H${params_row}*ROUTE_CONFIG!$D${zone_config_row}),"OVERFLOW: RENT!","OK"))')
            cell.border = thin_border
            
            row += 1
        
        data_end = row - 1
        zone_data_rows[zone] = {
            'start': data_start, 
            'end': data_end, 
            'params': params_row,
            'rent_cell': f'H{params_row}'
        }
        
        # Add conditional formatting for flags
        ws2.conditional_formatting.add(
            f'G{data_start}:G{data_end}',
            FormulaRule(formula=[f'LEFT(G{data_start},8)="STOCKOUT"'], fill=red_fill)
        )
        ws2.conditional_formatting.add(
            f'G{data_start}:G{data_end}',
            FormulaRule(formula=[f'LEFT(G{data_start},8)="OVERFLOW"'], fill=purple_fill)
        )
        
        row += 2
    
    # Column widths
    ws2.column_dimensions['A'].width = 12
    for col in range(2, 8):
        ws2.column_dimensions[get_column_letter(col)].width = 14
    
    # =========================================================================
    # TAB 3: SHIPMENT_BUILDER
    # =========================================================================
    ws3 = wb.create_sheet("SHIPMENT_BUILDER")
    
    ws3['A1'] = "SHIPMENT BUILDER - Plan Your Transfers"
    ws3['A1'].font = title_font
    ws3['A2'] = "Add shipments here. MANUALLY update Outgoing/Incoming in INVENTORY_TETRIS (shifted by Lead Time)."
    ws3['A2'].font = Font(italic=True, color="666666")
    
    ws3['A4'] = "IMPORTANT: After entering shipments here, update INVENTORY_TETRIS Tab 2:"
    ws3['A4'].font = Font(bold=True, color="C00000")
    ws3['A5'] = "• Add NEGATIVE quantity to Origin zone's 'Outgoing' column in the ORDER fortnight"
    ws3['A5'].font = Font(italic=True, color="666666")
    ws3['A6'] = "• Add POSITIVE quantity to Destination zone's 'Incoming' column in ARRIVAL fortnight (Order + Lead Time)"
    ws3['A6'].font = Font(italic=True, color="666666")
    
    # Shipment table
    ws3['A8'] = "SHIPMENT SCHEDULE"
    ws3['A8'].font = section_font
    
    ship_headers = ['#', 'Fortnight', 'Origin', 'Destination', 'Material', 'Mode', 'Quantity', 
                    'Lead Time', 'Arrival FN']
    for col, h in enumerate(ship_headers, start=1):
        cell = ws3.cell(row=9, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    
    # Pre-fill 10 empty shipment rows
    for i in range(10):
        row = 10 + i
        
        ws3.cell(row=row, column=1, value=i+1).border = thin_border
        
        cell = ws3.cell(row=row, column=2, value="")
        cell.border = thin_border
        cell.fill = input_fill
        
        cell = ws3.cell(row=row, column=3, value="")
        cell.border = thin_border
        cell.fill = input_fill
        
        cell = ws3.cell(row=row, column=4, value="")
        cell.border = thin_border
        cell.fill = input_fill
        
        cell = ws3.cell(row=row, column=5, value=DEFAULT_MATERIAL)
        cell.border = thin_border
        cell.fill = input_fill
        
        cell = ws3.cell(row=row, column=6, value="Truck")
        cell.border = thin_border
        cell.fill = input_fill
        
        cell = ws3.cell(row=row, column=7, value="")
        cell.border = thin_border
        cell.fill = input_fill
        
        # Lead Time lookup - reference ROUTE_CONFIG Tab 1 (Train=row 6, Truck=row 7, Plane=row 8)
        cell = ws3.cell(row=row, column=8, value='=IF(F' + str(row) + '="Train",ROUTE_CONFIG!$B$6,IF(F' + str(row) + '="Truck",ROUTE_CONFIG!$B$7,ROUTE_CONFIG!$B$8))')
        cell.border = thin_border
        cell.fill = calc_fill
        
        # Arrival FN
        cell = ws3.cell(row=row, column=9, value=f'=IF(B{row}<>"",B{row}+H{row},"")')
        cell.border = thin_border
        cell.fill = calc_fill
    
    # Column widths
    ws3.column_dimensions['A'].width = 5
    ws3.column_dimensions['B'].width = 12
    ws3.column_dimensions['C'].width = 12
    ws3.column_dimensions['D'].width = 14
    ws3.column_dimensions['E'].width = 14
    ws3.column_dimensions['F'].width = 10
    ws3.column_dimensions['G'].width = 12
    ws3.column_dimensions['H'].width = 12
    ws3.column_dimensions['I'].width = 12
    
    # =========================================================================
    # TAB 4: UPLOAD_READY_LOGISTICS
    # =========================================================================
    ws4 = wb.create_sheet("UPLOAD_READY_LOGISTICS")
    
    ws4['A1'] = "LOGISTICS DECISIONS - ExSim Upload Format (Side-by-Side)"
    ws4['A1'].font = title_font
    ws4['A2'] = "Copy these values to ExSim Logistics upload"
    ws4['A2'].font = Font(italic=True, color="666666")
    
    # Section 1: Warehouses (Left)
    ws4['A4'] = "Warehouses"
    ws4['A4'].font = section_font
    
    wh_headers = ['Zone', 'Buy Modules', 'Rent Modules']
    for col, h in enumerate(wh_headers, start=1):
        cell = ws4.cell(row=5, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    
    row = 6
    for zone in ZONES:
        cell = ws4.cell(row=row, column=1, value=zone)
        cell.border = thin_border
        cell.fill = zone_fills[zone]
        cell.font = Font(color="FFFFFF")
        
        # Buy modules (input)
        cell = ws4.cell(row=row, column=2, value=0)
        cell.border = thin_border
        cell.fill = input_fill
        
        # Rent modules (link to INVENTORY_TETRIS)
        zone_info = zone_data_rows.get(zone, {})
        rent_cell = zone_info.get('rent_cell', 'H5')
        cell = ws4.cell(row=row, column=3, value=f"=INVENTORY_TETRIS!{rent_cell}")
        cell.border = thin_border
        cell.fill = calc_fill
        
        row += 1
    
    # Section 2: Shipments (Right, starting at column 6)
    ship_col = 6
    ws4.cell(row=4, column=ship_col, value="Shipments").font = section_font
    
    ship_headers = ['Fortnight', 'Origin', 'Destination', 'Material', 'Transport', 'Quantity']
    for col, h in enumerate(ship_headers):
        cell = ws4.cell(row=5, column=ship_col+col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    
    # Link to SHIPMENT_BUILDER
    for i in range(10):
        row = 6 + i
        builder_row = 10 + i
        
        for col in range(6):
            # Map columns: B->Fortnight, C->Origin, D->Dest, E->Material, F->Mode, G->Qty
            builder_col = ['B', 'C', 'D', 'E', 'F', 'G'][col]
            cell = ws4.cell(row=row, column=ship_col+col, 
                value=f"=SHIPMENT_BUILDER!{builder_col}{builder_row}")
            cell.border = thin_border
            cell.fill = calc_fill
    
    # Column widths
    ws4.column_dimensions['A'].width = 10
    ws4.column_dimensions['B'].width = 14
    ws4.column_dimensions['C'].width = 14
    for col in range(ship_col, ship_col+6):
        ws4.column_dimensions[get_column_letter(col)].width = 12
    
    # Save
    wb.save(OUTPUT_FILE)
    print(f"[SUCCESS] Created '{OUTPUT_FILE}'")


def main():
    """Main function."""
    print("ExSim Logistics Dashboard Generator")
    print("=" * 50)
    
    print("\n[*] Loading data files...")
    
    # Finished Goods by Zone
    fg_path = DATA_FOLDER / "finished_goods_inventory.xlsx"
    if fg_path.exists():
        inventory_data = load_finished_goods_by_zone(fg_path)
        for zone in ZONES:
            inv = inventory_data[zone]
            if inv['capacity'] > 0 or inv['inventory'] > 0:
                print(f"  [OK] {zone}: Inv={inv['inventory']:.0f}, Cap={inv['capacity']}")
    else:
        inventory_data = load_finished_goods_by_zone(None)
        print("  [!] File not found. Using 0 values.")
    
    # Template
    template_path = DATA_FOLDER / "Logistics Decisions.xlsx"
    template_data = load_logistics_template(template_path)
    if template_data['exists']:
        print(f"  [OK] Loaded logistics template")
    else:
        print("  [!] Using default template layout")
    
    # Shipping Costs
    costs_path = DATA_FOLDER / "logistics.xlsx"
    if costs_path.exists():
        cost_data = load_shipping_costs(costs_path)
        print(f"  [OK] Loaded shipping costs")
    else:
        cost_data = load_shipping_costs(None)
        print("  [!] File not found. Using 0 values.")
    
    print("\n[*] Generating Logistics Dashboard...")
    
    create_logistics_dashboard(inventory_data, template_data, cost_data)
    
    print("\nSheets created:")
    print("  * ROUTE_CONFIG (Transport Modes & Warehouse Costs)")
    print("  * INVENTORY_TETRIS (Zone-by-Zone Balance)")
    print("  * SHIPMENT_BUILDER (Plan Transfers)")
    print("  * UPLOAD_READY_LOGISTICS (ExSim Format)")


if __name__ == "__main__":
    main()
