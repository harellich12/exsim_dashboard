"""
ExSim Production Dashboard Zones - Zone-Specific Production Planning

Handles physical separation of resources (Machines/Inventory) across
different geographic zones (Center, West, North, East, South).

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
OUTPUT_FILE = "Production_Dashboard_Zones.xlsx"

FORTNIGHTS = list(range(1, 9))  # 1-8
ZONES = ["Center", "West", "North", "East", "South"]
SECTIONS = ["Section 1", "Section 2", "Section 3"]
MACHINE_TYPES = ["M1", "M2", "M3-alpha", "M3-beta", "M4"]

# Default production parameters
DEFAULT_NOMINAL_RATE = 200  # Units per machine per FN
DEFAULT_VARIABLE_COST = 40  # $ per unit


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
# DATA LOADING (Zone-Specific)
# =============================================================================

def load_raw_materials_by_zone(filepath):
    """Load raw materials inventory grouped by zone."""
    df = load_excel_file(filepath)
    
    # Initialize all zones with defaults
    data = {zone: {'part_a': 0, 'part_b': 0} for zone in ZONES}
    
    if df is None:
        # Default: Only Center has inventory
        data['Center'] = {'part_a': 4000, 'part_b': 1000}
        return data
    
    current_zone = None
    current_part = None
    
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        
        # Detect zone from section headers (e.g., "Center - Section 1")
        for zone in ZONES:
            if zone.lower() in first_val.lower() and 'section' in first_val.lower():
                current_zone = zone
                break
        
        # Detect part type
        if 'part a' in first_val.lower():
            current_part = 'part_a'
        elif 'part b' in first_val.lower():
            current_part = 'part_b'
        
        # Get final inventory
        if 'final' in first_val.lower() and 'inventory' in first_val.lower():
            val = parse_numeric(row.iloc[8]) if len(row) > 8 else 0
            if current_zone and current_part:
                data[current_zone][current_part] = val
    
    return data


def load_finished_goods_by_zone(filepath):
    """Load finished goods inventory grouped by zone."""
    df = load_excel_file(filepath)
    
    data = {zone: {'capacity': 0, 'inventory': 0} for zone in ZONES}
    
    if df is None:
        data['Center'] = {'capacity': 4800, 'inventory': 500}
        return data
    
    current_zone = None
    
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        
        # Detect zone
        for zone in ZONES:
            if zone.lower() in first_val.lower() and ('section' in first_val.lower() or 'warehouse' in first_val.lower()):
                current_zone = zone
                break
        
        if 'capacity' in first_val.lower():
            import re
            match = re.search(r'(\d+)', first_val.replace(',', ''))
            if match and current_zone:
                data[current_zone]['capacity'] = int(match.group(1))
        
        if 'final' in first_val.lower() and 'inventory' in first_val.lower():
            val = parse_numeric(row.iloc[8]) if len(row) > 8 else 0
            if current_zone:
                data[current_zone]['inventory'] = val
    
    return data


def load_workers_by_zone(filepath):
    """Load workers balance grouped by zone."""
    df = load_excel_file(filepath)
    
    # Default: Workers only in Center and West
    data = {zone: {'workers': 0, 'absenteeism': 0} for zone in ZONES}
    
    if df is None:
        data['Center'] = {'workers': 219, 'absenteeism': 0}
        data['West'] = {'workers': 71, 'absenteeism': 0}
        return data
    
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        
        if 'workers assigned' in first_val.lower():
            # Columns: Center, West, North, East, South, Total
            for col_idx, zone in enumerate(['Center', 'West', 'North', 'East', 'South'], start=1):
                if col_idx < len(row):
                    data[zone]['workers'] = parse_numeric(row.iloc[col_idx])
        
        if 'absenteeism' in first_val.lower():
            for col_idx, zone in enumerate(['Center', 'West', 'North', 'East', 'South'], start=1):
                if col_idx < len(row):
                    data[zone]['absenteeism'] = parse_numeric(row.iloc[col_idx])
    
    return data


def load_machines_by_zone(filepath):
    """Load machine counts and modules grouped by zone."""
    df = load_excel_file(filepath)
    
    # Default: All machines in Center
    data = {zone: {'machines': 0, 'modules': 0, 'modules_used': 0} for zone in ZONES}
    
    if df is None:
        data['Center'] = {'machines': 57, 'modules': 72, 'modules_used': 69}
        return data
    
    # For now, assume all machines are in Center (most common scenario)
    total_machines = 0
    modules_available = 72
    modules_occupied = 0
    
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        
        # Get machine counts
        for mt in MACHINE_TYPES:
            if first_val == mt:
                for col_idx in reversed(range(1, min(11, len(row)))):
                    val = parse_numeric(row.iloc[col_idx])
                    if val > 0:
                        total_machines += int(val)
                        break
                break
        
        if 'available' in first_val.lower():
            modules_available = parse_numeric(row.iloc[1]) if len(row) > 1 else 72
        
        if 'occupied' in first_val.lower():
            modules_occupied = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
    
    # Assign all to Center (default scenario)
    data['Center'] = {
        'machines': total_machines,
        'modules': int(modules_available),
        'modules_used': int(modules_occupied)
    }
    
    return data


def load_production_template(filepath):
    """Load production decisions template."""
    try:
        df = pd.read_excel(filepath, sheet_name='Production', header=None)
        return {'df': df, 'exists': True}
    except:
        return {'df': None, 'exists': False}


# =============================================================================
# EXCEL GENERATION
# =============================================================================

def create_zones_dashboard(materials_data, fg_data, workers_data, 
                           machines_data, template_data):
    """Create the Zone-Specific Production Dashboard."""
    
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
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # Track zone output rows for linking
    zone_output_rows = {}
    
    # =========================================================================
    # TAB 1: ZONE_CALCULATORS
    # =========================================================================
    ws1 = wb.active
    ws1.title = "ZONE_CALCULATORS"
    
    ws1['A1'] = "ZONE-SPECIFIC PRODUCTION CALCULATORS"
    ws1['A1'].font = title_font
    ws1['A2'] = "Each zone has its own resources. Resources in Center do NOT count towards West capacity."
    ws1['A2'].font = Font(italic=True, color="666666")
    
    row = 4
    
    for zone in ZONES:
        zone_data = {
            'machines': machines_data.get(zone, {}).get('machines', 0),
            'materials': materials_data.get(zone, {}).get('part_a', 0),
            'workers': workers_data.get(zone, {}).get('workers', 0),
        }
        
        # Zone Header
        ws1.merge_cells(f'A{row}:H{row}')
        cell = ws1.cell(row=row, column=1, value=f"═══ {zone.upper()} ZONE ═══")
        cell.font = zone_font
        cell.fill = zone_fills[zone]
        cell.alignment = Alignment(horizontal='center')
        row += 1
        
        # Zone Parameters
        params = [
            ("Machines in Zone", zone_data['machines']),
            ("Material Stock (Part A)", zone_data['materials']),
            ("Workers in Zone", zone_data['workers']),
            ("Nominal Rate/Machine", DEFAULT_NOMINAL_RATE),
        ]
        
        for i, (label, value) in enumerate(params):
            ws1.cell(row=row, column=1, value=label).border = thin_border
            cell = ws1.cell(row=row, column=2, value=value)
            cell.border = thin_border
            cell.fill = ref_fill
            row += 1
        
        params_end = row - 1
        row += 1
        
        # Production Schedule Headers
        calc_headers = ['Fortnight', 'Target', 'Overtime', 
                        'Local Capacity', 'Material Cap', 'REAL OUTPUT', 'Shipment?']
        for col, h in enumerate(calc_headers, start=1):
            cell = ws1.cell(row=row, column=col, value=h)
            cell.font = header_font
            cell.fill = zone_fills[zone]
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center')
        row += 1
        
        data_start = row
        for fn in FORTNIGHTS:
            ws1.cell(row=row, column=1, value=f"FN{fn}").border = thin_border
            
            # Target (input)
            cell = ws1.cell(row=row, column=2, value=0 if zone_data['machines'] == 0 else 500)
            cell.border = thin_border
            cell.fill = input_fill
            
            # Overtime Y/N (input)
            cell = ws1.cell(row=row, column=3, value="N")
            cell.border = thin_border
            cell.fill = input_fill
            cell.alignment = Alignment(horizontal='center')
            
            # Local Capacity = Machines × Rate
            cell = ws1.cell(row=row, column=4, value=f"=$B${params_end-3}*$B${params_end}")
            cell.border = thin_border
            cell.fill = calc_fill
            
            # Material Cap (simplified)
            cell = ws1.cell(row=row, column=5, value=f"=$B${params_end-2}")
            cell.border = thin_border
            cell.fill = calc_fill
            
            # REAL OUTPUT = MIN(Target, Capacity, Material)
            cell = ws1.cell(row=row, column=6, value=f"=MIN(B{row},D{row},E{row})")
            cell.border = thin_border
            cell.fill = output_fill
            cell.font = Font(bold=True)
            
            # Shipment Alert
            cell = ws1.cell(row=row, column=7, 
                value=f'=IF(B{row}>E{row}, "SHIPMENT NEEDED!", "OK")')
            cell.border = thin_border
            
            row += 1
        
        # Total row
        ws1.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
        cell = ws1.cell(row=row, column=2, value=f"=SUM(B{data_start}:B{row-1})")
        cell.fill = input_fill
        cell = ws1.cell(row=row, column=6, value=f"=SUM(F{data_start}:F{row-1})")
        cell.fill = output_fill
        cell.font = Font(bold=True)
        
        zone_output_rows[zone] = {'total': row, 'data_start': data_start, 'data_end': row-1}
        
        row += 3  # Space between zones
    
    # Column widths
    ws1.column_dimensions['A'].width = 22
    for col in range(2, 8):
        ws1.column_dimensions[get_column_letter(col)].width = 14
    
    # =========================================================================
    # TAB 2: RESOURCE_MGR
    # =========================================================================
    ws2 = wb.create_sheet("RESOURCE_MGR")
    
    ws2['A1'] = "RESOURCE MANAGER - Zone-by-Zone Asset Allocation"
    ws2['A1'].font = title_font
    
    # Section A: Assignments by Zone
    ws2['A3'] = "SECTION A: MACHINE ASSIGNMENTS BY ZONE"
    ws2['A3'].font = section_font
    
    assign_headers = ['Zone', 'Section', 'Machines Assigned', 'Workers Needed', 'Status']
    for col, h in enumerate(assign_headers, start=1):
        cell = ws2.cell(row=5, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    
    row = 6
    for zone in ZONES:
        for section in SECTIONS:
            cell = ws2.cell(row=row, column=1, value=zone)
            cell.border = thin_border
            cell.fill = zone_fills[zone]
            cell.font = Font(color="FFFFFF")
            
            ws2.cell(row=row, column=2, value=section).border = thin_border
            
            # Default machines (only Center has machines initially)
            default_machines = 0
            if zone == "Center":
                default_machines = {'Section 1': 20, 'Section 2': 20, 'Section 3': 17}.get(section, 10)
            
            cell = ws2.cell(row=row, column=3, value=default_machines)
            cell.border = thin_border
            cell.fill = input_fill
            
            cell = ws2.cell(row=row, column=4, value=f"=C{row}*5")
            cell.border = thin_border
            cell.fill = calc_fill
            
            cell = ws2.cell(row=row, column=5, value="OK")
            cell.border = thin_border
            cell.fill = output_fill
            
            row += 1
    
    assign_end = row - 1
    row += 3
    
    # Section B: Expansion by Zone
    ws2.cell(row=row, column=1, value="SECTION B: EXPANSION RECOMMENDATIONS BY ZONE").font = section_font
    row += 2
    
    exp_headers = ['Zone', 'Target Capacity', 'Current Machines', 'Capacity Gap', 'Recommendation']
    for col, h in enumerate(exp_headers, start=1):
        cell = ws2.cell(row=row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    row += 1
    
    expansion_start = row
    for zone in ZONES:
        cell = ws2.cell(row=row, column=1, value=zone)
        cell.border = thin_border
        cell.fill = zone_fills[zone]
        cell.font = Font(color="FFFFFF")
        
        # Link to Zone Calculator total
        zone_total_row = zone_output_rows.get(zone, {}).get('total', 10)
        cell = ws2.cell(row=row, column=2, value=f"=ZONE_CALCULATORS!B{zone_total_row}")
        cell.border = thin_border
        cell.fill = ref_fill
        
        # Current machines
        current_machines = machines_data.get(zone, {}).get('machines', 0)
        ws2.cell(row=row, column=3, value=current_machines).border = thin_border
        
        # Capacity Gap
        cell = ws2.cell(row=row, column=4, value=f"=B{row}-(C{row}*{DEFAULT_NOMINAL_RATE})")
        cell.border = thin_border
        cell.fill = calc_fill
        
        # Recommendation
        cell = ws2.cell(row=row, column=5, 
            value=f'=IF(D{row}>0, "Buy "&ROUNDUP(D{row}/{DEFAULT_NOMINAL_RATE},0)&" machines", "OK")')
        cell.border = thin_border
        cell.fill = output_fill
        
        row += 1
    
    row += 3
    
    # Section C: Real Estate by Zone
    ws2.cell(row=row, column=1, value="SECTION C: REAL ESTATE (MODULES) BY ZONE").font = section_font
    row += 2
    
    re_headers = ['Zone', 'Machines', 'Module Slots', 'Free Slots', 'Recommendation']
    for col, h in enumerate(re_headers, start=1):
        cell = ws2.cell(row=row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    row += 1
    
    for zone in ZONES:
        zone_machines = machines_data.get(zone, {})
        
        cell = ws2.cell(row=row, column=1, value=zone)
        cell.border = thin_border
        cell.fill = zone_fills[zone]
        cell.font = Font(color="FFFFFF")
        
        ws2.cell(row=row, column=2, value=zone_machines.get('machines', 0)).border = thin_border
        ws2.cell(row=row, column=3, value=zone_machines.get('modules', 0)).border = thin_border
        
        cell = ws2.cell(row=row, column=4, value=f"=C{row}-B{row}")
        cell.border = thin_border
        cell.fill = calc_fill
        
        cell = ws2.cell(row=row, column=5, 
            value=f'=IF(D{row}<5, "Buy module in {zone}", "OK")')
        cell.border = thin_border
        cell.fill = output_fill
        
        row += 1
    
    # Column widths
    ws2.column_dimensions['A'].width = 12
    ws2.column_dimensions['B'].width = 16
    ws2.column_dimensions['C'].width = 18
    ws2.column_dimensions['D'].width = 14
    ws2.column_dimensions['E'].width = 28
    
    # =========================================================================
    # TAB 3: UPLOAD_READY_PRODUCTION
    # =========================================================================
    ws3 = wb.create_sheet("UPLOAD_READY_PRODUCTION")
    
    ws3['A1'] = "PRODUCTION DECISIONS - ExSim Upload (Zone-Mapped)"
    ws3['A1'].font = title_font
    ws3['A2'] = "Values linked from Zone Calculators"
    ws3['A2'].font = Font(italic=True, color="666666")
    
    # Block 1: Production Targets by Zone
    ws3['A4'] = "Production Targets"
    ws3['A4'].font = section_font
    
    headers = ['Zone', 'Product', 'Target', 'Overtime']
    for col, h in enumerate(headers, start=1):
        cell = ws3.cell(row=5, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    
    row = 6
    for zone in ZONES:
        cell = ws3.cell(row=row, column=1, value=zone)
        cell.border = thin_border
        cell.fill = zone_fills[zone]
        cell.font = Font(color="FFFFFF")
        
        ws3.cell(row=row, column=2, value="A").border = thin_border
        
        # Link to zone total output
        zone_total = zone_output_rows.get(zone, {}).get('total', 10)
        cell = ws3.cell(row=row, column=3, value=f"=ZONE_CALCULATORS!F{zone_total}")
        cell.border = thin_border
        cell.fill = calc_fill
        
        # Link to first FN overtime of zone
        zone_data_start = zone_output_rows.get(zone, {}).get('data_start', 10)
        cell = ws3.cell(row=row, column=4, value=f"=ZONE_CALCULATORS!C{zone_data_start}")
        cell.border = thin_border
        cell.fill = calc_fill
        
        row += 1
    
    row += 2
    
    # Block 2: Machine Purchases by Zone
    ws3.cell(row=row, column=1, value="Machine Purchases").font = section_font
    row += 1
    
    headers = ['Zone', 'Machine Type', 'Quantity']
    for col, h in enumerate(headers, start=1):
        cell = ws3.cell(row=row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    row += 1
    
    for zone in ZONES:
        for mt in ['M1', 'M2']:  # Most common machine types
            cell = ws3.cell(row=row, column=1, value=zone)
            cell.border = thin_border
            cell.fill = zone_fills[zone]
            cell.font = Font(color="FFFFFF")
            
            ws3.cell(row=row, column=2, value=mt).border = thin_border
            
            cell = ws3.cell(row=row, column=3, value=0)
            cell.border = thin_border
            cell.fill = input_fill
            
            row += 1
    
    row += 2
    
    # Block 3: Section Assignments by Zone
    ws3.cell(row=row, column=1, value="Section Assignments").font = section_font
    row += 1
    
    headers = ['Zone', 'Section', 'Machines', 'Workers']
    for col, h in enumerate(headers, start=1):
        cell = ws3.cell(row=row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    row += 1
    
    # Link to RESOURCE_MGR assignments
    assign_row = 6
    for zone in ZONES:
        for section in SECTIONS:
            cell = ws3.cell(row=row, column=1, value=zone)
            cell.border = thin_border
            cell.fill = zone_fills[zone]
            cell.font = Font(color="FFFFFF")
            
            ws3.cell(row=row, column=2, value=section).border = thin_border
            
            cell = ws3.cell(row=row, column=3, value=f"=RESOURCE_MGR!C{assign_row}")
            cell.border = thin_border
            cell.fill = calc_fill
            
            cell = ws3.cell(row=row, column=4, value=f"=RESOURCE_MGR!D{assign_row}")
            cell.border = thin_border
            cell.fill = calc_fill
            
            assign_row += 1
            row += 1
    
    # Column widths
    for col in range(1, 5):
        ws3.column_dimensions[get_column_letter(col)].width = 15
    
    # Save
    wb.save(OUTPUT_FILE)
    print(f"[SUCCESS] Created '{OUTPUT_FILE}'")


def main():
    """Main function."""
    print("ExSim Production Dashboard Zones Generator")
    print("=" * 50)
    
    print("\n[*] Loading zone-specific data files...")
    
    # Raw Materials by Zone
    materials_path = DATA_FOLDER / "raw_materials.xlsx"
    if materials_path.exists():
        materials_data = load_raw_materials_by_zone(materials_path)
        zones_with_materials = [z for z in ZONES if materials_data[z]['part_a'] > 0]
        print(f"  [OK] Loaded materials for zones: {zones_with_materials if zones_with_materials else 'Using defaults'}")
    else:
        materials_data = load_raw_materials_by_zone(None)
        print("  [!] Using default materials data")
    
    # Finished Goods by Zone
    fg_path = DATA_FOLDER / "finished_goods_inventory.xlsx"
    if fg_path.exists():
        fg_data = load_finished_goods_by_zone(fg_path)
        print(f"  [OK] Loaded finished goods by zone")
    else:
        fg_data = load_finished_goods_by_zone(None)
        print("  [!] Using default FG data")
    
    # Workers by Zone
    workers_path = DATA_FOLDER / "workers_balance_overtime.xlsx"
    if workers_path.exists():
        workers_data = load_workers_by_zone(workers_path)
        for zone in ZONES:
            if workers_data[zone]['workers'] > 0:
                print(f"  [OK] {zone}: {workers_data[zone]['workers']:.0f} workers")
    else:
        workers_data = load_workers_by_zone(None)
        print("  [!] Using default workers data")
    
    # Machines by Zone
    machines_path = DATA_FOLDER / "machine_spaces.xlsx"
    if machines_path.exists():
        machines_data = load_machines_by_zone(machines_path)
        for zone in ZONES:
            if machines_data[zone]['machines'] > 0:
                print(f"  [OK] {zone}: {machines_data[zone]['machines']} machines, {machines_data[zone]['modules']} module slots")
    else:
        machines_data = load_machines_by_zone(None)
        print("  [!] Using default machines data")
    
    # Template
    template_path = DATA_FOLDER / "Production Decisions.xlsx"
    template_data = load_production_template(template_path)
    if template_data['exists']:
        print(f"  [OK] Loaded production template")
    else:
        print("  [!] Using default template layout")
    
    print("\n[*] Generating Zone-Specific Dashboard...")
    
    create_zones_dashboard(materials_data, fg_data, workers_data,
                           machines_data, template_data)
    
    print("\nSheets created:")
    print("  * ZONE_CALCULATORS (5 Zone-Specific Production Blocks)")
    print("  * RESOURCE_MGR (Assignments/Expansion/Modules by Zone)")
    print("  * UPLOAD_READY_PRODUCTION (ExSim Format)")


if __name__ == "__main__":
    main()
