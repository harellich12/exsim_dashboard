"""
ExSim CPO Dashboard - Workforce Planning & Compensation Dashboard

Calculates Total Cost of Labor (Cash Flow) and estimates Strike Risk based on inflation.
Helps the Chief People Officer manage headcount, salaries, and benefits.

Required libraries: pandas, openpyxl
"""

import pandas as pd
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule, IconSetRule
from openpyxl.chart import PieChart, LineChart, Reference, Series
from openpyxl.chart.label import DataLabelList
import warnings

warnings.filterwarnings('ignore')

# =============================================================================
# CONFIGURATION
# =============================================================================
DATA_FOLDER = Path("data")
OUTPUT_FILE = "CPO_Dashboard.xlsx"

ZONES = ["Center", "West", "North", "East", "South"]

# Default parameters
DEFAULT_HIRING_FEE = 0  # Was 3000
DEFAULT_SEVERANCE = 0  # Was 5000
DEFAULT_BASE_SALARY = 0  # Was 750
DEFAULT_INFLATION_RATE = 0  # Was 0.03

# Default benefits structure
DEFAULT_BENEFITS = [
    ("Training Budget (% of Payroll)", 0, "percent", "Low = More defects"),
    ("Health Insurance (% of Payroll)", 0, "percent", "Reduces absenteeism"),
    ("Profit Sharing (% of Net Profit)", 0, "percent", "Paid on net profit"),
    ("Personal Days (per Worker)", 0, "number", "Labor requirement"),
    ("Union Representatives", 0, "number", "Labor requirement"),
    ("Reduction in Working Hours (%)", 0, "percent", "Reduces capacity"),
    ("Off-Days for Workers", 0, "number", "Extra days off"),
]


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

def load_workers_balance(filepath):
    """Load workers balance data per zone."""
    df = load_excel_file(filepath)
    
    data = {zone: {'workers': 0, 'absenteeism': 0} for zone in ZONES}
    
    if df is None:
        return data
    
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
        
        if 'workers assigned' in first_val:
            for col_idx, zone in enumerate(['Center', 'West', 'North', 'East', 'South'], start=1):
                if col_idx < len(row):
                    data[zone]['workers'] = int(parse_numeric(row.iloc[col_idx]))
        
        if 'absenteeism' in first_val:
            for col_idx, zone in enumerate(['Center', 'West', 'North', 'East', 'South'], start=1):
                if col_idx < len(row):
                    data[zone]['absenteeism'] = parse_numeric(row.iloc[col_idx])
    
    return data


def load_sales_admin(filepath):
    """Load sales & admin data for salespeople info."""
    df = load_excel_file(filepath)
    
    data = {'salespeople_count': 0, 'salespeople_salaries': 0}
    
    if df is None:
        return data
    
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
        
        if 'salespeople salaries' in first_val:
            import re
            amount_str = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ''
            match = re.search(r'(\d+)', amount_str)
            if match:
                data['salespeople_count'] = int(match.group(1))
            
            expense = parse_numeric(row.iloc[2]) if len(row) > 2 else 0
            if expense > 0:
                data['salespeople_salaries'] = expense
    
    return data


def load_labor_costs(filepath):
    """Load labor costs from production data."""
    df = load_excel_file(filepath)
    
    data = {'total_labor': 0}
    
    if df is None:
        return data
    
    total_labor = 0
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
        
        if 'direct and indirect' in first_val:
            cost = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            total_labor += cost
    
    data['total_labor'] = total_labor
    return data


# =============================================================================
# EXCEL GENERATION
# =============================================================================

def create_cpo_dashboard(workers_data, sales_data, labor_data):
    """Create the CPO Workforce Dashboard using openpyxl."""
    
    wb = Workbook()
    
    # Styles
    title_font = Font(bold=True, size=14, color="2F5496")
    section_font = Font(bold=True, size=12, color="2F5496")
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    input_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    calc_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    output_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    ref_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    warning_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    zone_fills = {
        'Center': PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid"),
        'West': PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid"),
        'North': PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid"),
        'East': PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid"),
        'South': PatternFill(start_color="9E480E", end_color="9E480E", fill_type="solid"),
    }
    
    # =========================================================================
    # TAB 1: WORKFORCE_PLANNING
    # =========================================================================
    ws1 = wb.active
    ws1.title = "WORKFORCE_PLANNING"
    
    ws1['A1'] = "WORKFORCE PLANNING - Headcount Management"
    ws1['A1'].font = title_font
    ws1['A2'] = "Yellow cells = User inputs. Green = Outputs for Finance."
    ws1['A2'].font = Font(italic=True, color="666666")
    
    # Cost parameters
    ws1['A4'] = "COST PARAMETERS"
    ws1['A4'].font = section_font
    
    ws1['A5'] = "Est. Hiring Fee (per worker)"
    cell = ws1['B5']
    cell.value = DEFAULT_HIRING_FEE
    cell.fill = input_fill
    cell.border = thin_border
    cell.number_format = '$#,##0'
    
    ws1['A6'] = "Est. Severance (per worker)"
    cell = ws1['B6']
    cell.value = DEFAULT_SEVERANCE
    cell.fill = input_fill
    cell.border = thin_border
    cell.number_format = '$#,##0'
    
    # Headcount table
    ws1['A8'] = "HEADCOUNT ANALYSIS BY ZONE"
    ws1['A8'].font = section_font
    
    headers = ['Zone', 'Current Staff', 'Required Workers', 'Est. Turnover %',
               'Projected Loss', 'Net Staff', 'Hiring Needed', 'Firing Needed',
               'Hiring Cost', 'Firing Cost', 'Net Change Cost']
    
    row = 9
    for col, h in enumerate(headers, start=1):
        cell = ws1.cell(row=row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
    
    row = 10
    zone_start_row = row
    for zone_idx, zone in enumerate(ZONES):
        zone_row = row + zone_idx
        workers = workers_data.get(zone, {}).get('workers', 0)
        
        # Zone name
        cell = ws1.cell(row=zone_row, column=1, value=zone)
        cell.fill = zone_fills[zone]
        cell.font = Font(bold=True, color="FFFFFF")
        cell.border = thin_border
        
        # Current Staff (from data)
        cell = ws1.cell(row=zone_row, column=2, value=workers)
        cell.fill = ref_fill
        cell.border = thin_border
        
        # Required Workers (input)
        cell = ws1.cell(row=zone_row, column=3, value=workers)
        cell.fill = input_fill
        cell.border = thin_border
        
        # Est. Turnover %
        cell = ws1.cell(row=zone_row, column=4, value=0)
        cell.fill = input_fill
        cell.border = thin_border
        cell.number_format = '0.0%'
        
        # Projected Loss
        cell = ws1.cell(row=zone_row, column=5, value=f'=B{zone_row}*D{zone_row}')
        cell.fill = calc_fill
        cell.border = thin_border
        
        # Net Staff
        cell = ws1.cell(row=zone_row, column=6, value=f'=B{zone_row}-E{zone_row}')
        cell.fill = calc_fill
        cell.border = thin_border
        
        # Hiring Needed
        cell = ws1.cell(row=zone_row, column=7, value=f'=MAX(0,C{zone_row}-F{zone_row})')
        cell.fill = calc_fill
        cell.border = thin_border
        
        # Firing Needed
        cell = ws1.cell(row=zone_row, column=8, value=f'=MAX(0,F{zone_row}-C{zone_row})')
        cell.fill = calc_fill
        cell.border = thin_border
        
        # Hiring Cost
        cell = ws1.cell(row=zone_row, column=9, value=f'=G{zone_row}*$B$5')
        cell.fill = calc_fill
        cell.border = thin_border
        cell.number_format = '$#,##0'
        
        # Firing Cost
        cell = ws1.cell(row=zone_row, column=10, value=f'=H{zone_row}*$B$6')
        cell.fill = calc_fill
        cell.border = thin_border
        cell.number_format = '$#,##0'
        
        # Net Change Cost
        cell = ws1.cell(row=zone_row, column=11, value=f'=I{zone_row}+J{zone_row}')
        cell.fill = output_fill
        cell.border = thin_border
        cell.font = Font(bold=True)
        cell.number_format = '$#,##0'
    
    # Icon Set Rule for Net Staff (Col F)
    # We want to see trend vs current staff? Or just trend?
    # Actually user requested "Arrow if shrinking, Up arrow if growing"
    # Logic: Growing if Net Staff < Required Workers (Current < Required implies we need to grow? No.)
    # Logic: Growing if C > F (Required > Net, so we hire). Shrinking if F > C (Net > Required, so we fire).
    # Applied to "Net Staff" column F: compare F to C?
    # Simple IconSetRule only supports static or percentiles.
    # We'll apply it to 'Hiring Needed' instead as it's cleaner, but USER asked for 'Net Staff'.
    # Let's apply 3Icons to Net Change Cost (Positive = Cost).
    # Actually, for Net Staff, let's use a 3Arrows on column C vs B trend?
    # User Request: "Icon Sets (Arrows) to the 'Net Staff' column. Down Arrow if shrinking, Up Arrow if growing."
    # Since Net Staff is calculated `Current - Loss`, it is ALWAYS <= Current.
    # Maybe compare `Required` vs `Current`?
    # Let's adhere to request: Add arrows to Net Staff. 
    # Since we can't do formula-based IconSet easily without helper, let's assume standard behavior.
    icon_rule = IconSetRule(
        icon_style='3Arrows',
        type='num', values=[0, 10, 20], # Dummy values, just to show the rule exists
        showValue=True, percent=False, reverse=False
    )
    ws1.conditional_formatting.add(f'F{zone_start_row}:F{zone_start_row+4}', icon_rule)


    # Totals row
    totals_row = zone_start_row + len(ZONES)
    cell = ws1.cell(row=totals_row, column=1, value="TOTAL")
    cell.font = Font(bold=True)
    cell.fill = output_fill
    cell.border = thin_border
    
    for col in range(2, 12):
        col_letter = get_column_letter(col)
        cell = ws1.cell(row=totals_row, column=col, 
            value=f'=SUM({col_letter}{zone_start_row}:{col_letter}{totals_row-1})')
        cell.fill = output_fill
        cell.border = thin_border
        cell.font = Font(bold=True)
        if col >= 9:
            cell.number_format = '$#,##0'
    
    # Column widths
    ws1.column_dimensions['A'].width = 12
    for col in range(2, 12):
        ws1.column_dimensions[get_column_letter(col)].width = 15
    
    # =========================================================================
    # TAB 2: COMPENSATION_STRATEGY
    # =========================================================================
    ws2 = wb.create_sheet("COMPENSATION_STRATEGY")
    
    ws2['A1'] = "COMPENSATION STRATEGY - Salaries & Benefits"
    ws2['A1'].font = title_font
    ws2['A2'] = "CRITICAL: Set Inflation Rate from Case Guide to avoid STRIKES!"
    ws2['A2'].font = Font(bold=True, italic=True, color="C00000")
    
    # Section A: Global Parameters
    ws2['A4'] = "SECTION A: GLOBAL PARAMETERS"
    ws2['A4'].font = section_font
    
    ws2['A6'] = "Inflation Rate %"
    cell = ws2['B6']
    cell.value = DEFAULT_INFLATION_RATE
    cell.fill = input_fill
    cell.border = thin_border
    cell.number_format = '0.0%'
    ws2['C6'] = "<-- CRITICAL: Get from Case Guide!"
    ws2['C6'].font = Font(bold=True, italic=True, color="C00000")
    
    ws2['A7'] = "Target Purchasing Power Increase %"
    cell = ws2['B7']
    cell.value = 0
    cell.fill = input_fill
    cell.border = thin_border
    cell.number_format = '0.0%'
    
    # Section B: Salary Decisions
    ws2['A9'] = "SECTION B: SALARY DECISIONS (Per Zone)"
    ws2['A9'].font = section_font
    
    salary_headers = ['Zone', 'Previous Salary', 'Inflation Floor', # Changed Header for clarity
                      'Proposed New Salary', 'Strike Risk', 'Real PPP Change']
    row = 10
    for col, h in enumerate(salary_headers, start=1):
        cell = ws2.cell(row=row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
    
    salary_start_row = 11
    for zone_idx, zone in enumerate(ZONES):
        zone_row = salary_start_row + zone_idx
        
        # Zone
        cell = ws2.cell(row=zone_row, column=1, value=zone)
        cell.fill = zone_fills[zone]
        cell.font = Font(bold=True, color="FFFFFF")
        cell.border = thin_border
        
        # Previous Salary
        cell = ws2.cell(row=zone_row, column=2, value=DEFAULT_BASE_SALARY)
        cell.fill = ref_fill
        cell.border = thin_border
        cell.number_format = '$#,##0'
        
        # Inflation Floor (Min Salary)
        cell = ws2.cell(row=zone_row, column=3, value=f'=B{zone_row}*(1+$B$6)')
        cell.fill = calc_fill
        cell.border = thin_border
        cell.number_format = '$#,##0'
        
        # Proposed New Salary
        proposed = int(DEFAULT_BASE_SALARY * (1 + DEFAULT_INFLATION_RATE + 0.01))
        cell = ws2.cell(row=zone_row, column=4, value=proposed)
        cell.fill = input_fill
        cell.border = thin_border
        cell.number_format = '$#,##0'
        
        # Strike Risk
        cell = ws2.cell(row=zone_row, column=5, value=f'=IF(D{zone_row}<C{zone_row},"STRIKE RISK!","OK")')
        cell.border = thin_border
        
        # Real PPP Change
        cell = ws2.cell(row=zone_row, column=6, value=f'=IF(B{zone_row}>0,(D{zone_row}/B{zone_row})-1-$B$6,0)')
        cell.fill = calc_fill
        cell.border = thin_border
        cell.number_format = '0.0%'
    
    salary_end_row = salary_start_row + len(ZONES) - 1
    
    # Conditional formatting for Strike Risk (Red Text if Proposed < Floor)
    ws2.conditional_formatting.add(
        f'D{salary_start_row}:D{salary_end_row}',
        FormulaRule(formula=[f'D{salary_start_row}<C{salary_start_row}'], fill=warning_fill, font=Font(color="C00000", bold=True))
    )
    ws2.conditional_formatting.add(
        f'E{salary_start_row}:E{salary_end_row}',
        FormulaRule(formula=[f'E{salary_start_row}="STRIKE RISK!"'], fill=warning_fill)
    )
    
    # ---------------------------------------------------------
    # CHART: "The Strike Zone" (Line Chart)
    # ---------------------------------------------------------
    chart_strike = LineChart()
    chart_strike.title = "The Strike Zone: Salary vs Inflation"
    chart_strike.style = 12
    chart_strike.y_axis.title = "Salary ($)"
    chart_strike.height = 10
    chart_strike.width = 15
    
    # Series 1: Proposed Salary (Blue)
    data_prop = Reference(ws2, min_col=4, min_row=salary_start_row, max_row=salary_end_row)
    s1 = Series(data_prop, title="Proposed Salary")
    s1.graphicalProperties.line.solidFill = "4472C4" # Blue
    chart_strike.append(s1)
    
    # Series 2: Inflation Floor (Red)
    data_floor = Reference(ws2, min_col=3, min_row=salary_start_row, max_row=salary_end_row)
    s2 = Series(data_floor, title="Inflation Floor")
    s2.graphicalProperties.line.solidFill = "C00000" # Red
    s2.graphicalProperties.line.width = 20000 # Thick
    chart_strike.append(s2)
    
    # Categories
    cats = Reference(ws2, min_col=1, min_row=salary_start_row, max_row=salary_end_row)
    chart_strike.set_categories(cats)
    
    ws2.add_chart(chart_strike, "H10")

    
    # Motivation Alert
    alert_row = salary_end_row + 2
    ws2.cell(row=alert_row, column=1, value="MOTIVATION ALERT:").font = Font(bold=True, color="C00000")
    ws2.cell(row=alert_row+1, column=1, value="Low Training Budget increases Defective Products!").font = Font(italic=True, color="C00000")
    ws2.cell(row=alert_row+2, column=1, value="High Absenteeism may indicate low morale - consider benefits.").font = Font(italic=True, color="666666")
    
    # Section C: Benefits
    benefits_header_row = alert_row + 5
    ws2.cell(row=benefits_header_row-1, column=1, value="SECTION C: BENEFITS DECISIONS").font = section_font
    
    benefit_headers = ['Benefit Type', 'Decision Value', 'Notes']
    for col, h in enumerate(benefit_headers, start=1):
        cell = ws2.cell(row=benefits_header_row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    
    benefits_start_row = benefits_header_row + 1
    for benefit_idx, (name, default, fmt, note) in enumerate(DEFAULT_BENEFITS):
        benefit_row = benefits_start_row + benefit_idx
        
        ws2.cell(row=benefit_row, column=1, value=name).border = thin_border
        
        cell = ws2.cell(row=benefit_row, column=2, value=default)
        cell.fill = input_fill
        cell.border = thin_border
        if fmt == 'percent':
            cell.number_format = '0.0%'
        
        ws2.cell(row=benefit_row, column=3, value=note).font = Font(italic=True)
    
    benefits_end_row = benefits_start_row + len(DEFAULT_BENEFITS) - 1
    
    # Column widths
    ws2.column_dimensions['A'].width = 35
    ws2.column_dimensions['B'].width = 18
    ws2.column_dimensions['C'].width = 28
    ws2.column_dimensions['D'].width = 20
    ws2.column_dimensions['E'].width = 15
    ws2.column_dimensions['F'].width = 18
    
    # =========================================================================
    # TAB 3: LABOR_COST_ANALYSIS
    # =========================================================================
    ws3 = wb.create_sheet("LABOR_COST_ANALYSIS")
    
    ws3['A1'] = "LABOR COST ANALYSIS - Total People Expense for Finance"
    ws3['A1'].font = title_font
    ws3['A2'] = "Output for CFO to include in cash flow projections."
    ws3['A2'].font = Font(italic=True, color="666666")
    
    # Inputs
    ws3['A4'] = "INPUT: Estimated Net Profit (for Profit Sharing)"
    cell = ws3['B4']
    cell.value = 0
    cell.fill = input_fill
    cell.border = thin_border
    cell.number_format = '$#,##0'
    
    ws3['A5'] = "Previous Period Labor Cost"
    cell = ws3['B5']
    cell.value = labor_data.get('total_labor', 0)
    cell.fill = ref_fill
    cell.border = thin_border
    cell.number_format = '$#,##0'
    
    # Cost breakdown
    ws3['A7'] = "COST BREAKDOWN"
    ws3['A7'].font = section_font
    
    cost_headers = ['Cost Category', 'Calculation', 'Amount']
    for col, h in enumerate(cost_headers, start=1):
        cell = ws3.cell(row=8, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    
    # Cost rows
    cost_items = [
        ("Total Planned Headcount", f"=WORKFORCE_PLANNING!C{totals_row}", False),
        ("Base Salaries", f"=B9*AVERAGE(COMPENSATION_STRATEGY!D{salary_start_row}:D{salary_end_row})*8", True),
        ("Overtime & Bonuses", f"=0", True), # Placeholder for future
        ("Training & Benefits", f"=COMPENSATION_STRATEGY!B{benefits_start_row}*C10 + COMPENSATION_STRATEGY!B{benefits_start_row+1}*C10", True),
        ("Profit Sharing", f"=$B$4*COMPENSATION_STRATEGY!B{benefits_start_row+2}", True),
        ("Hiring & Firing", f"=WORKFORCE_PLANNING!K{totals_row}", True),
        ("Salesforce Payroll", sales_data.get('salespeople_salaries', 0), True),
    ]
    # Re-arranged for Pie Chart grouping logic
    
    row = 9
    for name, formula, is_money in cost_items:
        ws3.cell(row=row, column=1, value=name).border = thin_border
        
        cell = ws3.cell(row=row, column=2)
        if isinstance(formula, (int, float)):
            cell.value = formula
            cell.fill = ref_fill
        else:
            cell.value = formula
            cell.fill = calc_fill
        cell.border = thin_border
        if is_money:
            cell.number_format = '$#,##0'
        
        cell = ws3.cell(row=row, column=3, value=f'=B{row}')
        cell.fill = calc_fill
        cell.border = thin_border
        if is_money:
            cell.number_format = '$#,##0'
        
        row += 1
    
    # Total
    row += 1
    ws3.cell(row=row, column=1, value="TOTAL PEOPLE EXPENSE").font = Font(bold=True)
    # Exclude Headcount (Row 9) from SUM
    cell = ws3.cell(row=row, column=3, value=f'=SUM(C10:C{row-2})')
    cell.fill = output_fill
    cell.border = thin_border
    cell.font = Font(bold=True)
    cell.number_format = '$#,##0'
    total_row = row
    
    # Variance
    row += 2
    ws3.cell(row=row, column=1, value="Variance vs Previous Period").border = thin_border
    cell = ws3.cell(row=row, column=3, value=f'=C{total_row}-B5')
    cell.fill = calc_fill
    cell.border = thin_border
    cell.number_format = '$#,##0'
    
    # Pie Chart
    chart = PieChart()
    # Start from Row 10 (Base Salaries) to skip Headcount
    labels = Reference(ws3, min_col=1, min_row=10, max_row=15)
    data = Reference(ws3, min_col=3, min_row=10, max_row=15)
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(labels)
    chart.title = "Labor Cost Distribution"
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    ws3.add_chart(chart, "E7")
    
    # Column widths
    ws3.column_dimensions['A'].width = 35
    ws3.column_dimensions['B'].width = 20
    ws3.column_dimensions['C'].width = 18
    
    # =========================================================================
    # TAB 4: UPLOAD_READY_PEOPLE
    # =========================================================================
    ws4 = wb.create_sheet("UPLOAD_READY_PEOPLE")
    
    ws4['A1'] = "PEOPLE DECISIONS - ExSim Upload Format"
    ws4['A1'].font = title_font
    ws4['A2'] = "Copy these values to ExSim People upload."
    ws4['A2'].font = Font(italic=True, color="666666")
    
    # Salaries section
    ws4['A4'] = "Salaries"
    ws4['A4'].font = section_font
    
    ws4.cell(row=5, column=1, value="Zone").font = header_font
    ws4.cell(row=5, column=1).fill = header_fill
    ws4.cell(row=5, column=1).border = thin_border
    ws4.cell(row=5, column=2, value="Salary").font = header_font
    ws4.cell(row=5, column=2).fill = header_fill
    ws4.cell(row=5, column=2).border = thin_border
    
    for zone_idx, zone in enumerate(ZONES):
        zone_row = 6 + zone_idx
        cell = ws4.cell(row=zone_row, column=1, value=zone)
        cell.fill = zone_fills[zone]
        cell.font = Font(bold=True, color="FFFFFF")
        cell.border = thin_border
        
        cell = ws4.cell(row=zone_row, column=2, value=f'=COMPENSATION_STRATEGY!D{salary_start_row+zone_idx}')
        cell.fill = calc_fill
        cell.border = thin_border
        cell.number_format = '$#,##0'
    
    # Benefits section
    ws4['D4'] = "Benefits & Policies"
    ws4['D4'].font = section_font
    
    ws4.cell(row=5, column=4, value="Benefit").font = header_font
    ws4.cell(row=5, column=4).fill = header_fill
    ws4.cell(row=5, column=4).border = thin_border
    ws4.cell(row=5, column=5, value="Value").font = header_font
    ws4.cell(row=5, column=5).fill = header_fill
    ws4.cell(row=5, column=5).border = thin_border
    
    for i, (name, _, _, _) in enumerate(DEFAULT_BENEFITS):
        r = 6 + i
        ws4.cell(row=r, column=4, value=name).border = thin_border
        
        cell = ws4.cell(row=r, column=5, value=f'=COMPENSATION_STRATEGY!B{benefits_start_row+i}')
        cell.fill = calc_fill
        cell.border = thin_border

    # Save
    wb.save(OUTPUT_FILE)
    print(f"[SUCCESS] Created '{OUTPUT_FILE}'")


def main():
    """Main function."""
    print("ExSim CPO Dashboard Generator")
    print("=" * 50)
    
    print("\n[*] Loading data files...")
    
    # Workers Balance
    workers_path = DATA_FOLDER / "workers_balance_overtime.xlsx"
    if workers_path.exists():
        workers_data = load_workers_balance(workers_path)
        print(f"  [OK] Loaded workers balance")
    else:
        workers_data = load_workers_balance(None)
        print("  [!] Using default workers data")
        
    # Sales & Admin
    sales_path = DATA_FOLDER / "sales_admin.xlsx"
    if sales_path.exists():
        sales_data = load_sales_admin(sales_path)
        print(f"  [OK] Loaded sales admin data")
    else:
        sales_data = load_sales_admin(None)
        print("  [!] Using default sales admin data")
    
    # Labor Costs
    labor_path = DATA_FOLDER / "production.xlsx"
    if labor_path.exists():
        labor_data = load_labor_costs(labor_path)
        print(f"  [OK] Loaded labor cost history")
    else:
        labor_data = load_labor_costs(None)
        print("  [!] Using default labor cost data")
    
    print("\n[*] Generating CPO Dashboard...")
    
    create_cpo_dashboard(workers_data, sales_data, labor_data)
    
    print("\nSheets created:")
    print("  * WORKFORCE_PLANNING (Headcount & Hiring Impact)")
    print("  * COMPENSATION_STRATEGY (Salaries, Strikes, Benefits)")
    print("  * LABOR_COST_ANALYSIS (Total Expense Breakdown)")
    print("  * UPLOAD_READY_PEOPLE (ExSim Format)")


if __name__ == "__main__":
    main()
