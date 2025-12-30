"""
ExSim Finance Dashboard Final - Financial Control & Liquidity

Integrates cash flow monitoring, income statement projection,
balance sheet health tracking, and debt management.

Required libraries: pandas, openpyxl
"""

import pandas as pd
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import LineChart, BarChart, Reference, Series
import warnings

warnings.filterwarnings('ignore')

# =============================================================================
# CONFIGURATION
# =============================================================================

# Required input files from centralized Reports folder
REQUIRED_FILES = [
    'initial_cash_flow.xlsx',
    'results_and_balance_statements.xlsx',
    'sales_admin_expenses.xlsx',
    'accounts_receivable_payable.xlsx',
    'Finance Decisions.xlsx'
]

# Data source: Primary = Reports folder at project root, Fallback = local /data
REPORTS_FOLDER = Path(__file__).parent.parent / "Reports"
LOCAL_DATA_FOLDER = Path(__file__).parent / "data"

def get_data_path(filename):
    """Get data file path, checking Reports folder first, then local fallback."""
    primary = REPORTS_FOLDER / filename
    fallback = LOCAL_DATA_FOLDER / filename
    if primary.exists():
        return primary
    elif fallback.exists():
        return fallback
    return None

OUTPUT_FILE = "Finance_Dashboard_Final.xlsx"

FORTNIGHTS = list(range(1, 9))  # 1-8

# Default financial parameters
DEFAULT_INTEREST_RATE = 0  # Was 0.08
DEFAULT_TAX_RATE = 0  # Was 0.30


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
    # Handle parentheses as negative
    if cleaned.startswith('(') and cleaned.endswith(')'):
        cleaned = '-' + cleaned[1:-1]
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

def load_initial_cash_flow(filepath):
    """Load initial cash flow for starting position."""
    df = load_excel_file(filepath)
    
    data = {
        'final_cash': 0,
        'tax_payments': 0
    }
    
    if df is None:
        return data
    
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
        
        if 'final cash' in first_val:
            for col_idx in range(1, min(10, len(row))):
                val = parse_numeric(row.iloc[col_idx])
                if val != 0:
                    data['final_cash'] = val
                    break
        
        if 'tax' in first_val and 'payment' in first_val:
            for col_idx in range(1, min(10, len(row))):
                val = parse_numeric(row.iloc[col_idx])
                if val != 0:
                    data['tax_payments'] = abs(val)
                    break
    
    return data


def load_balance_statements(filepath):
    """Load results and balance statements for historical data."""
    df = load_excel_file(filepath)
    
    data = {
        'net_sales': 0,
        'cogs': 0,
        'gross_income': 0,
        'net_profit': 0,
        'total_assets': 0,
        'total_liabilities': 0,
        'equity': 0,
        'retained_earnings': 0,
        'depreciation': 0,
        'gross_margin_pct': 0,
        'net_margin_pct': 0
    }
    
    if df is None:
        return data
    
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
        
        if 'net sales' in first_val or 'revenue' in first_val:
            val = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            if val > 0:
                data['net_sales'] = val
        
        if 'cost of goods sold' in first_val or 'cogs' in first_val:
            val = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            if val != 0:
                data['cogs'] = abs(val)
        
        if 'gross' in first_val and ('income' in first_val or 'profit' in first_val or 'margin' in first_val):
            val = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            if val != 0:
                data['gross_income'] = val
        
        if ('net profit' in first_val or 'net income' in first_val) and 'before' not in first_val:
            val = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            if val != 0:
                data['net_profit'] = val
        
        if 'total assets' in first_val:
            val = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            if val > 0:
                data['total_assets'] = val
        
        if 'total liabilities' in first_val:
            val = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            if val != 0:
                data['total_liabilities'] = abs(val)
        
        if first_val == 'equity' or 'total equity' in first_val:
            val = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            if val != 0:
                data['equity'] = val
        
        if 'retained earnings' in first_val:
            val = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            data['retained_earnings'] = val
        
        if 'depreciation' in first_val:
            val = parse_numeric(row.iloc[1]) if len(row) > 1 else 0
            if val != 0:
                data['depreciation'] = abs(val)
    
    # Calculate margins (with division protection)
    if data['net_sales'] > 0:
        data['gross_margin_pct'] = data['gross_income'] / data['net_sales']
        data['net_margin_pct'] = data['net_profit'] / data['net_sales']
    
    return data


def load_sales_admin_expenses(filepath):
    """Load S&A expenses for overhead estimation."""
    df = load_excel_file(filepath)
    
    data = {'total_sa_expenses': 0}
    
    if df is None:
        return data
    
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
        
        if 'total' in first_val and ('sales' in first_val or 'admin' in first_val or 's&a' in first_val):
            for col_idx in range(1, min(12, len(row))):
                val = parse_numeric(row.iloc[col_idx])
                if val > 0:
                    data['total_sa_expenses'] = val
                    break
    
    return data


def load_receivables_payables(filepath):
    """Load accounts receivable and payable."""
    df = load_excel_file(filepath)
    
    data = {
        'receivables': [0] * 8,
        'payables': [0] * 8
    }
    
    if df is None:
        return data
    
    # Parse by fortnight from cash flow report if available
    for idx, row in df.iterrows():
        first_val = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ''
        
        # Receivables (Cash In)
        if 'receipts' in first_val and 'customer' in first_val:
            for i in range(8):
                if i + 1 < len(row):
                    data['receivables'][i] = parse_numeric(row.iloc[i+1])
        
        # Payables (Cash Out)
        if 'payment' in first_val and 'supplier' in first_val:
            for i in range(8):
                if i + 1 < len(row):
                    data['payables'][i] = abs(parse_numeric(row.iloc[i+1]))
    
    return data


def load_finance_template(filepath):
    """Load finance decisions template."""
    try:
        df = pd.read_excel(filepath, sheet_name='Finance', header=None)
        return {'df': df, 'exists': True}
    except:
        return {'df': None, 'exists': False}


# =============================================================================
# EXCEL GENERATION
# =============================================================================

def create_finance_dashboard(cash_data, balance_data, sa_data, ar_ap_data, template_data):
    """Create the Finance Dashboard."""
    
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
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # =========================================================================
    # TAB 1: LIQUIDITY_MONITOR
    # =========================================================================
    ws1 = wb.active
    ws1.title = "LIQUIDITY_MONITOR"
    
    ws1['A1'] = "LIQUIDITY MONITOR - Cash Flow Engine"
    ws1['A1'].font = title_font
    
    # Section A: Initialization
    ws1['A3'] = "SECTION A: INITIALIZATION (Initial Cash Flow Bridge)"
    ws1['A3'].font = section_font
    
    ws1.cell(row=5, column=1, value="Cash at End of Last Period").border = thin_border
    cell = ws1.cell(row=5, column=2, value=cash_data.get('final_cash', 0))
    cell.border = thin_border
    cell.fill = ref_fill
    cell.number_format = '$#,##0'
    
    ws1.cell(row=6, column=1, value="Less: Tax Payments").border = thin_border
    cell = ws1.cell(row=6, column=2, value=cash_data.get('tax_payments', 0))
    cell.border = thin_border
    cell.fill = input_fill
    cell.number_format = '$#,##0'
    
    ws1.cell(row=7, column=1, value="Less: Dividend Payments").border = thin_border
    cell = ws1.cell(row=7, column=2, value=0)
    cell.border = thin_border
    cell.fill = input_fill
    cell.number_format = '$#,##0'
    
    ws1.cell(row=8, column=1, value="Less: Asset Purchases").border = thin_border
    cell = ws1.cell(row=8, column=2, value=0)
    cell.border = thin_border
    cell.fill = input_fill
    cell.number_format = '$#,##0'
    
    ws1.cell(row=9, column=1, value="STARTING CASH FOR FN1").font = Font(bold=True)
    cell = ws1.cell(row=9, column=2, value="=B5-B6-B7-B8")
    cell.border = thin_border
    cell.fill = output_fill
    cell.font = Font(bold=True)
    cell.number_format = '$#,##0'
    
    # Section B: Operational Cash Flow
    row = 12
    ws1.cell(row=row, column=1, value="SECTION B: OPERATIONAL CASH FLOW").font = section_font
    row += 2
    
    # Headers
    ws1.cell(row=row, column=1, value="Item").font = header_font
    ws1.cell(row=row, column=1).fill = header_fill
    for fn in FORTNIGHTS:
        cell = ws1.cell(row=row, column=1+fn, value=f"FN{fn}")
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
    row += 1
    
    # Opening Cash
    # Note: We need ending_cash_row which is defined later, so we pre-calculate it
    # Layout from open_cash_row (15):
    #   +1 Sales, +2 Procurement, +3 S&A, +4 Receivables, +5 Payables
    #   +6 gap, +7 Section C header, +8 Credit, +9 Invest, +10 Mortgage, +11 Dividends
    #   +12 gap, +13 Section D header, +14 Net Flow, +15 Ending Cash
    ending_cash_offset = 15  # open_cash_row + 15 = ending_cash_row
    
    ws1.cell(row=row, column=1, value="Opening Cash").border = thin_border
    open_cash_row = row
    for fn in FORTNIGHTS:
        if fn == 1:
            # FN1 Opening = Starting Cash from Section A
            cell = ws1.cell(row=row, column=1+fn, value="=$B$9")
        else:
            # FN2+ Opening = Previous FN's Ending Cash
            # Previous column is (1+fn-1) = fn, ending cash is at open_cash_row + ending_cash_offset
            prev_col = get_column_letter(fn)  # Previous fortnight column
            cell = ws1.cell(row=row, column=1+fn, value=f"={prev_col}${row + ending_cash_offset}")
        cell.border = thin_border
        cell.fill = ref_fill
        cell.number_format = '$#,##0'
    row += 1
    
    # Sales Receipts
    ws1.cell(row=row, column=1, value="Sales Receipts (Est.)").border = thin_border
    sales_row = row
    for fn in FORTNIGHTS:
        cell = ws1.cell(row=row, column=1+fn, value=100000)
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '$#,##0'
    row += 1
    
    # Procurement Spend
    ws1.cell(row=row, column=1, value="Procurement Spend (Est.)").border = thin_border
    for fn in FORTNIGHTS:
        cell = ws1.cell(row=row, column=1+fn, value=0)
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '$#,##0'
    procurement_row = row
    row += 1
    
    # Fixed Overhead (S&A)
    sa_per_fn = sa_data.get('total_sa_expenses', 0) / 8
    ws1.cell(row=row, column=1, value="Fixed Overhead (S&A)").border = thin_border
    for fn in FORTNIGHTS:
        cell = ws1.cell(row=row, column=1+fn, value=int(sa_per_fn))
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '$#,##0'
    sa_row = row
    row += 1
    
    # Receivables
    ws1.cell(row=row, column=1, value="Receivables (Hard)").border = thin_border
    recv_row = row
    for fn in FORTNIGHTS:
        cell = ws1.cell(row=row, column=1+fn, value=ar_ap_data['receivables'][fn-1])
        cell.border = thin_border
        cell.fill = ref_fill
        cell.number_format = '$#,##0'
    row += 1
    
    # Payables
    ws1.cell(row=row, column=1, value="Payables (Hard)").border = thin_border
    pay_row = row
    for fn in FORTNIGHTS:
        cell = ws1.cell(row=row, column=1+fn, value=ar_ap_data['payables'][fn-1])
        cell.border = thin_border
        cell.fill = ref_fill
        cell.number_format = '$#,##0'
    row += 2
    
    # Section C: Financing Decisions
    ws1.cell(row=row, column=1, value="SECTION C: FINANCING DECISIONS").font = section_font
    row += 1
    
    ws1.cell(row=row, column=1, value="Change in Credit Line (+/-)").border = thin_border
    credit_row = row
    for fn in FORTNIGHTS:
        cell = ws1.cell(row=row, column=1+fn, value=0)
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '$#,##0'
    row += 1
    
    ws1.cell(row=row, column=1, value="Change in Investments (+/-)").border = thin_border
    invest_row = row
    for fn in FORTNIGHTS:
        cell = ws1.cell(row=row, column=1+fn, value=0)
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '$#,##0'
    row += 1
    
    ws1.cell(row=row, column=1, value="New Mortgage Inflow").border = thin_border
    mortgage_row = row
    for fn in FORTNIGHTS:
        cell = ws1.cell(row=row, column=1+fn, value=0)
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '$#,##0'
    row += 1
    
    ws1.cell(row=row, column=1, value="Dividends Paid").border = thin_border
    dividend_row = row
    for fn in FORTNIGHTS:
        cell = ws1.cell(row=row, column=1+fn, value=0)
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '$#,##0'
    row += 2
    
    # Section D: The Balance
    ws1.cell(row=row, column=1, value="SECTION D: CASH BALANCE").font = section_font
    row += 1
    
    ws1.cell(row=row, column=1, value="Net Cash Flow").border = thin_border
    net_flow_row = row
    for fn in FORTNIGHTS:
        col = 1 + fn
        col_letter = get_column_letter(col)
        # Net = Sales + Receivables + Credit + Mortgage - Procurement - S&A - Payables - Investments - Dividends
        cell = ws1.cell(row=row, column=col,
            value=f"={col_letter}{sales_row}+{col_letter}{recv_row}+{col_letter}{credit_row}+{col_letter}{mortgage_row}-{col_letter}{procurement_row}-{col_letter}{sa_row}-{col_letter}{pay_row}-{col_letter}{invest_row}-{col_letter}{dividend_row}")
        cell.border = thin_border
        cell.fill = calc_fill
        cell.number_format = '$#,##0'
    row += 1
    
    ws1.cell(row=row, column=1, value="ENDING CASH BALANCE").font = Font(bold=True)
    ending_cash_row = row
    for fn in FORTNIGHTS:
        col = 1 + fn
        col_letter = get_column_letter(col)
        cell = ws1.cell(row=row, column=col,
            value=f"={col_letter}{open_cash_row}+{col_letter}{net_flow_row}")
        cell.border = thin_border
        cell.fill = output_fill
        cell.font = Font(bold=True)
        cell.number_format = '$#,##0'
    row += 1
    
    ws1.cell(row=row, column=1, value="Solvency Check").border = thin_border
    for fn in FORTNIGHTS:
        col = 1 + fn
        col_letter = get_column_letter(col)
        cell = ws1.cell(row=row, column=col,
            value=f'=IF({col_letter}{ending_cash_row}<0,"INSOLVENT!",IF({col_letter}{ending_cash_row}>200000,"Excess Cash","OK"))')
        cell.border = thin_border
    
    # Add conditional formatting
    ws1.conditional_formatting.add(
        f'B{ending_cash_row}:I{ending_cash_row}',
        FormulaRule(formula=[f'B{ending_cash_row}<0'], fill=red_fill)
    )
    ws1.conditional_formatting.add(
        f'B{ending_cash_row}:I{ending_cash_row}',
        FormulaRule(formula=[f'B{ending_cash_row}>200000'], fill=green_fill)
    )
    
    # ---------------------------------------------------------
    # CHART: Liquidity Forecast (Line Chart)
    # ---------------------------------------------------------
    chart_liq = LineChart()
    chart_liq.title = "Liquidity Forecast (Cash Balance)"
    chart_liq.style = 12
    chart_liq.y_axis.title = "Cash Balance ($)"
    chart_liq.x_axis.title = "Fortnight"
    chart_liq.height = 10
    chart_liq.width = 15

    # Data: Ending Cash (Row ending_cash_row, Cols B-I)
    # Fortnights are columns 2 to 9
    data_cash = Reference(ws1, min_col=2, min_row=ending_cash_row, max_col=9)
    s_cash = Series(data_cash, title="Ending Cash")
    chart_liq.append(s_cash)

    # Categories (FN1..FN8 headers at Row 14)
    cats_liq = Reference(ws1, min_col=2, min_row=14, max_col=9)
    chart_liq.set_categories(cats_liq)

    ws1.add_chart(chart_liq, "H2")
    
    # Column widths
    ws1.column_dimensions['A'].width = 28
    for col in range(2, 10):
        ws1.column_dimensions[get_column_letter(col)].width = 14
    
    # =========================================================================
    # TAB 2: PROFIT_CONTROL
    # =========================================================================
    ws2 = wb.create_sheet("PROFIT_CONTROL")
    
    ws2['A1'] = "PROFIT CONTROL - Income Statement Forecast vs Actuals"
    ws2['A1'].font = title_font
    
    ws2['A3'] = "HISTORICAL MARGINS (From results_and_balance_statements)"
    ws2['A3'].font = section_font
    
    ws2.cell(row=5, column=1, value="Historical Gross Margin %").border = thin_border
    cell = ws2.cell(row=5, column=2, value=balance_data.get('gross_margin_pct', 0))
    cell.border = thin_border
    cell.fill = ref_fill
    cell.number_format = '0.0%'
    
    ws2.cell(row=6, column=1, value="Historical Net Margin %").border = thin_border
    cell = ws2.cell(row=6, column=2, value=balance_data.get('net_margin_pct', 0))
    cell.border = thin_border
    cell.fill = ref_fill
    cell.number_format = '0.0%'
    
    # Income Statement Comparison
    row = 9
    ws2.cell(row=row, column=1, value="INCOME STATEMENT COMPARISON").font = section_font
    row += 1
    
    headers = ['Line Item', 'Last Round Actuals', 'This Round Projected', 'Variance %']
    for col, h in enumerate(headers, start=1):
        cell = ws2.cell(row=row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    row += 1
    
    # Revenue
    ws2.cell(row=row, column=1, value="Net Sales / Revenue").border = thin_border
    ws2.cell(row=row, column=2, value=balance_data.get('net_sales', 0)).border = thin_border
    ws2['B' + str(row)].fill = ref_fill
    ws2['B' + str(row)].number_format = '$#,##0'
    cell = ws2.cell(row=row, column=3, value=balance_data.get('net_sales', 0))
    cell.border = thin_border
    cell.fill = input_fill
    cell.number_format = '$#,##0'
    cell = ws2.cell(row=row, column=4, value=f"=IF(B{row}>0,(C{row}-B{row})/B{row},0)")
    cell.border = thin_border
    cell.fill = calc_fill
    cell.number_format = '0.0%'
    revenue_row = row
    row += 1
    
    # COGS
    ws2.cell(row=row, column=1, value="Cost of Goods Sold").border = thin_border
    ws2.cell(row=row, column=2, value=balance_data.get('cogs', 0)).border = thin_border
    ws2['B' + str(row)].fill = ref_fill
    ws2['B' + str(row)].number_format = '$#,##0'
    cell = ws2.cell(row=row, column=3, value=f"=C{revenue_row}*(1-$B$5)")
    cell.border = thin_border
    cell.fill = calc_fill
    cell.number_format = '$#,##0'
    cell = ws2.cell(row=row, column=4, value=f"=IF(B{row}>0,(C{row}-B{row})/B{row},0)")
    cell.border = thin_border
    cell.fill = calc_fill
    cell.number_format = '0.0%'
    cogs_row = row
    row += 1
    
    # Gross Margin
    ws2.cell(row=row, column=1, value="Gross Margin").border = thin_border
    ws2.cell(row=row, column=2, value=balance_data.get('gross_income', 0)).border = thin_border
    ws2['B' + str(row)].fill = ref_fill
    ws2['B' + str(row)].number_format = '$#,##0'
    cell = ws2.cell(row=row, column=3, value=f"=C{revenue_row}-C{cogs_row}")
    cell.border = thin_border
    cell.fill = calc_fill
    cell.number_format = '$#,##0'
    cell = ws2.cell(row=row, column=4, value=f"=IF(B{row}>0,(C{row}-B{row})/B{row},0)")
    cell.border = thin_border
    cell.fill = calc_fill
    cell.number_format = '0.0%'
    gross_margin_row = row
    row += 1
    
    # S&A Expenses
    ws2.cell(row=row, column=1, value="S&A Expenses").border = thin_border
    ws2.cell(row=row, column=2, value=sa_data.get('total_sa_expenses', 0)).border = thin_border
    ws2['B' + str(row)].fill = ref_fill
    ws2['B' + str(row)].number_format = '$#,##0'
    cell = ws2.cell(row=row, column=3, value=sa_data.get('total_sa_expenses', 0))
    cell.border = thin_border
    cell.fill = input_fill
    cell.number_format = '$#,##0'
    cell = ws2.cell(row=row, column=4, value=f"=IF(B{row}>0,(C{row}-B{row})/B{row},0)")
    cell.border = thin_border
    cell.fill = calc_fill
    cell.number_format = '0.0%'
    sa_expense_row = row
    row += 1
    
    # Depreciation
    ws2.cell(row=row, column=1, value="Depreciation").border = thin_border
    ws2.cell(row=row, column=2, value=balance_data.get('depreciation', 0)).border = thin_border
    ws2['B' + str(row)].fill = ref_fill
    ws2['B' + str(row)].number_format = '$#,##0'
    cell = ws2.cell(row=row, column=3, value=balance_data.get('depreciation', 0))
    cell.border = thin_border
    cell.fill = input_fill
    cell.number_format = '$#,##0'
    cell = ws2.cell(row=row, column=4, value=f"=IF(B{row}>0,(C{row}-B{row})/B{row},0)")
    cell.border = thin_border
    cell.fill = calc_fill
    cell.number_format = '0.0%'
    depreciation_row = row
    row += 1
    
    # Interest Expense
    ws2.cell(row=row, column=1, value="Interest Expense").border = thin_border
    ws2.cell(row=row, column=2, value=0).border = thin_border
    ws2['B' + str(row)].fill = ref_fill
    ws2['B' + str(row)].number_format = '$#,##0'
    cell = ws2.cell(row=row, column=3, value=0)
    cell.border = thin_border
    cell.fill = input_fill
    cell.number_format = '$#,##0'
    interest_row = row
    row += 1
    
    # Net Income
    ws2.cell(row=row, column=1, value="EST. NET INCOME").font = Font(bold=True)
    ws2.cell(row=row, column=2, value=balance_data.get('net_profit', 0)).border = thin_border
    ws2['B' + str(row)].fill = ref_fill
    ws2['B' + str(row)].font = Font(bold=True)
    ws2['B' + str(row)].number_format = '$#,##0'
    cell = ws2.cell(row=row, column=3, 
        value=f"=C{gross_margin_row}-C{sa_expense_row}-C{depreciation_row}-C{interest_row}")
    cell.border = thin_border
    cell.fill = output_fill
    cell.font = Font(bold=True)
    cell.number_format = '$#,##0'
    cell = ws2.cell(row=row, column=4, value=f"=IF(B{row}>0,(C{row}-B{row})/B{row},0)")
    cell.border = thin_border
    cell.fill = calc_fill
    cell.number_format = '0.0%'
    net_income_row = row
    row += 2
    
    # Accuracy Check
    ws2.cell(row=row, column=1, value="ACCURACY CHECK").font = section_font
    row += 1
    
    ws2.cell(row=row, column=1, value="Projected Net Margin").border = thin_border
    cell = ws2.cell(row=row, column=2, value=f"=IF(C{revenue_row}>0,C{net_income_row}/C{revenue_row},0)")
    cell.border = thin_border
    cell.fill = calc_fill
    cell.number_format = '0.0%'
    proj_margin_row = row
    row += 1
    
    ws2.cell(row=row, column=1, value="Profit Realism Flag").border = thin_border
    cell = ws2.cell(row=row, column=2, 
        value=f'=IF(B{proj_margin_row}>$B$6+0.05,"WARNING: Unrealistic profit jump!","Projection OK")')
    cell.border = thin_border
    cell.font = Font(bold=True)
    
    # Add conditional formatting for net income (Text Color instead of Fill)
    ws2.conditional_formatting.add(
        f'C{net_income_row}',
        FormulaRule(formula=[f'C{net_income_row}<0'], font=Font(color="C00000", bold=True))
    )
    ws2.conditional_formatting.add(
        f'C{net_income_row}',
        FormulaRule(formula=[f'C{net_income_row}>0'], font=Font(color="00B050", bold=True))
    )
    
    # Variance % Formatting (Yellow if > 10% deviation)
    # Variance is in Column 4 (D)
    yellow_fill = PatternFill(start_color="FFEB84", end_color="FFEB84", fill_type="solid")
    ws2.conditional_formatting.add(
        f'D12:D{net_income_row}',
        FormulaRule(formula=[f'ABS(D12)>0.1'], fill=yellow_fill)
    )
    
    # Column widths
    ws2.column_dimensions['A'].width = 28
    ws2.column_dimensions['B'].width = 20
    ws2.column_dimensions['C'].width = 22
    ws2.column_dimensions['D'].width = 14
    
    # =========================================================================
    # TAB 3: BALANCE_SHEET_HEALTH
    # =========================================================================
    ws3 = wb.create_sheet("BALANCE_SHEET_HEALTH")
    
    ws3['A1'] = "BALANCE SHEET HEALTH - Solvency & Debt Control"
    ws3['A1'].font = title_font
    
    ws3['A3'] = "CURRENT POSITION (From results_and_balance_statements)"
    ws3['A3'].font = section_font
    
    ws3.cell(row=5, column=1, value="Total Assets").border = thin_border
    cell = ws3.cell(row=5, column=2, value=balance_data.get('total_assets', 0))
    cell.border = thin_border
    cell.fill = ref_fill
    cell.number_format = '$#,##0'
    
    ws3.cell(row=6, column=1, value="Total Liabilities").border = thin_border
    cell = ws3.cell(row=6, column=2, value=balance_data.get('total_liabilities', 0))
    cell.border = thin_border
    cell.fill = ref_fill
    cell.number_format = '$#,##0'
    
    ws3.cell(row=7, column=1, value="Total Equity").border = thin_border
    cell = ws3.cell(row=7, column=2, value=balance_data.get('equity', 0))
    cell.border = thin_border
    cell.fill = ref_fill
    cell.number_format = '$#,##0'
    
    ws3.cell(row=8, column=1, value="Retained Earnings").border = thin_border
    cell = ws3.cell(row=8, column=2, value=balance_data.get('retained_earnings', 0))
    cell.border = thin_border
    cell.fill = ref_fill
    cell.number_format = '$#,##0'
    
    row = 11
    ws3.cell(row=row, column=1, value="DEBT ANALYSIS").font = section_font
    row += 2
    
    ws3.cell(row=row, column=1, value="Current Debt Ratio").border = thin_border
    cell = ws3.cell(row=row, column=2, value="=B6/B5")
    cell.border = thin_border
    cell.fill = calc_fill
    cell.number_format = '0.0%'
    current_debt_row = row
    row += 2
    
    ws3.cell(row=row, column=1, value="Projected New Credit Lines").border = thin_border
    cell = ws3.cell(row=row, column=2, value=0)
    cell.border = thin_border
    cell.fill = input_fill
    cell.number_format = '$#,##0'
    new_credit_row = row
    row += 1
    
    ws3.cell(row=row, column=1, value="Projected New Mortgages").border = thin_border
    cell = ws3.cell(row=row, column=2, value=0)
    cell.border = thin_border
    cell.fill = input_fill
    cell.number_format = '$#,##0'
    new_mortgage_row = row
    row += 1
    
    ws3.cell(row=row, column=1, value="Total New Debt").border = thin_border
    cell = ws3.cell(row=row, column=2, value=f"=B{new_credit_row}+B{new_mortgage_row}")
    cell.border = thin_border
    cell.fill = calc_fill
    cell.number_format = '$#,##0'
    total_new_debt_row = row
    row += 2
    
    ws3.cell(row=row, column=1, value="Est. Post-Decision Debt Ratio").border = thin_border
    cell = ws3.cell(row=row, column=2, value=f"=(B6+B{total_new_debt_row})/B5")
    cell.border = thin_border
    cell.fill = output_fill
    cell.font = Font(bold=True)
    cell.number_format = '0.0%'
    post_debt_row = row
    row += 2
    
    # Warning Flags
    ws3.cell(row=row, column=1, value="WARNING FLAGS").font = section_font
    row += 1
    
    ws3.cell(row=row, column=1, value="Debt Level Check").border = thin_border
    cell = ws3.cell(row=row, column=2, 
        value=f'=IF(B{post_debt_row}>0.6,"CRITICAL: Debt too high. Credit Rating Risk.","Health OK")')
    cell.border = thin_border
    cell.font = Font(bold=True)
    row += 1
    
    cell = ws3.cell(row=row, column=2, 
        value='=IF(B8<0,"CRITICAL: Equity Erosion. Retained earnings negative.","Equity OK")')
    cell.border = thin_border
    cell.font = Font(bold=True)
    
    # ---------------------------------------------------------
    # CHART: Solvency Gauge (Bar Chart with Limit Line)
    # ---------------------------------------------------------
    
    # Helper Data for Chart (Hidden Columns H-J)
    ws3['H12'] = "Metric"
    ws3['I12'] = "Ratio"
    ws3['J12'] = "Limit"
    
    ws3['H13'] = "Current"
    ws3['I13'] = f"=B{current_debt_row}"
    ws3['J13'] = 0.6
    
    ws3['H14'] = "Post-Decision"
    ws3['I14'] = f"=B{post_debt_row}"
    ws3['J14'] = 0.6
    
    # Chart
    c_bar = BarChart()
    c_bar.title = "Solvency Gauge (Debt Ratio)"
    c_bar.style = 10
    c_bar.height = 10
    c_bar.width = 10
    
    data_bar = Reference(ws3, min_col=9, min_row=12, max_row=14) # Includes header
    c_bar.add_data(data_bar, titles_from_data=True)
    c_bar.set_categories(Reference(ws3, min_col=8, min_row=13, max_row=14))
    
    c_line = LineChart()
    data_line = Reference(ws3, min_col=10, min_row=12, max_row=14) # Includes header
    c_line.add_data(data_line, titles_from_data=True)
    c_line.y_axis.axId = 200 # Separate axis? No, same axis for comparison.
    # Actually if we want same axis we don't set axId or we set it to same.
    # By default openpyxl combo puts them on same axis if not specified otherwise usually.
    
    # Style line
    s_line = c_line.series[0]
    s_line.graphicalProperties.line.solidFill = "FF0000"
    s_line.graphicalProperties.line.width = 20000
    
    c_bar += c_line
    
    ws3.add_chart(c_bar, "D12")
    
    # Column widths
    ws3.column_dimensions['A'].width = 30
    ws3.column_dimensions['B'].width = 50
    
    # =========================================================================
    # TAB 4: DEBT_MANAGER
    # =========================================================================
    ws4 = wb.create_sheet("DEBT_MANAGER")
    
    ws4['A1'] = "DEBT MANAGER - Mortgage Calculator"
    ws4['A1'].font = title_font
    
    ws4['A3'] = "MORTGAGE BLOCK"
    ws4['A3'].font = section_font
    
    mortgage_headers = ['Loan #', 'Amount', 'Interest Rate', 'Payment Period 1', 'Payment Period 2', 'Total Payments']
    for col, h in enumerate(mortgage_headers, start=1):
        cell = ws4.cell(row=5, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    
    # Pre-fill 3 loan rows
    for loan in range(1, 4):
        row = 5 + loan
        ws4.cell(row=row, column=1, value=f"Loan {loan}").border = thin_border
        
        cell = ws4.cell(row=row, column=2, value=0)
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '$#,##0'
        
        cell = ws4.cell(row=row, column=3, value=0)
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '0.0%'
        
        cell = ws4.cell(row=row, column=4, value=0)
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '$#,##0'
        
        cell = ws4.cell(row=row, column=5, value=0)
        cell.border = thin_border
        cell.fill = input_fill
        cell.number_format = '$#,##0'
        
        cell = ws4.cell(row=row, column=6, value=f"=D{row}+E{row}")
        cell.border = thin_border
        cell.fill = calc_fill
        cell.number_format = '$#,##0'
    
    # Totals
    row = 10
    ws4.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
    ws4.cell(row=row, column=2, value="=SUM(B6:B8)").fill = calc_fill
    ws4.cell(row=row, column=6, value="=SUM(F6:F8)").fill = calc_fill
    
    # Column widths
    for col in range(1, 7):
        ws4.column_dimensions[get_column_letter(col)].width = 18
    
    # =========================================================================
    # TAB 5: UPLOAD_READY_FINANCE
    # =========================================================================
    ws5 = wb.create_sheet("UPLOAD_READY_FINANCE")
    
    ws5['A1'] = "FINANCE DECISIONS - ExSim Upload Format"
    ws5['A1'].font = title_font
    ws5['A2'] = "Copy these values to ExSim Finance upload"
    ws5['A2'].font = Font(italic=True, color="666666")
    
    # Credit Lines
    ws5['A4'] = "Credit Lines"
    ws5['A4'].font = section_font
    
    credit_headers = ['FN1', 'FN2', 'FN3', 'FN4', 'FN5', 'FN6', 'FN7', 'FN8']
    for col, h in enumerate(credit_headers, start=2):
        cell = ws5.cell(row=5, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    
    ws5.cell(row=6, column=1, value="Amount").border = thin_border
    for fn in FORTNIGHTS:
        cell = ws5.cell(row=6, column=1+fn, value=f"=LIQUIDITY_MONITOR!{get_column_letter(1+fn)}{credit_row}")
        cell.border = thin_border
        cell.fill = calc_fill
        cell.number_format = '$#,##0'
    
    # Investments
    row = 9
    ws5.cell(row=row, column=1, value="Investments").font = section_font
    row += 1
    
    for col, h in enumerate(credit_headers, start=2):
        cell = ws5.cell(row=row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    row += 1
    
    ws5.cell(row=row, column=1, value="Amount").border = thin_border
    for fn in FORTNIGHTS:
        cell = ws5.cell(row=row, column=1+fn, value=f"=LIQUIDITY_MONITOR!{get_column_letter(1+fn)}{invest_row}")
        cell.border = thin_border
        cell.fill = calc_fill
        cell.number_format = '$#,##0'
    
    # Mortgages
    row += 3
    ws5.cell(row=row, column=1, value="Mortgages").font = section_font
    row += 1
    
    mortgage_upload_headers = ['Loan', 'Amount', 'Payment 1', 'Payment 2']
    for col, h in enumerate(mortgage_upload_headers, start=1):
        cell = ws5.cell(row=row, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
    row += 1
    
    for loan in range(1, 4):
        ws5.cell(row=row, column=1, value=f"Loan {loan}").border = thin_border
        ws5.cell(row=row, column=2, value=f"=DEBT_MANAGER!B{5+loan}").border = thin_border
        ws5.cell(row=row, column=3, value=f"=DEBT_MANAGER!D{5+loan}").border = thin_border
        ws5.cell(row=row, column=4, value=f"=DEBT_MANAGER!E{5+loan}").border = thin_border
        row += 1
    
    # Dividends
    row += 2
    ws5.cell(row=row, column=1, value="Dividends").font = section_font
    row += 1
    
    # Link to per-fortnight dividends row
    for fn in FORTNIGHTS:
        cell = ws5.cell(row=row, column=1+fn, value=f"=LIQUIDITY_MONITOR!{get_column_letter(1+fn)}{dividend_row}")
        cell.border = thin_border
        cell.fill = calc_fill
        cell.number_format = '$#,##0'
    
    # Column widths
    for col in range(1, 10):
        ws5.column_dimensions[get_column_letter(col)].width = 14
    
    # Save
    wb.save(OUTPUT_FILE)
    print(f"[SUCCESS] Created '{OUTPUT_FILE}'")


def main():
    """Main function."""
    print("ExSim Finance Dashboard Final Generator")
    print("=" * 50)
    
    print("\n[*] Loading data files...")
    print(f"    Primary source: {REPORTS_FOLDER}")
    print(f"    Fallback source: {LOCAL_DATA_FOLDER}")
    
    # Initial Cash Flow
    cash_path = get_data_path("initial_cash_flow.xlsx")
    if cash_path:
        cash_data = load_initial_cash_flow(cash_path)
        print(f"  [OK] Loaded initial cash from {cash_path.parent.name}/")
    else:
        cash_data = load_initial_cash_flow(None)
        print("  [!] Using default cash data")
    
    # Balance Statements
    balance_path = get_data_path("results_and_balance_statements.xlsx")
    if balance_path:
        balance_data = load_balance_statements(balance_path)
        print(f"  [OK] Loaded balance data")
    else:
        balance_data = load_balance_statements(None)
        print("  [!] Using default balance data")
    
    # S&A Expenses
    sa_path = get_data_path("sales_admin_expenses.xlsx")
    if sa_path:
        sa_data = load_sales_admin_expenses(sa_path)
        print(f"  [OK] Loaded S&A expenses")
    else:
        sa_data = load_sales_admin_expenses(None)
        print("  [!] Using default S&A data")
    
    # Receivables/Payables
    ar_ap_path = get_data_path("accounts_receivable_payable.xlsx")
    if ar_ap_path:
        ar_ap_data = load_receivables_payables(ar_ap_path)
        print(f"  [OK] Loaded AR/AP data")
    else:
        ar_ap_data = load_receivables_payables(None)
        print("  [!] Using default AR/AP data")
    
    # Template
    template_path = get_data_path("Finance Decisions.xlsx")
    template_data = load_finance_template(template_path)
    if template_data['exists']:
        print(f"  [OK] Loaded finance template")
    else:
        print("  [!] Using default template layout")
    
    print("\n[*] Generating Finance Dashboard...")
    
    create_finance_dashboard(cash_data, balance_data, sa_data, ar_ap_data, template_data)
    
    print("\nSheets created:")
    print("  * LIQUIDITY_MONITOR (Cash Flow Engine)")
    print("  * PROFIT_CONTROL (Income Statement Projection)")
    print("  * BALANCE_SHEET_HEALTH (Solvency & Debt)")
    print("  * DEBT_MANAGER (Mortgage Calculator)")
    print("  * UPLOAD_READY_FINANCE (ExSim Format)")


if __name__ == "__main__":
    main()
