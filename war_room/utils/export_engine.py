"""
ExSim War Room - Export Engine
Generates decision files in exact ExSim format.
"""

import io
import zipfile
import pandas as pd
import streamlit as st
from typing import Dict, Any


ZONES = ['Center', 'West', 'North', 'East', 'South']
FORTNIGHTS = list(range(1, 9))


def generate_marketing_csv(decisions: Dict[str, Any]) -> str:
    """Generate Marketing Decisions CSV in exact ExSim format."""
    lines = []
    lines.append("MARKETING DECISIONS")
    lines.append("")
    lines.append("TV BUDGET")
    lines.append(f"Amount,{decisions.get('tv_budget', 0)}")
    lines.append("")
    lines.append("RADIO BUDGET BY ZONE")
    lines.append("Zone,Amount")
    for zone in ZONES:
        lines.append(f"{zone},{decisions.get('radio', {}).get(zone, 0)}")
    lines.append("")
    lines.append("SALESPEOPLE BY ZONE")
    lines.append("Zone,Count,Salary")
    for zone in ZONES:
        sp = decisions.get('salespeople', {}).get(zone, {})
        lines.append(f"{zone},{sp.get('count', 0)},{sp.get('salary', 0)}")
    
    return '\n'.join(lines)


def generate_people_csv(decisions: Dict[str, Any]) -> str:
    """Generate People Decisions CSV in exact ExSim format."""
    lines = []
    lines.append("Zone,Salary")
    for zone in ZONES:
        lines.append(f"{zone},{decisions.get('salaries', {}).get(zone, 0)}")
    lines.append("")
    lines.append("Benefit,Amount")
    benefits = [
        "Personal days (per person and period)",
        "Budget for additional training (% of payroll)",
        "Health and safety budget (% of payroll)",
        "Union representatives (total of people by company)",
        "Reduction of working hours (hours per period)",
        "Profit sharing (% of profit from previous period)",
        "Health insurance (% of payroll)"
    ]
    for benefit in benefits:
        lines.append(f"{benefit},{decisions.get('benefits', {}).get(benefit, 0)}")
    
    return '\n'.join(lines)


def generate_finance_csv(decisions: Dict[str, Any]) -> str:
    """Generate Finance Decisions CSV in exact ExSim format."""
    lines = []
    
    # Credit Lines
    lines.append("Credit Lines," + ",".join([f"FN{fn}" for fn in FORTNIGHTS]))
    credit = decisions.get('credit_lines', [0]*8)
    lines.append("Amount," + ",".join([str(c) for c in credit]))
    lines.append("")
    
    # Investments
    lines.append("Investments," + ",".join([f"FN{fn}" for fn in FORTNIGHTS]))
    invest = decisions.get('investments', [0]*8)
    lines.append("Amount," + ",".join([str(i) for i in invest]))
    lines.append("")
    
    # Mortgages
    lines.append("Mortgages,Amount,Rate,Payment1,Payment2")
    for loan_num in [1, 2]:
        loan = decisions.get('mortgages', {}).get(f'Loan {loan_num}', {})
        lines.append(f"Loan {loan_num},{loan.get('amount', 0)},{loan.get('rate', 0.08)},{loan.get('payment1', 0)},{loan.get('payment2', 0)}")
    lines.append("")
    
    # Dividends
    lines.append(f"Dividends,{decisions.get('dividends', 0)}")
    
    return '\n'.join(lines)


def generate_procurement_csv(decisions: Dict[str, Any]) -> str:
    """Generate Procurement Decisions CSV in exact ExSim format."""
    lines = []
    headers = ["Parts", "Supplier", "Component"] + [str(fn) for fn in FORTNIGHTS]
    lines.append(",".join(headers))
    
    orders = decisions.get('orders', [])
    for order in orders:
        row = [
            order.get('zone', ''),
            order.get('supplier', ''),
            order.get('component', '')
        ] + [str(order.get(f'fn{fn}', 0)) for fn in FORTNIGHTS]
        lines.append(",".join(row))
    
    return '\n'.join(lines)


def generate_logistics_csv(decisions: Dict[str, Any]) -> str:
    """Generate Logistics Decisions CSV in exact ExSim format."""
    lines = []
    
    # Warehouses
    lines.append("Warehouses")
    lines.append("Zone,Modules")
    for zone in ZONES:
        lines.append(f"{zone},{decisions.get('warehouses', {}).get(zone, 0)}")
    lines.append("")
    
    # Shipments
    lines.append("Shipments")
    lines.append("Fortnight,Origin,Destination,Material,Transport,Quantity")
    shipments = decisions.get('shipments', [])
    for ship in shipments:
        lines.append(f"{ship.get('fortnight', '')},{ship.get('origin', '')},{ship.get('destination', '')},{ship.get('material', 'Electroclean')},{ship.get('transport', 'Train')},{ship.get('quantity', 0)}")
    
    return '\n'.join(lines)


def generate_esg_csv(decisions: Dict[str, Any]) -> str:
    """Generate ESG Decisions CSV in exact ExSim format."""
    lines = []
    lines.append("Component,Amount")
    
    initiatives = [
        ("Solar PV panels", "solar"),
        ("Tree plantation (groups of 80)", "trees"),
        ("Green electricity (%)", "green_electricity"),
        ("CO2 credits purchase (1 period)", "co2_1"),
        ("CO2 credits purchase (2 periods)", "co2_2"),
        ("CO2 credits purchase (3 periods)", "co2_3"),
    ]
    
    for label, key in initiatives:
        lines.append(f"{label},{decisions.get(key, 0)}")
    
    return '\n'.join(lines)


def create_decisions_zip() -> bytes:
    """Create a ZIP file with all decision CSVs."""
    buffer = io.BytesIO()
    
    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        # Marketing
        marketing = st.session_state.get('marketing_decisions', {})
        zf.writestr('Marketing_Decisions.csv', generate_marketing_csv(marketing or {}))
        
        # People
        people = st.session_state.get('people_decisions', {})
        zf.writestr('People_Decisions.csv', generate_people_csv(people or {}))
        
        # Finance
        finance = st.session_state.get('finance_decisions', {})
        zf.writestr('Finance_Decisions.csv', generate_finance_csv(finance or {}))
        
        # Procurement
        procurement = st.session_state.get('procurement_decisions', {})
        zf.writestr('Procurement_Decisions.csv', generate_procurement_csv(procurement or {}))
        
        # Logistics
        logistics = st.session_state.get('logistics_decisions', {})
        zf.writestr('Logistics_Decisions.csv', generate_logistics_csv(logistics or {}))
        
        # ESG
        esg = st.session_state.get('esg_decisions', {})
        zf.writestr('ESG_Decisions.csv', generate_esg_csv(esg or {}))
    
    buffer.seek(0)
    return buffer.getvalue()
