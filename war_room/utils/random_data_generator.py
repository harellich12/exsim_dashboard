"""
ExSim War Room - Internal Random Data Generator

Generates randomized but realistic data directly into session state
for live UI testing without requiring file I/O.

This is an internal adaptation of test_data/generate_mock_data.py
designed for in-app generation with custom seeds.
"""

import numpy as np
import pandas as pd
from io import BytesIO


# =============================================================================
# CONFIGURATION
# =============================================================================

ZONES = ["Center", "West", "North", "East", "South"]
FORTNIGHTS = 8

# Realistic bounds for generated values
BOUNDS = {
    "workers_per_zone": (50, 300),
    "absenteeism_rate": (0.00, 0.05),
    "overtime_pct": (0.00, 0.15),
    "salespeople": (20, 80),
    "salary_per_person": (500, 1500),
    "tv_spots": (10, 100),
    "radio_spots": (50, 500),
    "tv_cost_per_spot": (2500, 4000),
    "radio_cost_per_spot": (200, 400),
    "inventory_capacity": (1000, 5000),
    "production_units": (500, 2000),
    "unit_price": (50, 150),
    "discount_pct": (5, 15),
    "cash_balance": (20000, 200000),
    "receivables": (50000, 500000),
    "payables": (10000, 200000),
    "depreciation": (20000, 100000),
    "mortgage": (100000, 800000),
    "machines_per_zone": (0, 10),
    "transport_cost_per_unit": (5, 25),
}


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def format_currency(value):
    """Format number as currency string."""
    if value < 0:
        return f"$-{abs(value):,.0f}"
    return f"${value:,.0f}"


def format_percent(value):
    """Format number as percentage string."""
    return f"{value:.1%}"


# =============================================================================
# DATA GENERATORS - Return structured dicts for session state
# =============================================================================

def generate_market_data(rng):
    """Generate market_data structure for CMO dashboard.
    
    Must match structure expected by data_loader.load_market_report:
    - by_segment: dict with 'High' and 'Low' keys
    - Each segment has zone-level data with my_market_share, my_awareness, etc.
    - zones: aggregate zone data
    """
    SEGMENTS = ['High', 'Low']
    
    # Market data structure matching data_loader.py
    market_data = {
        'by_segment': {seg: {zone: {
            'my_market_share': rng.uniform(10, 35) if zone in ZONES[:3] else 0,
            'my_awareness': rng.uniform(40, 85) if zone in ZONES[:3] else 0,
            'my_attractiveness': rng.uniform(50, 90) if zone in ZONES[:3] else 0,
            'my_price': rng.integers(60, 150) if zone in ZONES[:3] else 0,
            'comp_avg_awareness': rng.uniform(50, 80) if zone in ZONES[:3] else 0,
            'comp_avg_price': rng.integers(70, 140) if zone in ZONES[:3] else 0
        } for zone in ZONES} for seg in SEGMENTS},
        'zones': {zone: {
            'my_price': rng.integers(70, 130) if zone in ZONES[:3] else 0,
            'comp_avg_price': rng.integers(75, 125) if zone in ZONES[:3] else 0,
            'my_awareness': rng.uniform(50, 85) if zone in ZONES[:3] else 0,
            'my_attractiveness': rng.uniform(55, 90) if zone in ZONES[:3] else 0,
            'my_market_share': rng.uniform(12, 30) if zone in ZONES[:3] else 0,
            'comp_avg_awareness': rng.uniform(55, 75) if zone in ZONES[:3] else 0
        } for zone in ZONES},
        'segments': SEGMENTS,
        'raw_df': None
    }
    
    return market_data


def generate_workers_data(rng):
    """Generate workers_data structure for CPO dashboard.
    
    Must match structure expected by data_loader.load_workers_balance:
    - zones: dict with zone names as keys
    - Each zone has 'workers' (count) and 'salary' (hourly rate)
    """
    workers_data = {
        'zones': {},
        'raw_df': None
    }
    
    for zone in ZONES:
        is_active = zone in ZONES[:3]
        workers = rng.integers(*BOUNDS["workers_per_zone"]) if is_active else 0
        salary = rng.integers(20, 40) if is_active else 25  # Hourly rate
        
        workers_data['zones'][zone] = {
            'workers': int(workers),
            'salary': float(salary),
        }
    
    return workers_data


def generate_materials_data(rng):
    """Generate materials_data structure for Purchasing dashboard.
    
    Must match structure expected by data_loader.load_raw_materials:
    - parts: dict with part names as keys
    - Each part has stock, cost, final_inventory
    """
    parts_list = ['Part A', 'Part B', 'Piece 1', 'Piece 2']
    
    materials_data = {
        'parts': {},
        'raw_df': None
    }
    
    for part in parts_list:
        stock = rng.integers(2000, 8000)
        cost = rng.uniform(2, 50)
        final = rng.integers(1500, 7000)
        
        materials_data['parts'][part] = {
            'stock': stock,
            'cost': cost,
            'final_inventory': final,
            'initial_inventory': stock,
            'received': rng.integers(1000, 5000),
            'consumed': rng.integers(2500, 6000),
        }
    
    return materials_data


def generate_finished_goods_data(rng):
    """Generate finished_goods_data structure for Logistics and CMO dashboards.
    
    Must match structure expected by data_loader.load_finished_goods:
    - zones: dict with zone names, each containing 'inventory', 'capacity', 'final'
    - is_stockout: boolean indicating if any zone has stockout
    - total_final_inventory: total inventory across zones
    """
    finished_goods_data = {
        'zones': {},
        'is_stockout': False,
        'total_final_inventory': 0,
        'raw_df': None
    }
    
    for zone in ZONES:
        is_active = zone in ZONES[:3]
        capacity = rng.integers(*BOUNDS["inventory_capacity"]) if is_active else 0
        inventory = rng.integers(500, capacity) if capacity > 0 else 0
        final = rng.integers(0, capacity) if capacity > 0 else 0
        
        # Randomly determine if there's a stockout (10% chance for active zones)
        if is_active and rng.random() < 0.1:
            final = 0
            finished_goods_data['is_stockout'] = True
        
        finished_goods_data['zones'][zone] = {
            'inventory': inventory,
            'capacity': capacity,
            'final': final,
        }
        finished_goods_data['total_final_inventory'] += final
    
    return finished_goods_data


def generate_production_data(rng):
    """Generate production_data structure for Production dashboard."""
    production_data = {
        'by_zone': {},
        'costs': {},
        'summary': {},
    }
    
    sections = ['Section 1', 'Section 2', 'Section 3', 'Section 4', 'Section 5']
    
    for zone in ZONES:
        is_active = zone in ZONES[:3]
        zone_sections = {}
        
        for section in sections:
            zone_sections[section] = {
                'produced': rng.integers(10000, 18000) if is_active else 0,
                'direct_cost': rng.integers(20000, 60000) if is_active else 0,
                'indirect_cost': rng.integers(5000, 15000) if is_active else 0,
                'holding_cost': rng.integers(10000, 25000) if is_active else 0,
            }
            zone_sections[section]['total_cost'] = (
                zone_sections[section]['direct_cost'] +
                zone_sections[section]['indirect_cost'] +
                zone_sections[section]['holding_cost']
            )
        
        production_data['by_zone'][zone] = zone_sections
    
    production_data['summary'] = {
        'total_production': sum(
            sum(s['produced'] for s in zdata.values())
            for zdata in production_data['by_zone'].values()
        ),
        'total_cost': sum(
            sum(s['total_cost'] for s in zdata.values())
            for zdata in production_data['by_zone'].values()
        ),
    }
    
    return production_data


def generate_esg_data(rng):
    """Generate esg_data structure for ESG dashboard.
    
    Must match structure expected by data_loader.load_esg_report:
    - emissions: total emissions value
    - energy: energy consumption value
    - energy_consumption: same as energy
    - tax_rate: tax rate percentage
    """
    emissions = rng.integers(5000, 25000)
    energy = rng.integers(500000, 2000000)
    
    esg_data = {
        'emissions': emissions,
        'energy': energy,
        'energy_consumption': energy,
        'tax_rate': rng.integers(20, 40),
        'raw_df': None
    }
    
    return esg_data


def generate_balance_data(rng):
    """Generate balance_data structure for CFO dashboard."""
    revenue = rng.integers(2000000, 8000000)
    cogs = int(revenue * rng.uniform(0.5, 0.7))
    gross_profit = revenue - cogs
    opex = int(revenue * rng.uniform(0.15, 0.25))
    operating_income = gross_profit - opex
    tax = int(max(0, operating_income * rng.uniform(0.2, 0.35)))
    net_income = operating_income - tax
    
    balance_data = {
        'income_statement': {
            'revenue': revenue,
            'cost_of_goods_sold': cogs,
            'gross_profit': gross_profit,
            'operating_expenses': opex,
            'operating_income': operating_income,
            'interest_expense': rng.integers(10000, 50000),
            'tax_expense': tax,
            'net_income': net_income,
        },
        'balance_sheet': {
            'assets': {
                'cash': rng.integers(*BOUNDS["cash_balance"]),
                'accounts_receivable': rng.integers(*BOUNDS["receivables"]),
                'inventory': rng.integers(300000, 800000),
                'ppe': rng.integers(500000, 2000000),
                'total_assets': 0,  # Calculated below
            },
            'liabilities': {
                'accounts_payable': rng.integers(*BOUNDS["payables"]),
                'short_term_debt': rng.integers(50000, 200000),
                'long_term_debt': rng.integers(*BOUNDS["mortgage"]),
                'total_liabilities': 0,  # Calculated below
            },
            'equity': {
                'common_stock': rng.integers(200000, 500000),
                'retained_earnings': rng.integers(100000, 500000),
                'total_equity': 0,  # Calculated below
            }
        },
        'ratios': {},
    }
    
    # Calculate totals
    assets = balance_data['balance_sheet']['assets']
    liabilities = balance_data['balance_sheet']['liabilities']
    equity = balance_data['balance_sheet']['equity']
    
    assets['total_assets'] = assets['cash'] + assets['accounts_receivable'] + assets['inventory'] + assets['ppe']
    liabilities['total_liabilities'] = liabilities['accounts_payable'] + liabilities['short_term_debt'] + liabilities['long_term_debt']
    equity['total_equity'] = equity['common_stock'] + equity['retained_earnings']
    
    # Ratios
    balance_data['ratios'] = {
        'current_ratio': (assets['cash'] + assets['accounts_receivable'] + assets['inventory']) / max(1, liabilities['accounts_payable'] + liabilities['short_term_debt']),
        'debt_to_equity': liabilities['total_liabilities'] / max(1, equity['total_equity']),
        'gross_margin': gross_profit / max(1, revenue),
        'net_margin': net_income / max(1, revenue),
        'roa': net_income / max(1, assets['total_assets']),
    }
    
    return balance_data


def generate_sales_admin_data(rng):
    """Generate sales_admin_data structure for CFO dashboard and CMO.
    
    Must match structure expected by data_loader.load_sales_admin_expenses:
    - by_zone: dict with zone names, each containing 'units' and 'price'
    - totals: dict with units, tv_spend, radio_spend, salespeople_cost
    """
    salespeople = rng.integers(*BOUNDS["salespeople"])
    salary_per = rng.integers(*BOUNDS["salary_per_person"])
    tv_spots = rng.integers(*BOUNDS["tv_spots"])
    tv_cost_per = rng.integers(*BOUNDS["tv_cost_per_spot"])
    radio_spots = rng.integers(*BOUNDS["radio_spots"])
    radio_cost_per = rng.integers(*BOUNDS["radio_cost_per_spot"])
    
    # Generate zone-level data (expected by CMO dashboard)
    by_zone = {}
    total_units = 0
    for zone in ZONES:
        is_active = zone in ZONES[:3]
        units = rng.integers(1000, 10000) if is_active else 0
        price = rng.integers(*BOUNDS["unit_price"]) if is_active else 0
        by_zone[zone] = {
            'units': units,
            'price': price,
        }
        total_units += units
    
    return {
        'by_zone': by_zone,
        'totals': {
            'units': total_units,
            'tv_spend': tv_spots * tv_cost_per,
            'radio_spend': radio_spots * radio_cost_per,
            'salespeople_cost': salespeople * salary_per,
        },
        'total_expenses': salespeople * salary_per + tv_spots * tv_cost_per + radio_spots * radio_cost_per,
        'categories': {},
        'raw_df': None
    }


def generate_subperiod_cash_data(rng):
    """Generate subperiod_cash_data structure for CFO dashboard."""
    cash_data = {
        'by_fortnight': [],
        'summary': {},
    }
    
    starting_cash = rng.integers(50000, 200000)
    current_cash = starting_cash
    
    for fn in range(FORTNIGHTS):
        inflows = rng.integers(100000, 400000)
        outflows = rng.integers(80000, 350000)
        net = inflows - outflows
        ending_cash = current_cash + net
        
        cash_data['by_fortnight'].append({
            'fortnight': fn + 1,
            'beginning_cash': current_cash,
            'inflows': inflows,
            'outflows': outflows,
            'net_cash_flow': net,
            'ending_cash': ending_cash,
        })
        
        current_cash = ending_cash
    
    cash_data['summary'] = {
        'starting_cash': starting_cash,
        'ending_cash': current_cash,
        'total_inflows': sum(f['inflows'] for f in cash_data['by_fortnight']),
        'total_outflows': sum(f['outflows'] for f in cash_data['by_fortnight']),
    }
    
    return cash_data


def generate_ar_ap_data(rng):
    """Generate ar_ap_data structure for CFO dashboard."""
    return {
        'accounts_receivable': {
            'current': rng.integers(100000, 300000),
            '30_days': rng.integers(50000, 150000),
            '60_days': rng.integers(20000, 80000),
            '90_plus_days': rng.integers(5000, 30000),
            'total': 0,  # Calculated
        },
        'accounts_payable': {
            'current': rng.integers(50000, 150000),
            '30_days': rng.integers(30000, 100000),
            '60_days': rng.integers(10000, 50000),
            '90_plus_days': rng.integers(2000, 20000),
            'total': 0,  # Calculated
        },
        'summary': {}
    }


def generate_financial_summary_data(rng):
    """Generate financial_summary_data structure for CFO dashboard."""
    revenue = rng.integers(2000000, 8000000)
    
    return {
        'revenue': revenue,
        'gross_profit': int(revenue * rng.uniform(0.3, 0.5)),
        'operating_income': int(revenue * rng.uniform(0.1, 0.25)),
        'net_income': int(revenue * rng.uniform(0.05, 0.15)),
        'ebitda': int(revenue * rng.uniform(0.15, 0.30)),
        'eps': rng.uniform(1.0, 5.0),
    }


def generate_initial_cash_data(rng):
    """Generate initial_cash_data structure for CFO dashboard."""
    cash_start = rng.integers(*BOUNDS["cash_balance"])
    
    return {
        'initial_cash': cash_start,
        'tax_payments': -rng.integers(10000, 50000),
        'supplier_payments': -rng.integers(5000, 30000),
        'hiring_expenses': -rng.integers(0, 5000),
        'warehouse_rental': -rng.integers(20000, 100000),
        'mortgage_interest': -rng.integers(10000, 40000),
        'financing_received': rng.integers(50000, 200000),
    }


def generate_logistics_data(rng):
    """Generate logistics_data structure for Logistics dashboard."""
    logistics_data = {
        'shipping': {},
        'warehouses': {},
        'costs': {},
    }
    
    for zone in ZONES:
        is_active = zone in ZONES[:3]
        logistics_data['shipping'][zone] = {
            'units_shipped': rng.integers(5000, 20000) if is_active else 0,
            'cost_per_unit': rng.uniform(*BOUNDS["transport_cost_per_unit"]) if is_active else 0,
            'total_cost': 0,  # Calculated
        }
        logistics_data['shipping'][zone]['total_cost'] = (
            logistics_data['shipping'][zone]['units_shipped'] *
            logistics_data['shipping'][zone]['cost_per_unit']
        )
        
        logistics_data['warehouses'][zone] = {
            'capacity': rng.integers(5000, 15000) if is_active else 0,
            'utilization': rng.uniform(0.5, 0.95) if is_active else 0,
            'monthly_cost': rng.integers(10000, 50000) if is_active else 0,
        }
    
    logistics_data['costs'] = {
        'total_shipping': sum(z['total_cost'] for z in logistics_data['shipping'].values()),
        'total_warehouse': sum(z['monthly_cost'] for z in logistics_data['warehouses'].values()),
    }
    
    return logistics_data


def generate_machine_spaces_data(rng):
    """Generate machine_spaces_data structure for Production dashboard."""
    machine_types = ['M1', 'M2', 'M3-alpha', 'M3-beta', 'M4']
    
    machine_data = {
        'by_zone': {},
        'totals': {},
    }
    
    for zone in ZONES:
        is_active = zone in ZONES[:3]
        zone_machines = {}
        
        for mtype in machine_types:
            count = rng.integers(0, 15) if is_active else 0
            cost_per = rng.integers(5000, 100000) if 'M3' in mtype else rng.integers(2000, 15000)
            
            zone_machines[mtype] = {
                'count': count,
                'acquisition_cost': cost_per,
                'total_value': count * cost_per,
                'capacity': count * rng.integers(500, 2000),
                'utilization': rng.uniform(0.6, 0.95) if count > 0 else 0,
            }
        
        machine_data['by_zone'][zone] = zone_machines
    
    machine_data['totals'] = {
        'total_machines': sum(
            sum(m['count'] for m in zdata.values())
            for zdata in machine_data['by_zone'].values()
        ),
        'total_capacity': sum(
            sum(m['capacity'] for m in zdata.values())
            for zdata in machine_data['by_zone'].values()
        ),
    }
    
    return machine_data


# =============================================================================
# MAIN GENERATOR FUNCTION
# =============================================================================

def generate_all_random_data(seed=None):
    """
    Generate all random data structures for the ExSim dashboards.
    
    Args:
        seed: Optional random seed for reproducibility. If None, uses random seed.
    
    Returns:
        dict: Dictionary mapping state_key -> generated data structure
    """
    if seed is None:
        rng = np.random.default_rng()
    else:
        rng = np.random.default_rng(seed)
    
    # Generate all data structures
    generated_data = {
        'market_data': generate_market_data(rng),
        'workers_data': generate_workers_data(rng),
        'materials_data': generate_materials_data(rng),
        'finished_goods_data': generate_finished_goods_data(rng),
        'production_data': generate_production_data(rng),
        'esg_data': generate_esg_data(rng),
        'balance_data': generate_balance_data(rng),
        'sales_admin_data': generate_sales_admin_data(rng),
        'subperiod_cash_data': generate_subperiod_cash_data(rng),
        'ar_ap_data': generate_ar_ap_data(rng),
        'financial_summary_data': generate_financial_summary_data(rng),
        'initial_cash_data': generate_initial_cash_data(rng),
        'logistics_data': generate_logistics_data(rng),
        'machine_spaces_data': generate_machine_spaces_data(rng),
    }
    
    return generated_data
