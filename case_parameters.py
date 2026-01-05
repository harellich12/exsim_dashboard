"""
ExSim Case Parameters - Mezquite Inc.
-------------------------------------
Single source of truth for all simulation constants derived from the
Mezquite Inc. / ElectroClean Case Study.

Usage:
    from case_parameters import FINANCIAL, MARKET, PRODUCTION, WORKFORCE, LOGISTICS
"""

# =============================================================================
# 0. COMMON PARAMETERS (Shared Across All Dashboards)
# =============================================================================
COMMON = {
    # Geographic zones - used by Production, Logistics, Marketing, CPO
    "ZONES": ["Center", "West", "North", "East", "South"],
    
    # Time periods - 8 fortnights per simulation period
    "FORTNIGHTS": list(range(1, 9)),  # [1, 2, 3, 4, 5, 6, 7, 8]
    
    # Market segments - used by CMO, Production
    "SEGMENTS": ["High", "Low"],
    
    # Production sections - used by Production Manager
    "SECTIONS": ["Section 1", "Section 2", "Section 3"],
    
    # Machine types - used by Production, CFO (depreciation)
    "MACHINE_TYPES": ["M1", "M2", "M3-alpha", "M3-beta", "M4"],
    
    # Transport modes - used by CLO
    "TRANSPORT_MODES": ["Train", "Truck", "Plane"],
    
    # Parts and pieces - used by Purchasing
    "PARTS": ["Part A", "Part B"],
    "PIECES": ["Piece 1", "Piece 2", "Piece 3", "Piece 4", "Piece 5", "Piece 6"],
    
    # Output file standard format
    "MY_COMPANY": "Company 3"  # Default company identifier
}

# =============================================================================
# 1. FINANCIAL PARAMETERS (CFO)
# =============================================================================
FINANCIAL = {
    "TAX_RATE": 0.25,           # 25% Income Tax
    "CASH_INTEREST_RATE": 0.02, # 2% Annual Interest on Cash
    
    "LOANS": {
        "SHORT_TERM": {
            "INTEREST_RATE_ANNUAL": 0.08,  # 8%
            "LIMIT": 500000
        },
        "LONG_TERM": {
            "INTEREST_RATE_ANNUAL": 0.06,  # 6%
            "LIMIT": 2000000
        }
    },
    
    # Payment Terms Impact (Discount %)
    "PAYMENT_TERMS": {
        "A": 0.13,  # 0 fortnights wait -> 13% discount cost
        "B": 0.09,  # 2 fortnights wait
        "C": 0.05,  # 4 fortnights wait
        "D": 0.00   # 8 fortnights wait -> 0% discount
    }
}

# =============================================================================
# 2. MARKET PARAMETERS (CMO)
# =============================================================================
MARKET = {
    "ZONES": ["Center", "West", "North", "East", "South"],
    "SEGMENTS": ["High", "Low"],
    
    # Total Addressable Market (Households/Population)
    "POPULATION": {
        "Center": {"High": 35200, "Low": 74800},
        "West":   {"High": 24000, "Low": 51000},
        "North":  {"High": 19600, "Low": 50400},
        "East":   {"High": 10000, "Low": 40000},
        "South":  {"High": 9600,  "Low": 20400}
    },
    
    "SEASONALITY_PEAKS": [3, 6], # Periods with high demand (Holiday seasons)
    
    # Innovation Features Cost Structure (Table II.1)
    "INNOVATION_COSTS": {
        "STAINLESS MATERIAL":         {"upfront": 15000, "variable": 0.15},
        "RECYCLABLE MATERIALS":       {"upfront": 30000, "variable": 0.15},
        "ENERGY EFFICIENCY":          {"upfront": 30000, "variable": 0.15},
        "LIGHTER AND MORE COMPACT":  {"upfront": 35000, "variable": 0.30},
        "IMPACT RESISTANCE":          {"upfront": 40000, "variable": 0.30},
        "NOISE REDUCTION":            {"upfront": 40000, "variable": 0.30},
        "IMPROVED BATTERY CAPACITY":  {"upfront": 45000, "variable": 0.45},
        "SELF-CLEANING":              {"upfront": 60000, "variable": 0.60},
        "SPEED SETTINGS":             {"upfront": 60000, "variable": 0.60},
        "DIGITAL CONTROLS":           {"upfront": 75000, "variable": 0.75},
        "VOICE ASSISTANCE INTEGRATION":{"upfront": 75000, "variable": 0.75},
        "AUTOMATION AND PROGRAMMABILITY":{"upfront": 90000, "variable": 0.90},
        "MULTIFUNCTIONAL ACCESSORIES": {"upfront": 90000, "variable": 0.90},
        "MAPPING TECHNOLOGY":         {"upfront": 100000, "variable": 1.00}
    }
}

# =============================================================================
# 3. PRODUCTION PARAMETERS (Production Manager)
# =============================================================================
PRODUCTION = {
    "MACHINERY": {
        "M1": {
            "cost": 10600,
            "capacity_units": 200,
            "workers_required": 10, # Implied/Estimated (verify?)
            "type": "Manual"
        },
        "M2": { # M2 implied exists but case focuses on M3? Adding generic if needed.
             "cost": 40000, # Placeholder
             "capacity_units": 300
        },
        "M3_ALPHA": {
            "cost": 88500,
            "capacity_units": 450,
            "workers_required": 30,
            "type": "Semi-Auto"
        },
        "M3_BETA": {
            "cost": 155400,
            "capacity_units": 600,
            "workers_required": 6,
            "type": "Automated"
        }
    },
    
    "FACILITIES": {
        "MODULE_COST": 25000,
        "MODULE_CAPACITY_MACHINES": 4 # Example: 4 machines per module?
    }
}

# =============================================================================
# 4. WORKFORCE PARAMETERS (CPO)
# =============================================================================
WORKFORCE = {
    "SALESFORCE": {
        "BASE_SALARY": 750,
        "HIRING_COST": 1000,
        "FIRING_COST": 2000 # Estimate (usually 2-3x salary or specific policy)
    },
    
    "PRODUCTION_WORKERS": {
        "BASE_SALARY": 650,
        "HIRING_COST": 1250,
        "TRAINING_COST": 350,
        "OVERTIME_MULTIPLIER": 1.4,
        "OVERTIME_CAPACITY_PCT": 0.20 # Max 20% extra capacity via overtime
    }
}

# =============================================================================
# 5. LOGISTICS PARAMETERS (CLO)
# =============================================================================
# Costs per unit (Placeholder structure - specific matrix needed from case p.26)
LOGISTICS = {
    # Costs from Factory (Center?) to Zones
    "TRANSPORT_COSTS": {
        "Center": {"Truck": 0.5, "Train": 0.3, "Plane": 2.0},
        "West":   {"Truck": 1.2, "Train": 0.8, "Plane": 3.0},
        "North":  {"Truck": 1.0, "Train": 0.7, "Plane": 2.5},
        "East":   {"Truck": 1.5, "Train": 1.0, "Plane": 3.5},
        "South":  {"Truck": 1.8, "Train": 1.2, "Plane": 4.0}
    }
}

# =============================================================================
# 6. PURCHASING PARAMETERS (Purchasing Manager)
# =============================================================================
PURCHASING = {
    "BOM_COSTS": {
        "Part A": 15.00,  # Estimated Base Cost
        "Part B": 25.00,  # Estimated Base Cost
        "Piece 1": 2.00,
        "Piece 2": 2.00,
        "Piece 3": 2.00,
        "Piece 4": 2.00,
        "Piece 5": 2.00,
        "Piece 6": 2.00
    }
}

# =============================================================================
# 7. ESG PARAMETERS (Strategy)
# =============================================================================
ESG = {
    "STRATEGY": {
        "CARBON_NEUTRAL_YEAR": 2030,
        "RECYCLABILITY_TARGET": 1.0, # 100%
        "FAIR_WAGE_MULTIPLIER": 1.1, # Target 10% above market
        "COMMUNITY_INVESTMENT_PCT": 0.01 # 1% of profits
    }
}
