"""
ExSim Case Parameters - Mezquite Inc.
-------------------------------------
Single source of truth for all simulation constants derived from the
Mezquite Inc. / ElectroClean Case Study (EXSIM Case.pdf).

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
    "SECTIONS": ["Section 1", "Section 2", "Section 3", "Section 4", "Section 5"],
    
    # Machine types - used by Production, CFO (depreciation)
    "MACHINE_TYPES": ["M1", "M2", "M3-alpha", "M3-beta", "M4"],
    
    # Transport modes - used by CLO
    "TRANSPORT_MODES": ["Train", "Truck", "Airplane"],
    
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
    
    # Payment Terms Impact (Table III.3)
    # Payment type -> (fortnights delay, discount %)
    "PAYMENT_TERMS": {
        "A": {"fortnights": 0, "discount": 0.130},  # Immediate payment -> 13% discount
        "B": {"fortnights": 2, "discount": 0.075},  # 2 fortnights -> 7.5% discount
        "C": {"fortnights": 4, "discount": 0.025},  # 4 fortnights -> 2.5% discount
        "D": {"fortnights": 8, "discount": 0.000}   # 8 fortnights -> 0% discount
    }
}

# =============================================================================
# 2. MARKET PARAMETERS (CMO)
# =============================================================================
MARKET = {
    "ZONES": ["Center", "West", "North", "East", "South"],
    "SEGMENTS": ["High", "Low"],
    
    # Total Addressable Market (Table I.1 - Households/Population per company)
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
        # ID 1
        "STAINLESS MATERIAL":         {"upfront": 15000, "variable": 0.15},
        # ID 2
        "RECYCLABLE MATERIALS":       {"upfront": 30000, "variable": 0.15},
        # ID 3
        "ENERGY EFFICIENCY":          {"upfront": 30000, "variable": 0.15},
        # ID 4
        "LIGHTER AND MORE COMPACT":   {"upfront": 30000, "variable": 0.30},
        # ID 5
        "IMPACT-RESISTANCE":          {"upfront": 30000, "variable": 0.30},
        # ID 6
        "NOISE REDUCTION":            {"upfront": 45000, "variable": 0.30},
        # ID 7
        "IMPROVED BATTERY CAPACITY":  {"upfront": 45000, "variable": 0.45},
        # ID 8
        "SELF-CLEANING":              {"upfront": 45000, "variable": 0.45},
        # ID 9
        "SPEED SETTINGS":             {"upfront": 45000, "variable": 0.45},
        # ID 10
        "DIGITAL CONTROLS":           {"upfront": 45000, "variable": 0.45},
        # ID 11
        "VOICE ASSISTANCE INTEGRATION": {"upfront": 45000, "variable": 0.75},
        # ID 12
        "AUTOMATION AND PROGRAMMABILITY": {"upfront": 45000, "variable": 0.75},
        # ID 13
        "MULTIFUNCTIONAL ACCESSORIES": {"upfront": 100000, "variable": 1.00},
        # ID 14
        "MAPPING TECHNOLOGY":         {"upfront": 100000, "variable": 1.00}
    },
    
    # CO2 Emissions from Improvements (Table VII.1)
    "INNOVATION_CO2": {
        "STAINLESS MATERIAL":         -0.025,
        "RECYCLABLE MATERIALS":        0.023,
        "ENERGY EFFICIENCY":          -0.025,
        "LIGHTER AND MORE COMPACT":    0.045,
        "IMPACT-RESISTANCE":          -0.050,
        "NOISE REDUCTION":             0.045,
        "IMPROVED BATTERY CAPACITY":  -0.075,
        "SELF-CLEANING":               0.068,
        "SPEED SETTINGS":              0.068,
        "DIGITAL CONTROLS":           -0.075,
        "VOICE ASSISTANCE INTEGRATION": 0.107,
        "AUTOMATION AND PROGRAMMABILITY": 0.107,
        "MULTIFUNCTIONAL ACCESSORIES": 0.143,
        "MAPPING TECHNOLOGY":          0.143
    }
}

# =============================================================================
# 3. PRODUCTION PARAMETERS (Production Manager)
# =============================================================================
PRODUCTION = {
    # Table IV.1 - Machine characteristics
    "MACHINERY": {
        "M1": {
            "price": 10600,
            "capacity_per_fortnight": 200,  # ASM-A
            "workers_required": 1,
            "spaces": 1,
            "lifespan_periods": 10,
            "power_kw": 10
        },
        "M2": {
            "price": 7000,
            "capacity_per_fortnight": 70,  # ASM-B
            "workers_required": 1,
            "spaces": 1,
            "lifespan_periods": 10,
            "power_kw": 9
        },
        "M3_ALPHA": {
            "price": 88500,
            "capacity_per_fortnight": 450,  # ASM-C
            "workers_required": 30,
            "spaces": 5,
            "lifespan_periods": 20,
            "power_kw": 15
        },
        "M3_BETA": {
            "price": 155400,
            "capacity_per_fortnight": 600,  # ASM-C
            "workers_required": 6,
            "spaces": 6,
            "lifespan_periods": 20,
            "power_kw": 16
        },
        "M4": {
            "price": 2500,
            "capacity_section_3": 400,  # ASM-C
            "capacity_section_4": 150,  # ASM-D
            "capacity_section_5": 130,  # Electrocleans
            "workers_required": 1,
            "spaces": 1,
            "lifespan_periods": 8,
            "power_kw": 6
        }
    },
    
    # Crew capacity (Section 5 manual assembly)
    "CREW": {
        "workers_per_crew": 3,
        "capacity_per_crew": 50  # Electrocleans per fortnight
    },
    
    # Machine resale value
    "RESALE_RATE": 0.65,  # 65% of net book value
    
    # Table IV.7 - Facilities ("Plant Modules")
    "FACILITIES": {
        "MODULE_PURCHASE_PRICE": 50000,
        "MODULE_RESALE_PRICE": 50000,
        "MODULE_LEASING_COST": 75000,  # Emergency leasing
        "MODULE_RENT_COST_PER_PERIOD": 7500,  # Regular rent per module per period
        "ADMIN_COST_PER_MODULE_PER_PERIOD": 10000,
        "SPACES_PER_MODULE": 18,
        # Current modules at Period 7
        "INITIAL_MODULES": {
            "Center": 4,
            "West": 2,
            "North": 0,
            "East": 0,
            "South": 0
        }
    },
    
    # Electricity costs (Table IV.3)
    "ELECTRICITY": {
        "POWER_COST_PER_KW_PER_PERIOD": 10,  # $10 per installed kW per period
        "CONSUMPTION_COST_PER_KWH": 0.06,    # $0.06 per kWh
        "CO2_PER_KWH": 0.4                    # 0.4 kg CO2 per kWh
    },
    
    # Table IV.2 - Machine Transfer Costs ($ per machine by airplane)
    "MACHINE_TRANSFER_COSTS": {
        "Center-West":  {"M1": 960, "M2": 840, "M3_ALPHA": 1440, "M3_BETA": 1800, "M4": 360},
        "Center-North": {"M1": 1320, "M2": 1155, "M3_ALPHA": 1980, "M3_BETA": 2475, "M4": 495},
        "Center-East":  {"M1": 2460, "M2": 2152.50, "M3_ALPHA": 3690, "M3_BETA": 4612.50, "M4": 922.50},
        "Center-South": {"M1": 1200, "M2": 1050, "M3_ALPHA": 1800, "M3_BETA": 2250, "M4": 450},
        "West-North":   {"M1": 1260, "M2": 1102.50, "M3_ALPHA": 1890, "M3_BETA": 2362.50, "M4": 472.50},
        "West-East":    {"M1": 1440, "M2": 1260, "M3_ALPHA": 2160, "M3_BETA": 2700, "M4": 540},
        "West-South":   {"M1": 1740, "M2": 1522.50, "M3_ALPHA": 2610, "M3_BETA": 3262.50, "M4": 652.50},
        "North-East":   {"M1": 1740, "M2": 1522.50, "M3_ALPHA": 2610, "M3_BETA": 3262.50, "M4": 652.50},
        "North-South":  {"M1": 2160, "M2": 1890, "M3_ALPHA": 3240, "M3_BETA": 4050, "M4": 810},
        "East-South":   {"M1": 3480, "M2": 3045, "M3_ALPHA": 5220, "M3_BETA": 6525, "M4": 1305}
    },
    
    # Table IV.1 - Initial Machine Counts per Region/Section (Period 7)
    "INITIAL_MACHINES": {
        "Center": {
            "Section 1": {"M1": 7},
            "Section 2": {"M2": 22},
            "Section 3": {"M3_ALPHA": 3, "M3_BETA": 0, "M4": 4},
            "Section 4": {"M4": 10},
            "Section 5": {"M4": 11}
        },
        "West": {
            "Section 1": {"M1": 3},
            "Section 2": {"M2": 9},
            "Section 3": {"M3_ALPHA": 1, "M3_BETA": 0, "M4": 2},
            "Section 4": {"M4": 3},
            "Section 5": {"M4": 3}
        }
    },
    
    # Table IV.5 - Initial Raw Materials Inventory (Period 7)
    "INITIAL_INVENTORY": {
        "Center": {
            "Part A": 3496, "Part B": 765,
            "Piece 1": 94, "Piece 2": 3844, "Piece 3": 3125,
            "Piece 4": 6131, "Piece 5": 16736, "Piece 6": 4439,
            "Assembly A": 7432, "Assembly B": 5392, "Assembly C": 3766, "Assembly D": 6606
        },
        "West": {
            "Part A": 1016, "Part B": 293,
            "Piece 1": 36, "Piece 2": 1424, "Piece 3": 1208,
            "Piece 4": 2356, "Piece 5": 6456, "Piece 6": 1596,
            "Assembly A": 3622, "Assembly B": 2210, "Assembly C": 2416, "Assembly D": 2266
        }
    }
}

# =============================================================================
# 4. WORKFORCE PARAMETERS (CPO)
# =============================================================================
WORKFORCE = {
    # Table III.2 - Distributor Sales Force
    "SALESFORCE": {
        "SALARY_PER_FORTNIGHT": 750,  # $750/fortnight -> $6000/period
        "HIRING_COST": 1000,          # One-time setup/training
        "LAYOFF_COST": 0,             # Per PDF: "Layoffs do not incur additional expenses"
        "INITIAL_COUNT": 44           # Period 6 force
    },
    
    # Table VI.1 - Factory Workers
    "PRODUCTION_WORKERS": {
        "SALARY_PER_FORTNIGHT": 27.3,  # Current salary
        "MINIMUM_SALARY": 26.0,         # Government minimum
        "HIRING_COST": 240,
        "LAYOFF_COST": 220,
        "OVERTIME_BONUS_PER_FORTNIGHT": 12,  # Flat $12/fortnight for Saturday work
        "OVERTIME_CAPACITY_INCREASE": 0.20,   # 20% extra production
        "INDIRECT_EXPENSE_RATE": 0.50         # 50% of salaries for foremen, QC, etc.
    },
    
    # Initial worker allocation (Table VI.1)
    "INITIAL_WORKERS": {
        "Center": {
            "M1": 7, "M2": 22, "M3_alpha": 90, "M4_S3": 4, "M4_S4": 10, "M4_S5": 11, "Crews": 75,
            "Total": 219
        },
        "West": {
            "M1": 3, "M2": 9, "M3_alpha": 30, "M4_S3": 2, "M4_S4": 3, "M4_S5": 3, "Crews": 21,
            "Total": 71
        }
    }
}

# =============================================================================
# 5. LOGISTICS PARAMETERS (CLO)
# =============================================================================
LOGISTICS = {
    # Table V.1 - Warehousing
    "WAREHOUSE": {
        "RENTAL_COST_PER_MODULE_PER_PERIOD": 800,
        "CAPACITY_PER_MODULE": 100,  # Electroclean units
        # Initial capacity at Period 6
        "INITIAL_MODULES": {
            "Center": 48,
            "West": 25,
            "North": 20,
            "East": 0,
            "South": 0
        }
    },
    
    # Table V.3 - Transport Costs ($ per Electroclean)
    # Format: {route: {mode: {size: cost}}}
    # Sizes: Small (1-999), Medium (1000-1999), Large (2000+)
    "TRANSPORT_COSTS": {
        "Center-West": {
            "Airplane": 19.20,
            "Truck": {"Small": 18.00, "Medium": 16.80, "Large": 15.60},
            "Train": {"Small": 13.20, "Medium": 12.00, "Large": 10.80}
        },
        "Center-North": {
            "Airplane": 26.40,
            "Truck": {"Small": 20.16, "Medium": 18.96, "Large": 18.48},
            "Train": {"Small": 19.69, "Medium": 18.24, "Large": 14.40}
        },
        "Center-East": {
            "Airplane": 49.20,
            "Truck": {"Small": 47.52, "Medium": 43.52, "Large": 40.80},
            "Train": {"Small": 46.32, "Medium": 42.00, "Large": 39.60}
        },
        "Center-South": {
            "Airplane": 24.00,
            "Truck": {"Small": 18.00, "Medium": 16.80, "Large": 16.32},
            "Train": {"Small": 16.80, "Medium": 15.60, "Large": 12.72}
        },
        "West-North": {
            "Airplane": 25.20,
            "Truck": {"Small": 22.08, "Medium": 21.60, "Large": 18.00},
            "Train": {"Small": 20.40, "Medium": 18.00, "Large": 16.80}
        },
        "West-East": {
            "Airplane": 28.80,
            "Truck": {"Small": 28.32, "Medium": 27.60, "Large": 26.64},
            "Train": {"Small": 27.60, "Medium": 26.40, "Large": 24.48}
        },
        "West-South": {
            "Airplane": 34.80,
            "Truck": {"Small": 33.60, "Medium": 32.40, "Large": 31.20},
            "Train": {"Small": 32.40, "Medium": 28.80, "Large": 24.00}
        },
        "North-East": {
            "Airplane": 34.80,
            "Truck": {"Small": 33.60, "Medium": 31.20, "Large": 28.80},
            "Train": {"Small": 32.40, "Medium": 28.80, "Large": 24.00}
        },
        "North-South": {
            "Airplane": 43.20,
            "Truck": {"Small": 38.40, "Medium": 37.20, "Large": 36.00},
            "Train": {"Small": 38.40, "Medium": 36.00, "Large": 28.80}
        },
        "East-South": {
            "Airplane": 69.60,
            "Truck": {"Small": 64.80, "Medium": 58.80, "Large": 57.60},
            "Train": {"Small": 63.60, "Medium": 57.60, "Large": 56.40}
        }
    },
    
    # Table V.4 - Transit Times & Reliability
    # Airplane: Always 1 fortnight, 100% reliable
    # Truck/Train: Probability distribution of delivery fortnights
    "TRANSIT_TIMES": {
        "Airplane": {"fortnights": 1, "reliability": 1.00},  # Always delivers in 1 FN
        "Truck": {
            "Center-West": {2: 0.70, 3: 0.30},
            "Center-North": {2: 0.50, 3: 0.50},
            "Center-East": {4: 0.40, 5: 0.40, 6: 0.20},
            "Center-South": {2: 0.60, 3: 0.40},
            "West-North": {2: 0.55, 3: 0.45},
            "West-East": {3: 0.40, 4: 0.40, 5: 0.20},
            "West-South": {3: 0.50, 4: 0.50},
            "North-East": {3: 0.50, 4: 0.50},
            "North-South": {3: 0.40, 4: 0.40, 5: 0.20},
            "East-South": {5: 0.50, 6: 0.50}
        },
        "Train": {
            "Center-West": {3: 0.60, 4: 0.40},
            "Center-North": {4: 0.50, 5: 0.50},
            "Center-East": {5: 0.30, 6: 0.40, 7: 0.30},
            "Center-South": {3: 0.50, 4: 0.50},
            "West-North": {4: 0.50, 5: 0.50},
            "West-East": {4: 0.40, 5: 0.40, 6: 0.20},
            "West-South": {4: 0.40, 5: 0.40, 6: 0.20},
            "North-East": {4: 0.40, 5: 0.40, 6: 0.20},
            "North-South": {5: 0.40, 6: 0.40, 7: 0.20},
            "East-South": {6: 0.40, 7: 0.40, 8: 0.20}
        }
    },
    
    # CO2 Emissions from Transport (Table VII.1) - kg CO2 per Electroclean
    "TRANSPORT_CO2": {
        "Center-West":  {"Airplane": 16.00, "Truck": {"Small": 5.00, "Medium": 4.67, "Large": 4.33}, "Train": {"Small": 1.30, "Medium": 1.18, "Large": 1.06}},
        "Center-North": {"Airplane": 22.00, "Truck": {"Small": 5.60, "Medium": 5.27, "Large": 5.13}, "Train": {"Small": 1.94, "Medium": 1.80, "Large": 1.42}},
        "Center-East":  {"Airplane": 41.00, "Truck": {"Small": 13.20, "Medium": 12.09, "Large": 11.33}, "Train": {"Small": 4.56, "Medium": 4.14, "Large": 3.90}},
        "Center-South": {"Airplane": 20.00, "Truck": {"Small": 5.00, "Medium": 4.67, "Large": 4.53}, "Train": {"Small": 1.65, "Medium": 1.54, "Large": 1.25}},
        "West-North":   {"Airplane": 21.00, "Truck": {"Small": 6.13, "Medium": 6.00, "Large": 5.00}, "Train": {"Small": 2.01, "Medium": 1.77, "Large": 1.65}},
        "West-East":    {"Airplane": 24.00, "Truck": {"Small": 7.87, "Medium": 7.67, "Large": 7.40}, "Train": {"Small": 2.72, "Medium": 2.60, "Large": 2.41}},
        "West-South":   {"Airplane": 29.00, "Truck": {"Small": 9.33, "Medium": 9.00, "Large": 8.67}, "Train": {"Small": 3.19, "Medium": 2.84, "Large": 2.36}},
        "North-East":   {"Airplane": 29.00, "Truck": {"Small": 9.33, "Medium": 8.67, "Large": 8.00}, "Train": {"Small": 3.19, "Medium": 2.84, "Large": 2.36}},
        "North-South":  {"Airplane": 36.00, "Truck": {"Small": 10.67, "Medium": 10.33, "Large": 10.00}, "Train": {"Small": 3.78, "Medium": 3.55, "Large": 2.84}},
        "East-South":   {"Airplane": 58.00, "Truck": {"Small": 18.00, "Medium": 16.33, "Large": 16.00}, "Train": {"Small": 6.26, "Medium": 5.67, "Large": 5.55}}
    }
}

# =============================================================================
# 6. PURCHASING PARAMETERS (Purchasing Manager)
# =============================================================================
PURCHASING = {
    # Table IV.4 - Raw Materials
    "PARTS": {
        "Part A": {
            "batch_size": 30,
            "suppliers": {
                "A": {"price": 125, "payment_fortnights": 0, "delivery_rate": 1.00},
                "B": {"price": 100, "discount_threshold": 150, "discount": 0.16, "payment_fortnights": 2, "delivery_rate": 0.80},
                "C": {"price": 140, "payment_fortnights": 8, "delivery_rate": 1.00}
            },
            "ordering_cost": 2300,
            "holding_cost_per_unit": 0.30,
            "consumption_per_electroclean": 1.00
        },
        "Part B": {
            "batch_size": 12,
            "suppliers": {
                "A": {"price": 330, "payment_fortnights": 0, "delivery_rate": 1.00},
                "B": {"price": 264, "discount_threshold": 80, "discount": 0.16, "payment_fortnights": 2, "delivery_rate": 0.80},
                "C": {"price": 370, "payment_fortnights": 8, "delivery_rate": 1.00}
            },
            "ordering_cost": 6000,
            "holding_cost_per_unit": 4.50,
            "consumption_per_electroclean": 0.20
        }
    },
    "PIECES": {
        "Piece 1": {"batch_size": 1, "price": 60, "consumption": 0.01},
        "Piece 2": {"batch_size": 100, "price": 7, "consumption": 8.00},
        "Piece 3": {"batch_size": 30, "price": 36, "consumption": 0.30},
        "Piece 4": {"batch_size": 60, "price": 24, "consumption": 0.60},
        "Piece 5": {"batch_size": 100, "price": 30, "consumption": 2.00},
        "Piece 6": {"batch_size": 150, "price": 28, "consumption": 0.75}
    },
    # Assembly holding costs
    "ASSEMBLIES": {
        "Assembly A": {"holding_cost": 0.20},
        "Assembly B": {"holding_cost": 0.40},
        "Assembly C": {"holding_cost": 0.60},
        "Assembly D": {"holding_cost": 0.20}
    },
    
    # Table IV.5 - Initial Raw Materials Inventory (Period 7)
    "INITIAL_INVENTORY": {
        "Center": {
            "Part A": 3496, "Part B": 765,
            "Piece 1": 94, "Piece 2": 3844, "Piece 3": 3125,
            "Piece 4": 6131, "Piece 5": 16736, "Piece 6": 14687,
            "Assembly A": 1045, "Assembly B": 789, "Assembly C": 1207, "Assembly D": 702
        },
        "West": {
            "Part A": 1016, "Part B": 293,
            "Piece 1": 36, "Piece 2": 1424, "Piece 3": 1208,
            "Piece 4": 2356, "Piece 5": 6456, "Piece 6": 5699,
            "Assembly A": 408, "Assembly B": 304, "Assembly C": 466, "Assembly D": 269
        }
    }
}

# =============================================================================
# 7. ESG PARAMETERS (Strategy)
# =============================================================================
ESG = {
    # Table VII.1 - CO2 from raw materials procurement (kg CO2/part)
    "RAW_MATERIALS_CO2": {
        "Part A": {"Supplier A": 3.67, "Supplier B": 6.18, "Supplier C": 3.58},
        "Part B": {"Supplier A": 5.46, "Supplier B": 8.33, "Supplier C": 5.33}
    },
    
    # Table VII.1 - Machine CO2 emissions (kg CO2/unit produced)
    "MACHINE_CO2": {
        "M1": {"emissions_at_capacity_kg": 320, "capacity": 200, "kg_per_unit": 1.60},
        "M2": {"emissions_at_capacity_kg": 288, "capacity": 70, "kg_per_unit": 4.11},
        "M3_alpha": {"emissions_at_capacity_kg": 480, "capacity": 450, "kg_per_unit": 1.07},
        "M3_beta": {"emissions_at_capacity_kg": 512, "capacity": 600, "kg_per_unit": 0.85},
        "M4_S3": {"emissions_at_capacity_kg": 192, "capacity": 400, "kg_per_unit": 0.48},
        "M4_S4": {"emissions_at_capacity_kg": 192, "capacity": 150, "kg_per_unit": 1.28},
        "M4_S5": {"emissions_at_capacity_kg": 192, "capacity": 130, "kg_per_unit": 1.48}
    },
    "ELECTRICITY_CO2_FACTOR": 0.4,  # kg CO2 per kWh
    
    # Table VII.1 - Other emissions
    "FACTORY_MODULE_CO2": 405000,  # kg CO2 per new module (distributed over 12 periods)
    "ELECTROCLEAN_DISPOSAL_CO2": 13.2,  # 12 kg recycling + 1.2 kg transport
    
    # Table VII.1 - Transport CO2 (kg CO2/electroclean by route and mode)
    "TRANSPORT_CO2": {
        "Central-Western": {"Airplane": {"Small": 16.00, "Medium": 5.00, "Large": 4.67}, 
                           "Truck": {"Small": 4.33, "Medium": 4.67, "Large": 4.33},
                           "Train": {"Small": 1.30, "Medium": 1.18, "Large": 1.06}},
        "Central-Northern": {"Airplane": {"Small": 22.00, "Medium": 5.60, "Large": 5.27},
                            "Truck": {"Small": 5.13, "Medium": 5.27, "Large": 5.13},
                            "Train": {"Small": 1.94, "Medium": 1.80, "Large": 1.42}},
        "Central-Eastern": {"Airplane": {"Small": 41.00, "Medium": 13.20, "Large": 12.09},
                           "Truck": {"Small": 11.33, "Medium": 12.09, "Large": 11.33},
                           "Train": {"Small": 4.56, "Medium": 4.14, "Large": 3.90}},
        "Central-Southern": {"Airplane": {"Small": 20.00, "Medium": 5.00, "Large": 4.67},
                            "Truck": {"Small": 4.53, "Medium": 4.67, "Large": 4.53},
                            "Train": {"Small": 1.65, "Medium": 1.54, "Large": 1.25}},
        "Western-Northern": {"Airplane": {"Small": 21.00, "Medium": 6.13, "Large": 6.00},
                            "Truck": {"Small": 5.00, "Medium": 6.00, "Large": 5.00},
                            "Train": {"Small": 2.01, "Medium": 1.77, "Large": 1.65}},
        "Western-Eastern": {"Airplane": {"Small": 24.00, "Medium": 7.87, "Large": 7.67},
                           "Truck": {"Small": 7.40, "Medium": 7.67, "Large": 7.40},
                           "Train": {"Small": 2.72, "Medium": 2.60, "Large": 2.41}},
        "Western-Southern": {"Airplane": {"Small": 29.00, "Medium": 9.33, "Large": 9.00},
                            "Truck": {"Small": 8.67, "Medium": 9.00, "Large": 8.67},
                            "Train": {"Small": 3.19, "Medium": 2.84, "Large": 2.36}},
        "Northern-Eastern": {"Airplane": {"Small": 29.00, "Medium": 9.33, "Large": 8.67},
                            "Truck": {"Small": 8.00, "Medium": 8.67, "Large": 8.00},
                            "Train": {"Small": 3.19, "Medium": 2.84, "Large": 2.36}},
        "Northern-Southern": {"Airplane": {"Small": 36.00, "Medium": 10.67, "Large": 10.33},
                             "Truck": {"Small": 10.00, "Medium": 10.33, "Large": 10.00},
                             "Train": {"Small": 3.78, "Medium": 3.55, "Large": 2.84}},
        "Eastern-Southern": {"Airplane": {"Small": 58.00, "Medium": 18.00, "Large": 16.33},
                            "Truck": {"Small": 16.00, "Medium": 16.33, "Large": 16.00},
                            "Train": {"Small": 6.26, "Medium": 5.67, "Large": 5.55}}
    },
    
    # Machine transport correction factors (% of electroclean emissions)
    "MACHINE_TRANSPORT_FACTORS": {
        "M1": 50.0, "M2": 43.75, "M3_alpha": 75.0, "M3_beta": 93.75, "M4": 18.75
    },
    
    # Table VII.1 - Product improvement emissions (kg CO2/unit)
    "IMPROVEMENT_CO2": {
        1: {"name": "Stainless Material", "kg_co2": -0.025},
        2: {"name": "Recyclable Materials", "kg_co2": 0.023},
        3: {"name": "Energy Efficiency", "kg_co2": -0.025},
        4: {"name": "Lighter and More Compact", "kg_co2": 0.045},
        5: {"name": "Impact-Resistance", "kg_co2": -0.05},
        6: {"name": "Noise Reduction", "kg_co2": 0.045},
        7: {"name": "Improved Battery Capacity", "kg_co2": -0.075},
        8: {"name": "Self-Cleaning", "kg_co2": 0.068},
        9: {"name": "Speed Settings", "kg_co2": 0.068},
        10: {"name": "Digital Controls", "kg_co2": -0.075},
        11: {"name": "Voice Assistance Integration", "kg_co2": 0.107},
        12: {"name": "Automation and Programmability", "kg_co2": 0.107},
        13: {"name": "Multifunctional Accessories", "kg_co2": 0.143},
        14: {"name": "Mapping Technology", "kg_co2": 0.143}
    },
    
    # Table VII.2 - CO2 Abatement Actions
    "ABATEMENT": {
        "SOLAR_PANELS": {
            "cost": 420,
            "lifespan_years": 25,
            "maintenance_per_period": 7,
            "energy_per_period_kwh": 266,
            "co2_reduction_per_period_kg": 106.4
        },
        "GREEN_ENERGY": {
            "regular_price_per_kwh": 0.06,
            "premium_rate": 0.20,  # 20% over regular price
            "co2_per_kwh_reduction": 0.4
        },
        "TREES": {
            "cost_per_tree": 6.25,
            "maintenance_per_period_per_80_trees": 16.67,
            "co2_absorbed_per_period_per_80_trees_kg": 333
        },
        "CO2_CREDITS": {
            "co2_per_credit_kg": 1000  # 1 ton = 1000 kg
        }
    },
    
    # Board targets
    "TARGETS": {
        "ANNUAL_CO2_REDUCTION": 0.15,  # 15% year-over-year reduction required
        "PERIOD_6_INTENSITY": 29.93  # kg CO2/unit baseline (example from case)
    }
}

# =============================================================================
# 8. FINANCE PARAMETERS (Chapter VIII)
# =============================================================================
FINANCE = {
    # Table VIII.1 - Financing Options (Period 7)
    "LINE_OF_CREDIT": {
        "interest_rate_per_period": 0.10,  # 10% per period
        "limit_pct_net_assets": 0.33,  # 33% of net fixed assets
        "current_balance": 113000,
        "net_fixed_assets_p6": 697625,  # For calculating limit
        "calculated_limit": 230216  # $697,625 * 0.33
    },
    
    "SHORT_TERM_DEPOSITS": {
        "interest_rate_per_period": 0.04,  # 4% per period
        "limit": None,  # No limit
        "current_balance": 200000
    },
    
    "MORTGAGES": {
        "interest_rate_per_period": 0.06,  # 6% per period
        "limit": 800000,
        "current_balance": 500000,
        "payment_schedule": [
            {"period": 10, "amount": 240000},
            {"period": 12, "amount": 130000},
            {"period": 18, "amount": 130000}
        ]
    },
    
    "EMERGENCY_LOAN": {
        "interest_rate_per_period": 0.30,  # 30% per period - very high!
        "warning": "Deliberate use = de facto bankruptcy"
    },
    
    # Table VIII.3 - Accounts Payable (Beginning of Period 7)
    "INITIAL_AP": {
        2: 17468,   # Due in fortnight 2
        3: 61630,   # Due in fortnight 3
        5: 11620,   # Due in fortnight 5
        6: 53250    # Due in fortnight 6
    },
    
    # Table VIII.4 - Accounts Receivable (Beginning of Period 7)
    "INITIAL_AR": {
        2: 295885.30  # Due in fortnight 2
    },
    
    # Table VIII.2 - Payment Schedule Summary
    "PAYMENT_TIMING": {
        "INITIAL": [
            "Supplier payments (pieces)", "Warehouse rental", "Equipment purchases",
            "Machine transfers", "Module leasing", "Product improvements",
            "Hiring/layoff costs", "Profit sharing", "Green investments",
            "Tax payments", "Mortgage interest", "Mortgage repayment",
            "Shares issued", "Dividend payments"
        ],
        "PER_FORTNIGHT": [
            "Energy costs", "Inventory holding", "Salaries", "Overtime",
            "Indirect labor", "Labor benefits", "Solar/tree maintenance",
            "Green energy", "Credit line interest", "Deposit interest"
        ],
        "SALES_FORTNIGHTS": [2, 4, 6, 8],  # Customer payments in these FNs
        "ORDERING_FORTNIGHTS": "Per order"  # Parts payments when ordered
    },
    
    # Cash flow timing constants
    "CASH_FLOW": {
        "initial_cash_p6": 500000,  # Approximate starting cash
        "min_cash_buffer": 50000    # Recommended safety buffer
    }
}

# =============================================================================
# 9. COMPANY REPORTS (Chapter IX)
# =============================================================================
FINANCIAL_STATEMENTS = {
    # Table IX.1 - Income Statement (Period 6 and Full Year 2)
    "INCOME_STATEMENT": {
        "NET_SALES": 1183541,
        "COGS": 481439,
        "GROSS_INCOME": 702101,
        "EXPENSES": {
            "warehouse": 74400,
            "freight": 64839,
            "installation": 0,
            "hiring_firing": 1100,
            "machine_rental": 0,
            "social": 0,
            "sales_admin": 316200,
            "energy": 29518,
            "co2_abatement": 0,
            "disposal": 0
        },
        "EBITDA": 216043,
        "DEPRECIATION": {
            "plant_equip": 60312,
            "improvements": 0,
            "esg": 0
        },
        "OPERATING_INCOME": 155730,
        "FINANCIAL_EXPENSES": {
            "credit_line_interest": 26037,
            "mortgage_interest": 30000,
            "emergency_interest": 0,
            "investment_income": 5000,
            "total_net": 51037
        },
        "NET_PROFIT_BEFORE_TAX": 104693,
        "TAXES": 52346,
        "NET_PROFIT": 52346
    },
    
    # Table IX.2 - Balance Sheet (End of Period 6)
    "BALANCE_SHEET": {
        "ASSETS": {
            "CURRENT": {
                "cash": 219615,
                "investments": 200000,
                "receivables": 295885,
                "inventory_rm": 70149,
                "inventory_wip": 308876,
                "inventory_fp": 132791,
                "total_inventory": 511817,
                "total_current": 1227318
            },
            "FIXED": {
                "plant_equip_gross": 1059500,
                "esg_gross": 0,
                "intangible_gross": 0,
                "accumulated_depreciation": 361875,
                "net_fixed": 697625
            },
            "TOTAL_ASSETS": 1924943
        },
        "LIABILITIES_EQUITY": {
            "LIABILITIES": {
                "payables": 143968,
                "credit_line": 113000,
                "interest_payable": 30000,
                "emergency_loan": 0,
                "mortgage_loans": 500000,
                "taxes_payable": 52346,
                "total_liabilities": 839314
            },
            "EQUITY": {
                "issued_capital": 850000,
                "retained_earnings": 183281,
                "period_profit": 52346,
                "total_equity": 1085628
            },
            "TOTAL_LIABILITIES_EQUITY": 1924943
        }
    }
}
