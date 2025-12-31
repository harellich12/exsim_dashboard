# ExSim Dashboard Suite

A comprehensive collection of decision-support dashboards for the ExSim business simulation. Each dashboard is designed to optimize a specific functional area while maintaining cross-functional integration.

## 📊 Dashboard Overview

### 1. [CMO Dashboard](CMO%20Dashboard/README.md) (Market Allocation & Strategy)

**Function:** Calculates "True Demand" (adjusting for previous stockouts), optimizes the 4Ps for High/Low segments, and manages Innovation/R&D to ensure the market allocation engine chooses you.

### 2. [Production Dashboard](Produciton%20Manager%20Dashboard/README.md) (Capacity & Cost Optimization)

**Function:** Converts the Marketing forecast into a feasible build plan per Zone, assigns machines/workers to specific sections, and calculates the "Real Unit Cost" to flag unprofitable overtime.

### 3. [Purchasing Dashboard](Purchasing%20Role/README.md) (MRP & Sourcing)

**Function:** Translates the Production plan into specific supplier orders, managing "Time Travel" (Lead Times) and optimizing Batch Sizes (EOQ) to minimize ordering and holding costs.

### 4. [Logistics Dashboard](CLO%20Dashboard/README.md) (Supply Network Optimization)

**Function:** Plays "Inventory Tetris" by scheduling shipments between Zones to balance supply with local demand, preventing stockouts and minimizing expensive warehouse overflow.

### 5. [CPO Dashboard](CPO%20Dashboard/README.md) (Workforce Planning & Compensation)

**Function:** Calculates hiring needs based on turnover/production targets and optimizes Salary/Benefit levels to prevent strikes (Inflation matching) while tracking total payroll cash flow.

### 6. [ESG Dashboard](ESG%20Dashboard/README.md) (CO2 Abatement Strategy)

**Function:** Performs a financial Cost-Benefit Analysis to decide if it is cheaper to invest in Green Tech (Solar/Credits) or simply pay the CO2 Tax.

### 7. [CFO Dashboard](CFO%20Dashboard/README.md) (Financial Control & Liquidity)

**Function:** Aggregates all spending plans to forecast solvency (Cash Flow), audits profitability (Income Statement), and manages Debt/Equity ratios to maintain credit ratings.

---

## 🚀 Quick Start

### Prerequisites

```bash
pip install pandas openpyxl
```

### Generate All Dashboards

```bash
# Generate each dashboard
python "CFO Dashboard/generate_finance_dashboard_final.py"
python "CLO Dashboard/generate_logistics_dashboard.py"
python "CPO Dashboard/generate_cpo_dashboard.py"
python "CMO Dashboard/generate_cmo_dashboard_complete.py"
python "Purchasing Role/generate_purchasing_dashboard_v2.py"
python "ESG Dashboard/generate_esg_dashboard.py"
```

### Run Validation Tests

```bash
# Run structural validation (27 tests)
python validate_dashboards.py

# Run column-by-column formula verification (187 tests)
python self_test_dashboards.py
```

---

## 📁 Project Structure

```
EXSIM models/
├── CFO Dashboard/
│   ├── generate_finance_dashboard_final.py
│   ├── Finance_Dashboard_Final.xlsx
│   └── data/
├── CLO Dashboard/
│   ├── generate_logistics_dashboard.py
│   ├── Logistics_Dashboard.xlsx
│   └── data/
├── CPO Dashboard/
│   ├── generate_cpo_dashboard.py
│   ├── CPO_Dashboard.xlsx
│   └── data/
├── CMO Dashboard/
│   ├── generate_cmo_dashboard_complete.py
│   ├── CMO_Dashboard_Complete.xlsx
│   └── data/
├── Purchasing Role/
│   ├── generate_purchasing_dashboard_v2.py
│   ├── Purchasing_Dashboard.xlsx
│   └── data/
├── ESG Dashboard/
│   ├── generate_esg_dashboard.py
│   ├── ESG_Dashboard.xlsx
│   └── data/
├── validate_dashboards.py
└── self_test_dashboards.py
```

---

## ✅ Test Coverage

| Dashboard | Validation Tests | Self-Tests |
|-----------|------------------|------------|
| CFO | Liquidity cascade, UPLOAD refs | 18 formula checks |
| CLO | Zone config, inventory flags | 15 formula checks |
| CPO | Hiring/firing, strike risk | 75 formula checks |
| CMO | Strategy cockpit, innovation | 29 formula checks |
| Purchasing | MRP cascade, cumulative spend | 10 formula checks |
| ESG | CO2 abatement, payback calc | 40 formula checks |

**Total: 214 automated tests**

---

## 📝 License

This project is for educational use with the ExSim business simulation.
