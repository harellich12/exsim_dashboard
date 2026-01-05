"""
ExSim Shared Outputs - Inter-Dashboard Data Exchange

Provides a standardized way for dashboards to share their outputs
with other dashboards in the cascade:

    CMO (demand) → Production (plan) → Purchasing (MRP) → CLO (shipping) → CFO (cash flow)

This module enables the cascade dashboard pattern by allowing each
dashboard to export key outputs and import upstream dependencies.

Usage:
    # In a dashboard script:
    from shared_outputs import SharedOutputManager
    
    # Export data after generating dashboard
    manager = SharedOutputManager()
    manager.export('CMO', {
        'demand_forecast': {...},
        'marketing_spend': 50000
    })
    
    # Import upstream data before generating dashboard
    cmo_data = manager.import_data('CMO')
"""

import json
from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime


# Default path for shared outputs file
SHARED_OUTPUTS_FILE = Path(__file__).parent / "shared_outputs.json"


class SharedOutputManager:
    """Manages inter-dashboard data exchange via JSON file."""
    
    # Define which dashboards can read from which other dashboards
    DEPENDENCY_GRAPH = {
        'CMO': [],  # CMO is typically the starting point
        'Production': ['CMO'],  # Production reads demand from CMO
        'Purchasing': ['Production'],  # Purchasing reads production plan
        'CLO': ['Production', 'CMO'],  # CLO reads production output and CMO demand
        'CPO': ['Production'],  # CPO reads production for workforce sizing
        'ESG': ['Production'],  # ESG reads production for emissions
        'CFO': ['CMO', 'Production', 'Purchasing', 'CLO', 'CPO', 'ESG'],  # CFO aggregates all
    }
    
    # Define the expected output keys for each dashboard
    OUTPUT_SCHEMA = {
        'CMO': ['demand_forecast', 'marketing_spend', 'pricing', 'innovation_costs'],
        'Production': ['production_plan', 'capacity_utilization', 'overtime_hours', 'unit_costs'],
        'Purchasing': ['material_orders', 'supplier_spend', 'lead_time_schedule'],
        'CLO': ['shipping_schedule', 'logistics_costs', 'inventory_by_zone'],
        'CPO': ['workforce_headcount', 'payroll_forecast', 'hiring_costs'],
        'ESG': ['co2_emissions', 'abatement_investment', 'tax_liability'],
        'CFO': ['cash_flow_projection', 'debt_levels', 'liquidity_status'],
    }
    
    def __init__(self, filepath: Optional[Path] = None):
        """Initialize the manager with optional custom filepath."""
        self.filepath = filepath or SHARED_OUTPUTS_FILE
        self._ensure_file_exists()
    
    def _ensure_file_exists(self):
        """Ensure the shared outputs file exists with proper structure."""
        if not self.filepath.exists():
            self._write_data({
                'metadata': {
                    'created': datetime.now().isoformat(),
                    'last_updated': datetime.now().isoformat(),
                    'version': '1.0'
                },
                'dashboards': {}
            })
    
    def _read_data(self) -> Dict[str, Any]:
        """Read the shared outputs file."""
        try:
            with open(self.filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, FileNotFoundError):
            return {'metadata': {}, 'dashboards': {}}
    
    def _write_data(self, data: Dict[str, Any]):
        """Write data to the shared outputs file."""
        data['metadata']['last_updated'] = datetime.now().isoformat()
        with open(self.filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, default=str)
    
    def export(self, dashboard_name: str, outputs: Dict[str, Any]) -> bool:
        """
        Export dashboard outputs to the shared file.
        
        Args:
            dashboard_name: Name of the dashboard (e.g., 'CMO', 'Production')
            outputs: Dictionary of output values to share
        
        Returns:
            True if successful, False otherwise
        """
        if dashboard_name not in self.OUTPUT_SCHEMA:
            print(f"Warning: Unknown dashboard '{dashboard_name}'")
            return False
        
        data = self._read_data()
        
        data['dashboards'][dashboard_name] = {
            'timestamp': datetime.now().isoformat(),
            'outputs': outputs
        }
        
        self._write_data(data)
        print(f"[SHARED] Exported {len(outputs)} keys from {dashboard_name}")
        return True
    
    def import_data(self, dashboard_name: str) -> Optional[Dict[str, Any]]:
        """
        Import outputs from a specific dashboard.
        
        Args:
            dashboard_name: Name of the dashboard to import from
        
        Returns:
            Dictionary of outputs, or None if not available
        """
        data = self._read_data()
        dashboard_data = data.get('dashboards', {}).get(dashboard_name, {})
        return dashboard_data.get('outputs')
    
    def import_dependencies(self, dashboard_name: str) -> Dict[str, Dict[str, Any]]:
        """
        Import all outputs from dashboards that this dashboard depends on.
        
        Args:
            dashboard_name: Name of the dashboard requesting dependencies
        
        Returns:
            Dictionary mapping upstream dashboard names to their outputs
        """
        dependencies = self.DEPENDENCY_GRAPH.get(dashboard_name, [])
        result = {}
        
        for dep in dependencies:
            dep_data = self.import_data(dep)
            if dep_data:
                result[dep] = dep_data
            else:
                print(f"[SHARED] Warning: No data available from {dep}")
        
        return result
    
    def clear(self):
        """Clear all shared outputs (for reset/testing)."""
        self._write_data({
            'metadata': {
                'created': datetime.now().isoformat(),
                'last_updated': datetime.now().isoformat(),
                'version': '1.0'
            },
            'dashboards': {}
        })
        print("[SHARED] Cleared all shared outputs")
    
    def get_status(self) -> Dict[str, str]:
        """Get status of all dashboards' shared data."""
        data = self._read_data()
        status = {}
        
        for name in self.OUTPUT_SCHEMA.keys():
            dashboard_data = data.get('dashboards', {}).get(name, {})
            if dashboard_data:
                timestamp = dashboard_data.get('timestamp', 'unknown')
                keys = list(dashboard_data.get('outputs', {}).keys())
                status[name] = f"[OK] {len(keys)} keys @ {timestamp[:16]}"
            else:
                status[name] = "[--] No data"
        
        return status


# Convenience functions for quick access
def export_dashboard_data(dashboard_name: str, outputs: Dict[str, Any]) -> bool:
    """Quick export function."""
    return SharedOutputManager().export(dashboard_name, outputs)


def import_dashboard_data(dashboard_name: str) -> Optional[Dict[str, Any]]:
    """Quick import function."""
    return SharedOutputManager().import_data(dashboard_name)


def get_all_status() -> Dict[str, str]:
    """Get status of all shared dashboard data."""
    return SharedOutputManager().get_status()


# Example of cascade execution order
EXECUTION_ORDER = [
    'CMO',        # 1. Marketing forecasts demand
    'Production', # 2. Production plans to meet demand
    'Purchasing', # 3. Purchasing sources materials for production
    'CLO',        # 4. Logistics ships finished goods
    'CPO',        # 5. HR manages workforce for production
    'ESG',        # 6. Sustainability tracks emissions from production
    'CFO',        # 7. Finance aggregates all cash flows
]


if __name__ == "__main__":
    # Demo usage
    print("Shared Outputs Manager - Status Check")
    print("=" * 50)
    
    manager = SharedOutputManager()
    status = manager.get_status()
    
    for dashboard, state in status.items():
        print(f"  {dashboard}: {state}")
    
    print("\nDependency Graph:")
    for dashboard, deps in manager.DEPENDENCY_GRAPH.items():
        if deps:
            print(f"  {dashboard} ← {', '.join(deps)}")
        else:
            print(f"  {dashboard} (no dependencies)")
