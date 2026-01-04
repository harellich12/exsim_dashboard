"""
ExSim Dashboard Suite - Full Cascade Execution

Runs all dashboards in dependency order, passing shared outputs between them:
    CMO → Production → Purchasing → CLO → CPO → ESG → CFO

Usage:
    python run_full_suite.py
    
Options:
    --clean     Clear shared outputs before running
    --verify    Verify all outputs exist after running
"""

import sys
import subprocess
from pathlib import Path
from datetime import datetime

# Dashboard execution order (based on dependency graph)
EXECUTION_ORDER = [
    ("CMO", "CMO Dashboard", "generate_cmo_dashboard_complete.py"),
    ("Production", "Production Manager Dashboard", "generate_production_dashboard_zones.py"),
    ("Purchasing", "Purchasing Role", "generate_purchasing_dashboard_v2.py"),
    ("CLO", "CLO Dashboard", "generate_logistics_dashboard.py"),
    ("CPO", "CPO Dashboard", "generate_cpo_dashboard.py"),
    ("ESG", "ESG Dashboard", "generate_esg_dashboard.py"),
    ("CFO", "CFO Dashboard", "generate_finance_dashboard_final.py"),
]

# Expected output files
EXPECTED_OUTPUTS = {
    "CMO": "CMO_Dashboard_Complete.xlsx",
    "Production": "Production_Dashboard_Zones.xlsx",
    "Purchasing": "Purchasing_Dashboard.xlsx",
    "CLO": "Logistics_Dashboard.xlsx",
    "CPO": "CPO_Dashboard.xlsx",
    "ESG": "ESG_Dashboard.xlsx",
    "CFO": "Finance_Dashboard_Final.xlsx",
}


def run_dashboard(name, folder, script):
    """Run a single dashboard generator script."""
    base_path = Path(__file__).parent
    script_path = base_path / folder / script
    
    if not script_path.exists():
        print(f"  [ERROR] Script not found: {script_path}")
        return False
    
    try:
        result = subprocess.run(
            [sys.executable, str(script_path)],
            cwd=str(script_path.parent),
            capture_output=True,
            text=True,
            timeout=120  # 2 minute timeout per dashboard
        )
        
        if result.returncode == 0:
            # Check for SUCCESS in output
            if "[SUCCESS]" in result.stdout:
                return True
            else:
                print(f"  [WARN] No SUCCESS message found")
                return True  # Still consider success if no error
        else:
            print(f"  [ERROR] Exit code {result.returncode}")
            if result.stderr:
                print(f"  {result.stderr[:200]}")
            return False
            
    except subprocess.TimeoutExpired:
        print(f"  [ERROR] Timeout (>120s)")
        return False
    except Exception as e:
        print(f"  [ERROR] {str(e)}")
        return False


def verify_outputs():
    """Verify all expected output files exist."""
    base_path = Path(__file__).parent
    results = {}
    
    for name, folder, _ in EXECUTION_ORDER:
        expected_file = EXPECTED_OUTPUTS.get(name)
        if expected_file:
            output_path = base_path / folder / expected_file
            results[name] = output_path.exists()
    
    return results


def clear_shared_outputs():
    """Clear the shared outputs file."""
    try:
        from shared_outputs import SharedOutputManager
        manager = SharedOutputManager()
        manager.clear()
        return True
    except Exception as e:
        print(f"[WARN] Could not clear shared outputs: {e}")
        return False


def get_shared_status():
    """Get status of shared outputs."""
    try:
        from shared_outputs import get_all_status
        return get_all_status()
    except Exception:
        return {}


def main():
    """Run the full dashboard suite in cascade order."""
    print("=" * 60)
    print("  ExSim Dashboard Suite - Cascade Execution")
    print("=" * 60)
    print(f"  Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    # Parse arguments
    clean_mode = "--clean" in sys.argv
    verify_mode = "--verify" in sys.argv
    
    if clean_mode:
        print("[*] Clearing shared outputs...")
        clear_shared_outputs()
        print()
    
    # Run dashboards in order
    print("[*] Executing dashboards in cascade order:")
    print()
    
    results = {}
    for i, (name, folder, script) in enumerate(EXECUTION_ORDER, 1):
        print(f"  [{i}/7] {name}...", end=" ", flush=True)
        success = run_dashboard(name, folder, script)
        results[name] = success
        print("[OK]" if success else "[FAIL]")
    
    print()
    
    # Summary
    success_count = sum(1 for v in results.values() if v)
    print(f"[*] Results: {success_count}/7 dashboards completed successfully")
    
    if verify_mode:
        print()
        print("[*] Verifying outputs...")
        outputs = verify_outputs()
        for name, exists in outputs.items():
            status = "[OK]" if exists else "[MISSING]"
            print(f"    {name}: {EXPECTED_OUTPUTS[name]} {status}")
    
    # Show shared data status
    print()
    print("[*] Shared Data Status:")
    status = get_shared_status()
    for name, state in status.items():
        print(f"    {name}: {state}")
    
    print()
    print(f"  Completed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)
    
    # Return exit code
    return 0 if success_count == 7 else 1


if __name__ == "__main__":
    sys.exit(main())
