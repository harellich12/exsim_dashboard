
import os
import sys
import subprocess
from pathlib import Path

BASE_DIR = Path(__file__).parent

def run_script(script_path):
    print(f"[*] Running {script_path.name}...")
    try:
        result = subprocess.run(
            [sys.executable, str(script_path)],
            cwd=str(script_path.parent),
            capture_output=True,
            text=True,
            timeout=120
        )
        if result.returncode == 0:
            print(f"  [OK] {script_path.name} completed successfully.")
            return True
        else:
            print(f"  [FAIL] {script_path.name} failed.")
            print(result.stderr)
            return False
    except Exception as e:
        print(f"  [ERROR] {e}")
        return False

def main():
    print("=" * 60)
    print("ExSim Full Suite Runner")
    print("=" * 60)

    # 1. Generators
    generators = [
        BASE_DIR / "CFO Dashboard/generate_finance_dashboard_final.py",
        BASE_DIR / "CLO Dashboard/generate_logistics_dashboard.py",
        BASE_DIR / "CPO Dashboard/generate_cpo_dashboard.py",
        BASE_DIR / "CMO Dashboard/generate_cmo_dashboard_complete.py",
        BASE_DIR / "Purchasing Role/generate_purchasing_dashboard_v2.py",
        BASE_DIR / "ESG Dashboard/generate_esg_dashboard.py",
        BASE_DIR / "Production Manager Dashboard/generate_production_dashboard_zones.py"
    ]

    gen_success = True
    print("\n--- GENERATION PHASE ---")
    for gen in generators:
        if not run_script(gen):
            gen_success = False

    if not gen_success:
        print("\n[!] Generation phase encountered errors. Proceeding to tests might be unsafe.")
        # We continue anyway to see what fails

    # 2. Validation
    validators = [
        BASE_DIR / "validate_dashboards.py",
        BASE_DIR / "self_test_dashboards.py",
        BASE_DIR / "verify_integrity_suite.py"
    ]

    val_success = True
    print("\n--- VERIFICATION PHASE ---")
    for val in validators:
        if not run_script(val):
            val_success = False

    print("\n" + "=" * 60)
    if gen_success and val_success:
        print("OVERALL STATUS: SUCCESS")
        sys.exit(0)
    else:
        print("OVERALL STATUS: FAILURE")
        sys.exit(1)

if __name__ == "__main__":
    main()
