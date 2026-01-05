from pathlib import Path
import os

# Define project root (where this file is located)
ROOT_DIR = Path(__file__).parent.resolve()

# Define standard data directories
REPORTS_DIR = ROOT_DIR / "Reports"
DATA_DIR = ROOT_DIR / "data"

# Define output directory
OUTPUT_DIR = ROOT_DIR / "dashboards_v2"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def get_data_path(filename: str) -> Path:
    """
    Locate a data file in standard directories.
    Checks REPORTS_DIR first, then DATA_DIR.
    Returns the resolved Path if found.
    Raises FileNotFoundError if not found.
    """
    # 1. Check Reports folder
    report_path = REPORTS_DIR / filename
    if report_path.exists():
        return report_path
    
    # 2. Check data folder
    data_path = DATA_DIR / filename
    if data_path.exists():
        return data_path
    
    # If not found in either, return None to allow scripts to fallback to defaults
    return None
