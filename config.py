from pathlib import Path
import os

# Define project root (where this file is located)
ROOT_DIR = Path(__file__).parent.resolve()

# Define standard data directories
REPORTS_DIR = ROOT_DIR / "Reports"
DATA_DIR = ROOT_DIR / "data"
TEMPLATES_DIR = REPORTS_DIR / "Decision Templates"

# Define output directory
OUTPUT_DIR = ROOT_DIR / "dashboards_v2"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

def get_data_path(filename: str, required: bool = True) -> Path:
    """
    Locate a data file in standard directories.
    Checks REPORTS_DIR, TEMPLATES_DIR, then DATA_DIR.
    
    Args:
        filename: Name of the file to find
        required: If True, raises FileNotFoundError when not found.
                  If False, returns None when not found.
    
    Returns the resolved Path if found.
    """
    # 1. Check Reports folder
    report_path = REPORTS_DIR / filename
    if report_path.exists():
        return report_path
    
    # 2. Check Decision Templates subfolder
    template_path = TEMPLATES_DIR / filename
    if template_path.exists():
        return template_path
    
    # 3. Check data folder
    data_path = DATA_DIR / filename
    if data_path.exists():
        return data_path
    
    # If not found in any location
    if required:
        raise FileNotFoundError(f"Could not find {filename} in {REPORTS_DIR}, {TEMPLATES_DIR}, or {DATA_DIR}")
    else:
        return None

