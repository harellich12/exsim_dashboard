"""
ExSim Dashboards - Shared Data Path Helper
Centralized data source configuration for all dashboard scripts.
"""

from pathlib import Path


def get_data_path(filename, script_path):
    """
    Get data file path, checking Reports folder first, then local fallback.
    
    Args:
        filename: Name of the Excel file to find
        script_path: __file__ from the calling script
        
    Returns:
        Path object if file found, None otherwise
    """
    script_dir = Path(script_path).parent
    reports_folder = script_dir.parent / "Reports"
    local_folder = script_dir / "data"
    
    primary = reports_folder / filename
    fallback = local_folder / filename
    
    if primary.exists():
        return primary
    elif fallback.exists():
        return fallback
    return None


def print_data_sources(script_path):
    """Print data source paths for logging."""
    script_dir = Path(script_path).parent
    reports_folder = script_dir.parent / "Reports"
    local_folder = script_dir / "data"
    
    print(f"    Primary source: {reports_folder}")
    print(f"    Fallback source: {local_folder}")
