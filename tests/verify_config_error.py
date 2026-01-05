import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent.parent))

try:
    from config import get_data_path, REPORTS_DIR, DATA_DIR
    import os
    
    filename = "NON_EXISTENT_FILE_XYZ_123.xlsx"
    print(f"Attempting to find: {filename}")
    
    try:
        path = get_data_path(filename)
        print(f"FAILED: Expected FileNotFoundError, but got: {path}")
        sys.exit(1)
    except FileNotFoundError as e:
        print(f"SUCCESS: Caught expected error: {e}")
        sys.exit(0)
    except Exception as e:
        print(f"FAILED: Caught unexpected exception: {type(e).__name__}: {e}")
        sys.exit(1)

except ImportError:
    print("FAILED: Could not import config.py")
    sys.exit(1)
