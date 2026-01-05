import sys
import io
from pathlib import Path

# Add shared utils path
sys.path.append(str(Path(__file__).resolve().parent.parent))

# Redirect stderr to capture output
stderr_capture = io.StringIO()
original_stderr = sys.stderr
sys.stderr = stderr_capture

try:
    # Import a dashboard generator script to test load_excel_file
    # We will invoke the function directly if possible, or trigger a load failure
    
    # We'll use CPO dashboard's load_excel_file as a representative test
    # (Since we patched all of them identically)
    sys.path.append(str(Path(__file__).resolve().parent.parent / "CPO Dashboard"))
    import generate_cpo_dashboard as cpo_gen
    
    print("Testing load_excel_file with non-existent file...")
    result = cpo_gen.load_excel_file("NON_EXISTENT_FILE_FOR_TESTING.xlsx")
    
    # Check if stderr caught the error message
    captured_err = stderr_capture.getvalue()
    
    if "[ERROR]" in captured_err and "NON_EXISTENT_FILE_FOR_TESTING.xlsx" in captured_err:
        print("SUCCESS: Error correctly logged to stderr.")
    else:
        print("FAILED: Error message not found in stderr.")
        print(f"Captured Stderr: {captured_err}")

except Exception as e:
    print(f"FAILED: Unexpected exception during test: {e}")
finally:
    sys.stderr = original_stderr
