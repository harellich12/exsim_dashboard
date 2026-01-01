"""
Shared pytest fixtures for ExSim Dashboard tests.
"""
import pytest
import tempfile
import shutil
from pathlib import Path


@pytest.fixture
def temp_output_dir():
    """Create a temporary directory for test output files."""
    temp_dir = tempfile.mkdtemp(prefix="exsim_test_")
    yield Path(temp_dir)
    # Cleanup after test
    shutil.rmtree(temp_dir, ignore_errors=True)


@pytest.fixture
def mock_reports_path():
    """Return path to existing mock reports for integration tests."""
    return Path(__file__).parent.parent / "test_data" / "mock_reports"
