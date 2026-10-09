"""
Shared test fixtures for the Excel Intelligence System.
"""
import sys
from pathlib import Path

import pytest
import pandas as pd
import numpy as np

# Add project root to path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from tests.sample_data import generate_sales_data, save_as_excel


@pytest.fixture(scope="session")
def test_data_dir(tmp_path_factory):
    """Create a temporary directory with test data."""
    d = tmp_path_factory.mktemp("test_data")
    return d


@pytest.fixture(scope="session")
def sample_sales_df():
    """Return the sample sales DataFrame."""
    return generate_sales_data(num_rows=100, seed=42)


@pytest.fixture(scope="session")
def sample_excel_path(test_data_dir, sample_sales_df):
    """Create a sample sales.xlsx in the test data dir and return its path."""
    path = test_data_dir / "sales.xlsx"
    save_as_excel(sample_sales_df, path)
    return path


@pytest.fixture(scope="session")
def output_dir(test_data_dir):
    """Create an output directory for test results."""
    d = test_data_dir / "output"
    d.mkdir(exist_ok=True)
    return d
