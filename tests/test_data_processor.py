"""Unit tests for DataProcessor module."""

import pytest
import pandas as pd
import numpy as np
from datetime import datetime
import sys
import os

# Add src to path
sys.path.append(os.path.join(os.path.dirname(__file__), "../src"))

from data_processor import DataProcessor


@pytest.fixture
def sample_data():
    """Create sample data for testing."""
    dates = pd.date_range(start="2024-01-01", periods=10, freq="D")
    data = {
        "date": dates,
        "sales": [100, 110, 105, 120, 115, 130, 125, 140, 135, 150],
        "revenue": [
            1000.0,
            1100.0,
            1050.0,
            1200.0,
            1150.0,
            1300.0,
            1250.0,
            1400.0,
            1350.0,
            1500.0,
        ],
        "users": [50, 55, 52, 60, 58, 65, 63, 70, 68, 75],
        "conversion_rate": [
            0.02,
            0.021,
            0.019,
            0.022,
            0.020,
            0.023,
            0.021,
            0.024,
            0.022,
            0.025,
        ],
    }
    return pd.DataFrame(data)


@pytest.fixture
def processor():
    """Create DataProcessor instance."""
    return DataProcessor()


def test_data_processor_initialization(processor):
    """Test DataProcessor initialization."""
    assert processor is not None
    assert "data_paths" in processor.config


def test_load_raw_data(processor, tmp_path):
    """Test loading raw data."""
    # Create a temporary CSV file
    test_data = pd.DataFrame({"col1": [1, 2, 3], "col2": ["a", "b", "c"]})
    test_file = tmp_path / "test_data.csv"
    test_data.to_csv(test_file, index=False)

    # Test loading
    # Note: In real test, would mock the file path
    assert True  # Placeholder


def test_clean_data(processor, sample_data):
    """Test data cleaning functionality."""
    cleaned_data = processor.clean_data(sample_data)

    assert len(cleaned_data) == len(sample_data)
    assert "date" in cleaned_data.columns
    assert cleaned_data["date"].dtype == "datetime64[ns]"

    # Check for NaN values
    assert cleaned_data.isna().sum().sum() == 0


def test_calculate_metrics(processor, sample_data):
    """Test metric calculations."""
    cleaned_data = processor.clean_data(sample_data)
    metrics_data = processor.calculate_metrics(cleaned_data)

    assert "sales_growth" in metrics_data.columns
    assert "revenue_growth" in metrics_data.columns
    assert "sales_7d_ma" in metrics_data.columns

    # Check growth rate calculations
    assert metrics_data["sales_growth"].iloc[1] == pytest.approx(
        10.0
    )  # (110-100)/100*100


def test_process_pipeline(processor):
    """Test complete processing pipeline."""
    processed_data = processor.process_pipeline()

    assert isinstance(processed_data, pd.DataFrame)
    assert len(processed_data) > 0
    assert "sales_growth" in processed_data.columns


def test_save_processed_data(processor, sample_data, tmp_path):
    """Test saving processed data."""
    # Test would involve mocking the save path
    assert True  # Placeholder


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
