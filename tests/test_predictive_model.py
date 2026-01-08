"""
Unit tests for PredictiveModel module.
"""

import pytest
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import sys
import os

# Add src to path
sys.path.append(os.path.join(os.path.dirname(__file__), "../src"))

from predictive_model import PredictiveModel
from data_processor import DataProcessor


@pytest.fixture
def sample_data():
    """Create sample data for testing."""
    dates = pd.date_range(start="2024-01-01", periods=20, freq="D")
    data = {
        "date": dates,
        "sales": [
            100,
            110,
            105,
            120,
            115,
            130,
            125,
            140,
            135,
            150,
            145,
            160,
            155,
            170,
            165,
            180,
            175,
            190,
            185,
            200,
        ],
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
            1450.0,
            1600.0,
            1550.0,
            1700.0,
            1650.0,
            1800.0,
            1750.0,
            1900.0,
            1850.0,
            2000.0,
        ],
        "users": [
            50,
            55,
            52,
            60,
            58,
            65,
            63,
            70,
            68,
            75,
            73,
            80,
            78,
            85,
            83,
            90,
            88,
            95,
            93,
            100,
        ],
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
            0.024,
            0.027,
            0.025,
            0.028,
            0.026,
            0.029,
            0.027,
            0.030,
            0.028,
            0.031,
        ],
    }
    df = pd.DataFrame(data)
    return df


@pytest.fixture
def processor():
    """Create DataProcessor instance."""
    return DataProcessor()


@pytest.fixture
def model():
    """Create PredictiveModel instance."""
    return PredictiveModel()


def test_predictive_model_initialization(model):
    """Test PredictiveModel initialization."""
    assert model is not None
    assert hasattr(model, "model")
    assert hasattr(model, "is_trained")
    assert model.is_trained == False


def test_prepare_features(model, sample_data):
    """Test feature preparation."""
    X, y, feature_names = model.prepare_features(sample_data, target_column="sales")

    assert X is not None
    assert y is not None
    assert len(feature_names) > 0
    assert len(X) == len(y)

    # Check that features are created
    assert X.shape[1] > 0
    assert "day_of_week" in feature_names or "lag_1" in feature_names


def test_train_model(model, sample_data):
    """Test model training."""
    metrics = model.train(sample_data, target_column="sales", test_size=0.2)

    assert model.is_trained == True
    assert metrics is not None
    assert "r2_test" in metrics
    assert "mse_test" in metrics
    assert "feature_importance" in metrics

    # Check metrics are reasonable
    assert isinstance(metrics["r2_test"], float)
    assert isinstance(metrics["mse_test"], float)
    assert isinstance(metrics["feature_importance"], dict)


def test_train_model_with_different_target(model, sample_data):
    """Test model training with different target column."""
    metrics = model.train(sample_data, target_column="revenue", test_size=0.3)

    assert model.is_trained == True
    assert metrics["r2_test"] <= 1.0  # R² should be <= 1
    assert metrics["r2_test"] >= -1.0  # R² should be >= -1


def test_forecast_without_training(model, sample_data):
    """Test that forecast raises error without training."""
    with pytest.raises(ValueError, match="Model must be trained before forecasting"):
        model.forecast(sample_data, periods=7, target_column="sales")


def test_forecast_with_training(model, sample_data):
    """Test forecasting after training."""
    # First train the model
    model.train(sample_data, target_column="sales")

    # Then forecast
    forecast_df = model.forecast(sample_data, periods=7, target_column="sales")

    assert forecast_df is not None
    assert len(forecast_df) == 7
    assert "date" in forecast_df.columns
    assert "predicted_sales" in forecast_df.columns
    assert "confidence_interval_lower" in forecast_df.columns
    assert "confidence_interval_upper" in forecast_df.columns

    # Check dates are in future
    last_date = sample_data["date"].max()
    forecast_dates = forecast_df["date"]
    assert all(forecast_dates > last_date)


def test_forecast_different_periods(model, sample_data):
    """Test forecasting with different period lengths."""
    model.train(sample_data, target_column="sales")

    # Test 3-day forecast
    forecast_3 = model.forecast(sample_data, periods=3, target_column="sales")
    assert len(forecast_3) == 3

    # Test 14-day forecast
    forecast_14 = model.forecast(sample_data, periods=14, target_column="sales")
    assert len(forecast_14) == 14


def test_save_forecast(model, sample_data, tmp_path):
    """Test saving forecast results."""
    import tempfile
    import os

    model.train(sample_data, target_column="sales")
    forecast_df = model.forecast(sample_data, periods=7, target_column="sales")

    # Create temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        output_path = os.path.join(temp_dir, "test_forecast.csv")

        # Save forecast
        forecast_df.to_csv(output_path, index=False)

        # Check file was created
        assert os.path.exists(output_path)

        # Load and verify
        loaded_df = pd.read_csv(output_path)
        assert len(loaded_df) == len(forecast_df)
        assert "predicted_sales" in loaded_df.columns


def test_model_metrics_range(model, sample_data):
    """Test that model metrics are within reasonable ranges."""
    metrics = model.train(sample_data, target_column="sales")

    # R² should be between -1 and 1
    assert -1.0 <= metrics["r2_test"] <= 1.0

    # MSE should be positive
    assert metrics["mse_test"] >= 0
    assert metrics["rmse_test"] >= 0

    # Should have feature importance
    assert len(metrics["feature_importance"]) > 0


def test_feature_importance_structure(model, sample_data):
    """Test feature importance structure."""
    metrics = model.train(sample_data, target_column="sales")

    feature_importance = metrics["feature_importance"]

    # Should be a dictionary
    assert isinstance(feature_importance, dict)

    # Should have feature names as keys
    for feature_name in feature_importance.keys():
        assert isinstance(feature_name, str)

    # Should have coefficients as values
    for coefficient in feature_importance.values():
        assert isinstance(coefficient, (int, float, np.floating))


def test_forecast_confidence_intervals(model, sample_data):
    """Test that confidence intervals are reasonable."""
    model.train(sample_data, target_column="sales")
    forecast_df = model.forecast(sample_data, periods=7, target_column="sales")

    # Check confidence intervals
    for idx, row in forecast_df.iterrows():
        lower = row["confidence_interval_lower"]
        upper = row["confidence_interval_upper"]
        predicted = row["predicted_sales"]

        # Lower bound should be less than or equal to predicted
        assert lower <= predicted

        # Upper bound should be greater than or equal to predicted
        assert upper >= predicted

        # Bounds should be positive (for sales)
        assert lower >= 0
        assert upper >= 0


def test_empty_data_handling(model):
    """Test model behavior with empty data."""
    empty_df = pd.DataFrame()

    with pytest.raises(Exception):
        model.train(empty_df, target_column="sales")


def test_insufficient_data_handling(model):
    """Test model with insufficient data."""
    # Create DataFrame with only 2 rows (insufficient for training)
    small_df = pd.DataFrame(
        {
            "date": pd.date_range(start="2024-01-01", periods=2, freq="D"),
            "sales": [100, 110],
        }
    )

    with pytest.raises(Exception):
        model.train(small_df, target_column="sales")


def test_plot_forecast(model, sample_data, tmp_path):
    """Test forecast plotting functionality."""
    import matplotlib

    matplotlib.use("Agg")  # Use non-interactive backend

    model.train(sample_data, target_column="sales")
    forecast_df = model.forecast(sample_data, periods=7, target_column="sales")

    # Create plot
    fig = model.plot_forecast(sample_data, forecast_df, target_column="sales")

    assert fig is not None
    assert isinstance(fig, matplotlib.figure.Figure)


def test_end_to_end_pipeline(processor, model):
    """Test complete pipeline from data processing to forecasting."""
    # Process data
    processed_data = processor.process_pipeline()
    assert len(processed_data) > 0

    # Train model
    metrics = model.train(processed_data, target_column="sales")
    assert model.is_trained == True

    # Generate forecast
    forecast = model.forecast(processed_data, periods=7, target_column="sales")
    assert len(forecast) == 7

    # Save forecast
    model.save_forecast(forecast, filename="test_forecast_end2end.csv")

    # Check file was created
    forecast_path = "data/forecasts/test_forecast_end2end.csv"
    assert os.path.exists(forecast_path)

    # Clean up
    if os.path.exists(forecast_path):
        os.remove(forecast_path)


def test_model_persistence(model, sample_data, tmp_path):
    """Test that model can be retrained and produces consistent results."""
    # Train first time
    metrics1 = model.train(sample_data, target_column="sales")
    forecast1 = model.forecast(sample_data, periods=3, target_column="sales")

    # Create new model instance and retrain
    model2 = PredictiveModel()
    metrics2 = model2.train(sample_data, target_column="sales")
    forecast2 = model2.forecast(sample_data, periods=3, target_column="sales")

    # Results should be similar (not necessarily identical due to random split)
    assert abs(metrics1["r2_test"] - metrics2["r2_test"]) < 0.5
    assert len(forecast1) == len(forecast2)


if __name__ == "__main__":
    # Run tests directly
    pytest.main([__file__, "-v", "--tb=short"])
