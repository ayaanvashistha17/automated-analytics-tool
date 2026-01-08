from datetime import datetime

"""
Data processing module for cleaning and preparing data for analysis.
"""

import pandas as pd
import numpy as np
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class DataProcessor:
    """Process and clean raw data for analysis."""

    def __init__(self, config_path="config/config.yaml"):
        """
        Initialize DataProcessor with configuration.

        Args:
            config_path (str): Path to configuration file
        """
        self.config = self._load_config(config_path)
        logger.info("DataProcessor initialized")

    def _load_config(self, config_path):
        """Load configuration file."""
        # Simplified config for demo
        return {
            "data_paths": {"raw": "data/raw/", "processed": "data/processed/"},
            "date_columns": ["date", "timestamp", "created_at"],
            "numeric_columns": ["sales", "revenue", "users", "conversion_rate"],
        }

    def load_raw_data(self, filename="daily_metrics.csv"):
        """
        Load raw data from CSV file.

        Args:
            filename (str): Name of the CSV file

        Returns:
            pd.DataFrame: Loaded data
        """
        try:
            filepath = f"{self.config['data_paths']['raw']}{filename}"
            df = pd.read_csv(filepath)
            logger.info(f"Loaded data from {filepath}")
            return df
        except FileNotFoundError:
            logger.warning(f"File {filename} not found. Generating sample data.")
            return self._generate_sample_data()

    def _generate_sample_data(self):
        """Generate sample data for demonstration."""
        dates = pd.date_range(start="2024-01-01", end="2024-01-30", freq="D")
        data = {
            "date": dates,
            "sales": np.random.randint(100, 500, len(dates)),
            "revenue": np.random.uniform(1000, 5000, len(dates)),
            "users": np.random.randint(50, 200, len(dates)),
            "conversion_rate": np.random.uniform(0.01, 0.05, len(dates)),
        }
        df = pd.DataFrame(data)
        df.to_csv(f"{self.config['data_paths']['raw']}daily_metrics.csv", index=False)
        logger.info("Generated sample data")
        return df

    def clean_data(self, df):
        """
        Clean and prepare data for analysis.

        Args:
            df (pd.DataFrame): Raw data

        Returns:
            pd.DataFrame: Cleaned data
        """
        # Make a copy to avoid modifying original
        df_clean = df.copy()

        # Convert date columns
        for col in self.config["date_columns"]:
            if col in df_clean.columns:
                df_clean[col] = pd.to_datetime(df_clean[col], errors="coerce")

        # Handle missing values
        for col in self.config["numeric_columns"]:
            if col in df_clean.columns:
                df_clean[col] = pd.to_numeric(df_clean[col], errors="coerce")
                # Fill missing with forward fill then backward fill
                df_clean[col] = df_clean[col].ffill().bfill()

        # Remove duplicates
        df_clean = df_clean.drop_duplicates()

        logger.info(
            f"Cleaned data: {len(df_clean)} rows, {len(df_clean.columns)} columns"
        )
        return df_clean

    def calculate_metrics(self, df):
        """
        Calculate derived metrics.

        Args:
            df (pd.DataFrame): Cleaned data

        Returns:
            pd.DataFrame: Data with calculated metrics
        """
        df_metrics = df.copy()

        # Calculate daily growth rates
        if "sales" in df_metrics.columns:
            df_metrics["sales_growth"] = df_metrics["sales"].pct_change() * 100

        if "revenue" in df_metrics.columns:
            df_metrics["revenue_growth"] = df_metrics["revenue"].pct_change() * 100

        # Calculate moving averages
        if "sales" in df_metrics.columns:
            df_metrics["sales_7d_ma"] = df_metrics["sales"].rolling(window=7).mean()

        # Calculate cumulative metrics
        if "revenue" in df_metrics.columns:
            df_metrics["cumulative_revenue"] = df_metrics["revenue"].cumsum()

        logger.info("Calculated derived metrics")
        return df_metrics

    def save_processed_data(self, df, filename="processed_data.csv"):
        """
        Save processed data to CSV.

        Args:
            df (pd.DataFrame): Processed data
            filename (str): Output filename
        """
        filepath = f"{self.config['data_paths']['processed']}{filename}"
        df.to_csv(filepath, index=False)
        logger.info(f"Saved processed data to {filepath}")

    def process_pipeline(self):
        """
        Run complete data processing pipeline.

        Returns:
            pd.DataFrame: Processed data
        """
        # Load data
        raw_data = self.load_raw_data()

        # Clean data
        clean_data = self.clean_data(raw_data)

        # Calculate metrics
        processed_data = self.calculate_metrics(clean_data)

        # Save processed data
        self.save_processed_data(processed_data)

        return processed_data


if __name__ == "__main__":
    processor = DataProcessor()
    processed_df = processor.process_pipeline()
    print(f"Processed data shape: {processed_df.shape}")
    print(processed_df.head())
