"""
Predictive modeling module using linear regression for forecasting.
"""

import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error, r2_score
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime, timedelta
import logging
import json

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class PredictiveModel:
    """Linear regression model for trend analysis and forecasting."""
    
    def __init__(self):
        """Initialize PredictiveModel."""
        self.model = LinearRegression()
        self.is_trained = False
        self.model_metrics = {}
        logger.info("PredictiveModel initialized")
    
    def prepare_features(self, df, target_column='sales'):
        """
        Prepare features for training.
        
        Args:
            df (pd.DataFrame): Input data
            target_column (str): Column to predict
            
        Returns:
            tuple: Features and target arrays
        """
        df_features = df.copy()
        
        # Ensure we have a date column
        if 'date' not in df_features.columns:
            df_features['date'] = pd.date_range(start='2024-01-01', periods=len(df_features))
        
        # Create time-based features
        df_features['day_of_week'] = df_features['date'].dt.dayofweek
        df_features['day_of_month'] = df_features['date'].dt.day
        df_features['month'] = df_features['date'].dt.month
        df_features['quarter'] = df_features['date'].dt.quarter
        
        # Create lag features
        for lag in [1, 7, 14]:
            df_features[f'lag_{lag}'] = df_features[target_column].shift(lag)
        
        # Create rolling statistics
        df_features['rolling_mean_7'] = df_features[target_column].rolling(window=7).mean()
        df_features['rolling_std_7'] = df_features[target_column].rolling(window=7).std()
        
        # Drop rows with NaN values
        df_features = df_features.dropna()
        
        # Select features
        feature_columns = [
            'day_of_week', 'day_of_month', 'month', 'quarter',
            'lag_1', 'lag_7', 'lag_14',
            'rolling_mean_7', 'rolling_std_7'
        ]
        
        # Filter available features
        available_features = [col for col in feature_columns if col in df_features.columns]
        
        X = df_features[available_features]
        y = df_features[target_column]
        
        logger.info(f"Prepared features: {X.shape}")
        return X, y, available_features
    
    def train(self, df, target_column='sales', test_size=0.2):
        """
        Train the linear regression model.
        
        Args:
            df (pd.DataFrame): Training data
            target_column (str): Column to predict
            test_size (float): Proportion of data for testing
            
        Returns:
            dict: Model metrics
        """
        try:
            # Prepare features
            X, y, feature_names = self.prepare_features(df, target_column)
            
            # Split data
            X_train, X_test, y_train, y_test = train_test_split(
                X, y, test_size=test_size, random_state=42, shuffle=False
            )
            
            # Train model
            self.model.fit(X_train, y_train)
            self.is_trained = True
            self.feature_names = feature_names
            
            # Make predictions
            y_pred_train = self.model.predict(X_train)
            y_pred_test = self.model.predict(X_test)
            
            # Calculate metrics
            self.model_metrics = {
                'r2_train': r2_score(y_train, y_pred_train),
                'r2_test': r2_score(y_test, y_pred_test),
                'mse_train': mean_squared_error(y_train, y_pred_train),
                'mse_test': mean_squared_error(y_test, y_pred_test),
                'rmse_train': np.sqrt(mean_squared_error(y_train, y_pred_train)),
                'rmse_test': np.sqrt(mean_squared_error(y_test, y_pred_test)),
                'feature_importance': dict(zip(feature_names, self.model.coef_)),
                'intercept': float(self.model.intercept_),
                'num_features': len(feature_names),
                'train_samples': len(X_train),
                'test_samples': len(X_test)
            }
            
            logger.info(f"Model trained. R² Test: {self.model_metrics['r2_test']:.3f}")
            return self.model_metrics
            
        except Exception as e:
            logger.error(f"Error training model: {str(e)}")
            raise
    
    def forecast(self, df, periods=7, target_column='sales'):
        """
        Generate future forecasts.
        
        Args:
            df (pd.DataFrame): Historical data
            periods (int): Number of periods to forecast
            target_column (str): Column to forecast
            
        Returns:
            pd.DataFrame: Forecast results
        """
        if not self.is_trained:
            raise ValueError("Model must be trained before forecasting")
        
        # Prepare the last available data point
        X_last, _, _ = self.prepare_features(df, target_column)
        if len(X_last) == 0:
            raise ValueError("No data available for forecasting")
        
        # Start from the last date
        last_date = df['date'].max() if 'date' in df.columns else datetime.now()
        
        # Generate future dates
        future_dates = pd.date_range(
            start=last_date + timedelta(days=1),
            periods=periods,
            freq='D'
        )
        
        forecasts = []
        
        # Make iterative predictions
        current_features = X_last.iloc[-1:].copy()
        
        for i in range(periods):
            # Update date features for the forecast period
            forecast_date = future_dates[i]
            current_features['day_of_week'] = forecast_date.dayofweek
            current_features['day_of_month'] = forecast_date.day
            current_features['month'] = forecast_date.month
            current_features['quarter'] = (forecast_date.month - 1) // 3 + 1
            
            # Make prediction
            prediction = self.model.predict(current_features)[0]
            
            forecasts.append({
                'date': forecast_date,
                f'predicted_{target_column}': prediction,
                'forecast_period': i + 1,
                'confidence_interval_lower': prediction * 0.9,  # Simplified
                'confidence_interval_upper': prediction * 1.1   # Simplified
            })
            
            # Update lag features for next prediction (simplified)
            if i == 0:
                current_features['lag_1'] = df[target_column].iloc[-1]
            else:
                current_features['lag_1'] = forecasts[-2][f'predicted_{target_column}']
        
        forecast_df = pd.DataFrame(forecasts)
        
        logger.info(f"Generated {periods}-day forecast")
        return forecast_df
    
    def save_forecast(self, forecast_df, filename='forecast_results.csv'):
        """
        Save forecast results to CSV.
        
        Args:
            forecast_df (pd.DataFrame): Forecast data
            filename (str): Output filename
        """
        filepath = f"data/forecasts/{filename}"
        forecast_df.to_csv(filepath, index=False)
        logger.info(f"Saved forecast to {filepath}")
    
    def plot_forecast(self, historical_df, forecast_df, target_column='sales'):
        """
        Create visualization of forecast vs historical.
        
        Args:
            historical_df (pd.DataFrame): Historical data
            forecast_df (pd.DataFrame): Forecast data
            target_column (str): Target column name
        """
        plt.figure(figsize=(12, 6))
        
        # Plot historical data
        plt.plot(
            historical_df['date'],
            historical_df[target_column],
            'b-',
            label='Historical',
            alpha=0.7
        )
        
        # Plot forecast
        plt.plot(
            forecast_df['date'],
            forecast_df[f'predicted_{target_column}'],
            'r--',
            label='Forecast',
            linewidth=2
        )
        
        # Plot confidence interval
        plt.fill_between(
            forecast_df['date'],
            forecast_df['confidence_interval_lower'],
            forecast_df['confidence_interval_upper'],
            alpha=0.2,
            color='red',
            label='90% Confidence Interval'
        )
        
        plt.title(f'{target_column.title()} Forecast', fontsize=16)
        plt.xlabel('Date', fontsize=12)
        plt.ylabel(target_column.title(), fontsize=12)
        plt.legend()
        plt.grid(True, alpha=0.3)
        plt.tight_layout()
        
        # Save plot
        plot_filename = f"reports/forecast_plot_{datetime.now().strftime('%Y%m%d')}.png"
        plt.savefig(plot_filename, dpi=300)
        logger.info(f"Saved forecast plot to {plot_filename}")
        
        return plt.gcf()

if __name__ == "__main__":
    # Example usage
    from data_processor import DataProcessor
    
    # Process data
    processor = DataProcessor()
    processed_data = processor.process_pipeline()
    
    # Train model and forecast
    model = PredictiveModel()
    metrics = model.train(processed_data, target_column='sales')
    
    print("Model Metrics:")
    print(json.dumps(metrics, indent=2))
    
    # Generate forecast
    forecast = model.forecast(processed_data, periods=14, target_column='sales')
    model.save_forecast(forecast)
    
    print("\nForecast Results:")
    print(forecast[['date', 'predicted_sales', 'confidence_interval_lower', 'confidence_interval_upper']])