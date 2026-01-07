import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# Create more realistic data with trends and seasonality
np.random.seed(42)  # For reproducibility

# Generate 90 days of data
dates = pd.date_range(start='2024-01-01', periods=90, freq='D')

# Create realistic trends
base_sales = 200
trend = np.linspace(0, 50, 90)  # Increasing trend
seasonality = 50 * np.sin(np.linspace(0, 4*np.pi, 90))  # Weekly seasonality
noise = np.random.normal(0, 20, 90)

sales = base_sales + trend + seasonality + noise
sales = np.maximum(sales, 100)  # Ensure no negative sales

# Create correlated metrics
revenue = sales * np.random.uniform(8, 12, 90)  # $8-12 per sale
users = sales / np.random.uniform(2, 4, 90)  # Conversion rate affects users
conversion_rate = np.random.uniform(0.02, 0.05, 90) * (1 + 0.3 * np.sin(np.linspace(0, 2*np.pi, 90)))  # Seasonal conversion

# Create DataFrame
data = {
    'date': dates,
    'sales': sales.round().astype(int),
    'revenue': revenue.round(2),
    'users': users.round().astype(int),
    'conversion_rate': np.clip(conversion_rate, 0.01, 0.08)  # Keep between 1-8%
}

df = pd.DataFrame(data)

# Save to file
df.to_csv('data/raw/daily_metrics.csv', index=False)
print(f"✓ Improved sample data created with {len(df)} records")
print(f"Date range: {df['date'].min().date()} to {df['date'].max().date()}")
print(f"Sales range: {df['sales'].min()} to {df['sales'].max()}")
print(f"Revenue range: ${df['revenue'].min():.2f} to ${df['revenue'].max():.2f}")
