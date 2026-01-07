#!/bin/bash

# Automated Analytics Tool - Setup Script
# For MacBook/Linux systems

echo "=========================================="
echo "Automated Analytics Tool - Setup"
echo "=========================================="

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Function to print colored output
print_status() {
    echo -e "${GREEN}[✓]${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}[!]${NC} $1"
}

print_error() {
    echo -e "${RED}[✗]${NC} $1"
}

# Check Python version
echo "Checking Python version..."
python_version=$(python3 --version 2>&1 | awk '{print $2}')
if [[ $python_version =~ ^3\.(8|9|10|11|12)\. ]]; then
    print_status "Python $python_version detected"
else
    print_error "Python 3.8+ is required. Found: $python_version"
    exit 1
fi

# Create virtual environment
echo "Creating virtual environment..."
python3 -m venv venv
if [ $? -eq 0 ]; then
    print_status "Virtual environment created"
else
    print_error "Failed to create virtual environment"
    exit 1
fi

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate
if [ $? -eq 0 ]; then
    print_status "Virtual environment activated"
else
    print_error "Failed to activate virtual environment"
    exit 1
fi

# Upgrade pip
echo "Upgrading pip..."
pip install --upgrade pip
print_status "pip upgraded"

# Install requirements
echo "Installing requirements..."
pip install -r requirements.txt
if [ $? -eq 0 ]; then
    print_status "Requirements installed"
else
    print_error "Failed to install requirements"
    exit 1
fi

# Create necessary directories
echo "Creating project directories..."
mkdir -p data/raw data/processed data/forecasts
mkdir -p reports/daily_reports reports/forecast_reports
mkdir -p excel_files/macro_scripts
mkdir -p logs config tests
mkdir -p .github/workflows

print_status "Directories created"

# Generate sample data
echo "Generating sample data..."
python3 -c "
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# Generate sample daily metrics
dates = pd.date_range(start='2024-01-01', end='2024-01-30', freq='D')
data = {
    'date': dates,
    'sales': np.random.randint(100, 500, len(dates)),
    'revenue': np.random.uniform(1000, 5000, len(dates)),
    'users': np.random.randint(50, 200, len(dates)),
    'conversion_rate': np.random.uniform(0.01, 0.05, len(dates))
}

df = pd.DataFrame(data)
df.to_csv('data/raw/daily_metrics.csv', index=False)
print('Sample data generated: data/raw/daily_metrics.csv')

# Generate historical data
historical_dates = pd.date_range(start='2023-01-01', end='2023-12-31', freq='D')
historical_data = {
    'date': historical_dates,
    'sales': np.random.randint(50, 300, len(historical_dates)),
    'revenue': np.random.uniform(500, 3000, len(historical_dates))
}

historical_df = pd.DataFrame(historical_data)
historical_df.to_csv('data/raw/historical_data.csv', index=False)
print('Historical data generated: data/raw/historical_data.csv')
"

print_status "Sample data generated"

# Create default config if not exists
if [ ! -f "config/config.yaml" ]; then
    echo "Creating default configuration..."
    cat > config/config.yaml << 'EOF'
# Configuration for Automated Analytics Tool

paths:
  raw_data: "data/raw/"
  processed_data: "data/processed/"
  forecasts: "data/forecasts/"
  reports: "reports/"
  logs: "logs/"

data:
  source_files:
    daily_metrics: "daily_metrics.csv"
    historical_data: "historical_data.csv"
  
  target_column: "sales"

model:
  test_size: 0.2
  forecast_periods: 7

reporting:
  company_name: "Demo Company"
  report_title: "Daily Analytics Report"

logging:
  level: "INFO"
  file: "logs/analytics_tool.log"
EOF
    print_status "Default configuration created"
fi

# Set up git (if not already)
if [ ! -d ".git" ]; then
    echo "Initializing git repository..."
    git init
    print_status "Git repository initialized"
fi

# Make scripts executable
chmod +x scripts/*.sh

echo ""
echo "=========================================="
echo "Setup Complete!"
echo "=========================================="
echo ""
echo "Next steps:"
echo "1. Activate virtual environment:"
echo "   source venv/bin/activate"
echo ""
echo "2. Run a test report:"
echo "   python src/main.py --report daily"
echo ""
echo "3. View generated report:"
echo "   open reports/daily_reports/*.xlsx"
echo ""
echo "4. Run tests:"
echo "   python -m pytest tests/"
echo ""
echo "For more information, see README.md"
echo "=========================================="