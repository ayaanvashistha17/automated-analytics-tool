"""
Main entry point for Automated Analytics & Predictive Modeling Tool.
"""

import argparse
import logging
import sys
from datetime import datetime
import os

from data_processor import DataProcessor
from predictive_model import PredictiveModel
from report_generator import ReportGenerator
from excel_automation import ExcelAutomation

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/analytics_tool.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

def setup_directories():
    """Create necessary directories."""
    directories = [
        'logs',
        'data/raw',
        'data/processed',
        'data/forecasts',
        'reports/daily_reports',
        'reports/forecast_reports',
        'excel_files/macro_scripts',
        'config'
    ]
    
    for directory in directories:
        os.makedirs(directory, exist_ok=True)
        logger.debug(f"Created directory: {directory}")

def generate_daily_report():
    """Generate daily status report with forecasts."""
    try:
        logger.info("Starting daily report generation...")
        
        # Step 1: Process data
        processor = DataProcessor()
        processed_data = processor.process_pipeline()
        logger.info(f"Data processed: {len(processed_data)} records")
        
        # Step 2: Train model and generate forecasts
        model = PredictiveModel()
        metrics = model.train(processed_data, target_column='sales')
        logger.info(f"Model trained. R²: {metrics['r2_test']:.3f}")
        
        forecast = model.forecast(processed_data, periods=7, target_column='sales')
        model.save_forecast(forecast)
        logger.info(f"Forecast generated: {len(forecast)} days")
        
        # Step 3: Generate report
        generator = ReportGenerator()
        report_path = generator.create_daily_report(processed_data, forecast)
        logger.info(f"Report generated: {report_path}")
        
        # Step 4: Excel automation (if on Windows)
        automation = ExcelAutomation()
        automation.create_vba_macro_file()
        
        if os.name == 'nt':  # Windows
            logger.info("Running Excel automation...")
            # This would execute Excel automation in production
            pass
        
        logger.info("Daily report generation completed successfully!")
        return report_path
        
    except Exception as e:
        logger.error(f"Error generating daily report: {str(e)}")
        raise

def generate_forecast_only():
    """Generate only forecasts without full report."""
    try:
        logger.info("Generating forecasts...")
        
        processor = DataProcessor()
        processed_data = processor.process_pipeline()
        
        model = PredictiveModel()
        model.train(processed_data, target_column='sales')
        
        forecast = model.forecast(processed_data, periods=14, target_column='sales')
        
        # Create forecast-specific report
        generator = ReportGenerator()
        
        # Save forecast to dedicated file
        forecast_date = datetime.now().strftime('%Y-%m-%d')
        forecast_path = f"reports/forecast_reports/{forecast_date}_forecast.xlsx"
        
        # Create a simple forecast report
        import pandas as pd
        forecast.to_excel(forecast_path, index=False)
        
        logger.info(f"Forecast saved: {forecast_path}")
        return forecast_path
        
    except Exception as e:
        logger.error(f"Error generating forecast: {str(e)}")
        raise

def run_data_pipeline():
    """Run only data processing pipeline."""
    try:
        logger.info("Running data pipeline...")
        
        processor = DataProcessor()
        processed_data = processor.process_pipeline()
        
        logger.info(f"Data pipeline completed. Processed {len(processed_data)} records.")
        return processed_data
        
    except Exception as e:
        logger.error(f"Error in data pipeline: {str(e)}")
        raise

def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description='Automated Analytics & Predictive Modeling Tool'
    )
    parser.add_argument(
        '--report',
        choices=['daily', 'forecast', 'data', 'all'],
        default='all',
        help='Type of report to generate (default: all)'
    )
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='Enable verbose logging'
    )
    
    args = parser.parse_args()
    
    # Set log level
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # Setup directories
    setup_directories()
    
    logger.info("=" * 60)
    logger.info("Automated Analytics & Predictive Modeling Tool")
    logger.info(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"Mode: {args.report}")
    logger.info("=" * 60)
    
    try:
        if args.report == 'daily':
            report_path = generate_daily_report()
            print(f"\n✅ Daily report generated: {report_path}")
            
        elif args.report == 'forecast':
            forecast_path = generate_forecast_only()
            print(f"\n✅ Forecast generated: {forecast_path}")
            
        elif args.report == 'data':
            processed_data = run_data_pipeline()
            print(f"\n✅ Data processed: {len(processed_data)} records")
            
        elif args.report == 'all':
            report_path = generate_daily_report()
            print(f"\n✅ Complete automation finished:")
            print(f"   - Report: {report_path}")
            print(f"   - Forecasts saved in: data/forecasts/")
            print(f"   - Logs: logs/analytics_tool.log")
        
        logger.info("Process completed successfully!")
        
    except Exception as e:
        logger.error(f"Process failed: {str(e)}")
        print(f"\n❌ Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()