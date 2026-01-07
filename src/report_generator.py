"""
Report generation module for creating daily status reports.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import logging
import os

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class ReportGenerator:
    """Generate daily status reports in Excel format."""
    
    def __init__(self):
        """Initialize ReportGenerator."""
        self.report_date = datetime.now().strftime('%Y-%m-%d')
        logger.info(f"ReportGenerator initialized for {self.report_date}")
    
    def create_daily_report(self, data_df, forecast_df=None):
        """
        Create daily status report Excel file.
        
        Args:
            data_df (pd.DataFrame): Daily metrics data
            forecast_df (pd.DataFrame): Forecast data (optional)
            
        Returns:
            str: Path to generated report
        """
        try:
            # Create workbook
            wb = openpyxl.Workbook()
            
            # Remove default sheet
            default_sheet = wb.active
            wb.remove(default_sheet)
            
            # Create summary sheet
            summary_ws = wb.create_sheet(title="Executive Summary")
            self._create_summary_sheet(summary_ws, data_df)
            
            # Create detailed metrics sheet
            metrics_ws = wb.create_sheet(title="Detailed Metrics")
            self._create_metrics_sheet(metrics_ws, data_df)
            
            # Create forecast sheet if forecast data provided
            if forecast_df is not None:
                forecast_ws = wb.create_sheet(title="Forecast")
                self._create_forecast_sheet(forecast_ws, forecast_df)
            
            # Create trends sheet
            trends_ws = wb.create_sheet(title="Trends Analysis")
            self._create_trends_sheet(trends_ws, data_df)
            
            # Create recommendations sheet
            rec_ws = wb.create_sheet(title="Recommendations")
            self._create_recommendations_sheet(rec_ws, data_df, forecast_df)
            
            # Save report
            report_dir = "reports/daily_reports"
            os.makedirs(report_dir, exist_ok=True)
            
            filename = f"{report_dir}/{self.report_date}_daily_report.xlsx"
            wb.save(filename)
            
            logger.info(f"Daily report created: {filename}")
            return filename
            
        except Exception as e:
            logger.error(f"Error creating report: {str(e)}")
            raise
    
    def _create_summary_sheet(self, ws, data_df):
        """Create executive summary sheet."""
        # Title
        ws['A1'] = "DAILY PERFORMANCE REPORT"
        ws['A1'].font = Font(size=16, bold=True, color="1F497D")
        ws.merge_cells('A1:F1')
        
        ws['A2'] = f"Date: {self.report_date}"
        ws['A2'].font = Font(bold=True)
        
        # Key Metrics Header
        ws['A4'] = "KEY PERFORMANCE INDICATORS"
        ws['A4'].font = Font(size=14, bold=True, color="1F497D")
        ws.merge_cells('A4:F4')
        
        # Get latest metrics
        latest_data = data_df.iloc[-1] if len(data_df) > 0 else pd.Series()
        
        # Define KPI rows
        kpis = [
            ("Total Sales", latest_data.get('sales', 0), "units"),
            ("Total Revenue", latest_data.get('revenue', 0), "$"),
            ("Active Users", latest_data.get('users', 0), "users"),
            ("Conversion Rate", f"{latest_data.get('conversion_rate', 0)*100:.2f}", "%"),
            ("Sales Growth", f"{latest_data.get('sales_growth', 0):.2f}", "%"),
            ("Revenue Growth", f"{latest_data.get('revenue_growth', 0):.2f}", "%")
        ]
        
        # Populate KPIs
        row = 6
        for kpi_name, value, unit in kpis:
            ws[f'A{row}'] = kpi_name
            ws[f'B{row}'] = value
            ws[f'C{row}'] = unit
            
            # Format based on value
            cell = ws[f'B{row}']
            if isinstance(value, (int, float)):
                if 'growth' in kpi_name.lower():
                    if value > 0:
                        cell.font = Font(color="00B050", bold=True)  # Green
                    elif value < 0:
                        cell.font = Font(color="FF0000", bold=True)  # Red
            
            row += 1
        
        # Performance Summary
        row += 2
        ws[f'A{row}'] = "PERFORMANCE SUMMARY"
        ws[f'A{row}'].font = Font(size=14, bold=True, color="1F497D")
        ws.merge_cells(f'A{row}:F{row}')
        
        row += 2
        summary_text = self._generate_summary_text(data_df)
        ws[f'A{row}'] = summary_text
        ws.merge_cells(f'A{row}:F{row+5}')
        ws[f'A{row}'].alignment = Alignment(wrap_text=True, vertical='top')
        
        # Apply formatting
        self._apply_sheet_formatting(ws)
    
    def _generate_summary_text(self, data_df):
        """Generate performance summary text."""
        if len(data_df) < 2:
            return "Insufficient data for summary analysis."
        
        latest = data_df.iloc[-1]
        previous = data_df.iloc[-2]
        
        summary = []
        summary.append(f"Performance Summary for {self.report_date}:")
        summary.append("")
        
        # Sales analysis
        sales_growth = latest.get('sales_growth', 0)
        if sales_growth > 5:
            summary.append(f"✓ Sales showed strong growth of {sales_growth:.1f}% compared to yesterday.")
        elif sales_growth > 0:
            summary.append(f"✓ Sales increased by {sales_growth:.1f}% compared to yesterday.")
        elif sales_growth < 0:
            summary.append(f"⚠ Sales decreased by {abs(sales_growth):.1f}% compared to yesterday.")
        else:
            summary.append(f"○ Sales remained stable compared to yesterday.")
        
        # Revenue analysis
        revenue_growth = latest.get('revenue_growth', 0)
        if revenue_growth > sales_growth:
            summary.append(f"✓ Revenue growth ({revenue_growth:.1f}%) outpaced sales growth, indicating higher average order value.")
        
        # User analysis
        users_change = latest.get('users', 0) - previous.get('users', 0)
        if users_change > 10:
            summary.append(f"✓ User base increased by {users_change} active users.")
        
        # Conversion rate analysis
        conv_rate = latest.get('conversion_rate', 0) * 100
        if conv_rate > 3:
            summary.append(f"✓ Conversion rate remains strong at {conv_rate:.2f}%.")
        elif conv_rate < 1:
            summary.append(f"⚠ Conversion rate is low at {conv_rate:.2f}%, needs attention.")
        
        return "\n".join(summary)
    
    def _create_metrics_sheet(self, ws, data_df):
        """Create detailed metrics sheet."""
        # Title
        ws['A1'] = "DETAILED DAILY METRICS"
        ws['A1'].font = Font(size=16, bold=True, color="1F497D")
        ws.merge_cells('A1:H1')
        
        # Write data
        if not data_df.empty:
            # Write headers
            headers = list(data_df.columns)
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=3, column=col_idx, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            
            # Write data rows
            for row_idx, row_data in enumerate(data_df.itertuples(index=False), 4):
                for col_idx, value in enumerate(row_data, 1):
                    ws.cell(row=row_idx, column=col_idx, value=value)
            
            # Apply number formatting
            for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=2, max_col=len(headers)):
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        if 'rate' in str(cell.column):
                            cell.number_format = '0.00%'
                        elif 'growth' in str(cell.column):
                            cell.number_format = '0.00%'
                        elif 'revenue' in str(cell.column):
                            cell.number_format = '$#,##0.00'
                        else:
                            cell.number_format = '#,##0'
        
        self._apply_sheet_formatting(ws)
    
    def _create_forecast_sheet(self, ws, forecast_df):
        """Create forecast analysis sheet."""
        # Title
        ws['A1'] = "PREDICTIVE FORECAST"
        ws['A1'].font = Font(size=16, bold=True, color="1F497D")
        ws.merge_cells('A1:F1')
        
        ws['A2'] = "AI-Powered 7-Day Forecast"
        ws['A2'].font = Font(size=12, bold=True, color="4472C4")
        
        # Write forecast data
        if not forecast_df.empty:
            # Headers
            headers = ['Date', 'Predicted Sales', 'Lower Bound', 'Upper Bound', 'Confidence']
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=4, column=col_idx, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
            
            # Data rows
            for row_idx, row in forecast_df.iterrows():
                ws.cell(row=row_idx+5, column=1, value=row['date'].strftime('%Y-%m-%d'))
                ws.cell(row=row_idx+5, column=2, value=row['predicted_sales'])
                ws.cell(row=row_idx+5, column=3, value=row['confidence_interval_lower'])
                ws.cell(row=row_idx+5, column=4, value=row['confidence_interval_upper'])
                
                # Calculate confidence range
                range_pct = ((row['confidence_interval_upper'] - row['confidence_interval_lower']) / 
                            row['predicted_sales'] * 100)
                ws.cell(row=row_idx+5, column=5, value=f"{range_pct:.1f}%")
                
                # Color code based on confidence
                if range_pct < 10:
                    ws.cell(row=row_idx+5, column=5).font = Font(color="00B050", bold=True)
                elif range_pct < 20:
                    ws.cell(row=row_idx+5, column=5).font = Font(color="FFC000", bold=True)
                else:
                    ws.cell(row=row_idx+5, column=5).font = Font(color="FF0000", bold=True)
            
            # Forecast insights
            last_row = len(forecast_df) + 6
            ws.cell(row=last_row, column=1, value="FORECAST INSIGHTS:").font = Font(bold=True)
            
            insights = self._generate_forecast_insights(forecast_df)
            for i, insight in enumerate(insights, 1):
                ws.cell(row=last_row + i, column=1, value=f"• {insight}")
        
        self._apply_sheet_formatting(ws)
    
    def _generate_forecast_insights(self, forecast_df):
        """Generate insights from forecast data."""
        insights = []
        
        if len(forecast_df) > 0:
            avg_sales = forecast_df['predicted_sales'].mean()
            min_sales = forecast_df['predicted_sales'].min()
            max_sales = forecast_df['predicted_sales'].max()
            
            insights.append(f"Average predicted sales: {avg_sales:.0f} units per day")
            insights.append(f"Expected range: {min_sales:.0f} to {max_sales:.0f} units")
            
            # Trend analysis
            if len(forecast_df) > 1:
                first = forecast_df.iloc[0]['predicted_sales']
                last = forecast_df.iloc[-1]['predicted_sales']
                trend = ((last - first) / first) * 100
                
                if trend > 5:
                    insights.append(f"Positive trend: {trend:.1f}% increase over forecast period")
                elif trend < -5:
                    insights.append(f"Negative trend: {abs(trend):.1f}% decrease over forecast period")
                else:
                    insights.append("Stable trend expected over forecast period")
            
            insights.append("Recommendation: Maintain current inventory levels to meet forecasted demand")
        
        return insights
    
    def _create_trends_sheet(self, ws, data_df):
        """Create trends analysis sheet."""
        # Title
        ws['A1'] = "TRENDS ANALYSIS"
        ws['A1'].font = Font(size=16, bold=True, color="1F497D")
        ws.merge_cells('A1:F1')
        
        # Calculate trends
        if len(data_df) >= 7:
            # Weekly trends
            recent_week = data_df.iloc[-7:]
            previous_week = data_df.iloc[-14:-7] if len(data_df) >= 14 else data_df.iloc[-7:]
            
            trends_data = []
            
            for metric in ['sales', 'revenue', 'users']:
                if metric in recent_week.columns:
                    current_avg = recent_week[metric].mean()
                    previous_avg = previous_week[metric].mean() if len(previous_week) > 0 else current_avg
                    
                    if previous_avg > 0:
                        change_pct = ((current_avg - previous_avg) / previous_avg) * 100
                    else:
                        change_pct = 0
                    
                    trends_data.append({
                        'Metric': metric.title(),
                        'Current Week Avg': current_avg,
                        'Previous Week Avg': previous_avg,
                        'Change %': change_pct
                    })
            
            # Write trends table
            if trends_data:
                headers = ['Metric', 'Current Week Avg', 'Previous Week Avg', 'Change %', 'Trend']
                for col_idx, header in enumerate(headers, 1):
                    cell = ws.cell(row=3, column=col_idx, value=header)
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                
                for row_idx, trend in enumerate(trends_data, 4):
                    ws.cell(row=row_idx, column=1, value=trend['Metric'])
                    ws.cell(row=row_idx, column=2, value=trend['Current Week Avg'])
                    ws.cell(row=row_idx, column=3, value=trend['Previous Week Avg'])
                    ws.cell(row=row_idx, column=4, value=trend['Change %'] / 100)
                    
                    # Add trend indicator
                    change = trend['Change %']
                    if change > 5:
                        trend_indicator = "↗ Strong Growth"
                        color = "00B050"
                    elif change > 0:
                        trend_indicator = "↗ Moderate Growth"
                        color = "92D050"
                    elif change < -5:
                        trend_indicator = "↘ Significant Decline"
                        color = "FF0000"
                    elif change < 0:
                        trend_indicator = "↘ Moderate Decline"
                        color = "FFC000"
                    else:
                        trend_indicator = "→ Stable"
                        color = "000000"
                    
                    ws.cell(row=row_idx, column=5, value=trend_indicator).font = Font(color=color, bold=True)
                    
                    # Format percentages
                    ws.cell(row=row_idx, column=4).number_format = '0.00%'
        
        self._apply_sheet_formatting(ws)
    
    def _create_recommendations_sheet(self, ws, data_df, forecast_df):
        """Create actionable recommendations sheet."""
        # Title
        ws['A1'] = "ACTIONABLE RECOMMENDATIONS"
        ws['A1'].font = Font(size=16, bold=True, color="1F497D")
        ws.merge_cells('A1:F1')
        
        ws['A2'] = "AI-Generated Insights for Proactive Management"
        ws['A2'].font = Font(size=12, italic=True, color="4472C4")
        
        # Generate recommendations
        recommendations = self._generate_recommendations(data_df, forecast_df)
        
        row = 4
        for i, rec in enumerate(recommendations, 1):
            # Recommendation title
            title_cell = ws.cell(row=row, column=1, value=f"Recommendation #{i}: {rec['title']}")
            title_cell.font = Font(bold=True, color="1F497D")
            ws.merge_cells(f'A{row}:F{row}')
            row += 1
            
            # Description
            desc_cell = ws.cell(row=row, column=1, value=rec['description'])
            desc_cell.alignment = Alignment(wrap_text=True)
            ws.merge_cells(f'A{row}:F{row}')
            row += 1
            
            # Impact
            impact_cell = ws.cell(row=row, column=1, value=f"Expected Impact: {rec['impact']}")
            impact_cell.font = Font(bold=True)
            ws.merge_cells(f'A{row}:F{row}')
            row += 1
            
            # Priority
            priority_cell = ws.cell(row=row, column=1, value=f"Priority: {rec['priority']}")
            priority_cell.font = Font(bold=True, color=rec['priority_color'])
            ws.merge_cells(f'A{row}:F{row}')
            row += 2
        
        # Add footer note
        ws.cell(row=row, column=1, 
                value="Note: These recommendations are generated using predictive analytics and should be reviewed by domain experts.")
        ws.merge_cells(f'A{row}:F{row}')
        
        self._apply_sheet_formatting(ws)
    
    def _generate_recommendations(self, data_df, forecast_df):
        """Generate AI-powered recommendations."""
        recommendations = []
        
        if len(data_df) >= 7:
            latest = data_df.iloc[-1]
            week_avg = data_df.iloc[-7:].mean()
            
            # Recommendation 1: Based on conversion rate
            conv_rate = latest.get('conversion_rate', 0) * 100
            if conv_rate < 2:
                recommendations.append({
                    'title': 'Optimize Conversion Funnel',
                    'description': 'Conversion rate is below optimal levels. Conduct A/B testing on checkout process and improve CTAs.',
                    'impact': 'Potential 15-25% increase in conversion rate',
                    'priority': 'High',
                    'priority_color': 'FF0000'
                })
            
            # Recommendation 2: Based on sales trend
            sales_growth = latest.get('sales_growth', 0)
            if sales_growth < -5:
                recommendations.append({
                    'title': 'Boost Sales Initiatives',
                    'description': 'Sales showing negative trend. Implement promotional campaigns and review pricing strategy.',
                    'impact': 'Potential to reverse negative trend within 7-10 days',
                    'priority': 'High',
                    'priority_color': 'FF0000'
                })
            
            # Recommendation 3: Based on forecast
            if forecast_df is not None and not forecast_df.empty:
                forecast_avg = forecast_df['predicted_sales'].mean()
                current_avg = week_avg.get('sales', 0)
                
                if forecast_avg > current_avg * 1.1:
                    recommendations.append({
                        'title': 'Increase Inventory Preparation',
                        'description': f'Forecast predicts {((forecast_avg/current_avg)-1)*100:.0f}% increase in demand. Prepare inventory accordingly.',
                        'impact': 'Prevent stockouts and capitalize on increased demand',
                        'priority': 'Medium',
                        'priority_color': 'FFC000'
                    })
            
            # Recommendation 4: Based on user metrics
            users_growth = latest.get('users', 0) - data_df.iloc[-2].get('users', 0) if len(data_df) >= 2 else 0
            if users_growth < 0:
                recommendations.append({
                    'title': 'Enhance User Engagement',
                    'description': 'Active users decreased. Launch re-engagement campaign and improve onboarding experience.',
                    'impact': 'Potential to regain lost users and improve retention',
                    'priority': 'Medium',
                    'priority_color': 'FFC000'
                })
        
        # Default recommendation if none generated
        if not recommendations:
            recommendations.append({
                'title': 'Continue Monitoring Key Metrics',
                'description': 'Current performance metrics are within expected ranges. Maintain regular monitoring and focus on incremental improvements.',
                'impact': 'Sustained performance stability',
                'priority': 'Low',
                'priority_color': '00B050'
            })
        
        return recommendations
    
    def _apply_sheet_formatting(self, ws):
        """Apply consistent formatting to worksheet."""
        # Auto-adjust column widths
        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # Apply borders to data tables
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.border = thin_border
                if cell.row % 2 == 0 and cell.row > 3:
                    cell.fill = PatternFill(start_color="F8F8F8", end_color="F8F8F8", fill_type="solid")

if __name__ == "__main__":
    # Example usage
    from data_processor import DataProcessor
    from predictive_model import PredictiveModel
    
    # Generate sample data and report
    processor = DataProcessor()
    processed_data = processor.process_pipeline()
    
    model = PredictiveModel()
    model.train(processed_data)
    forecast = model.forecast(processed_data, periods=7)
    
    generator = ReportGenerator()
    report_path = generator.create_daily_report(processed_data, forecast)
    
    print(f"Report generated: {report_path}")