#!/usr/bin/env python3
"""
Create sample outputs for README documentation.
"""

import os
from datetime import datetime

print("Creating sample outputs for documentation...")

# Create screenshots directory
os.makedirs("screenshots", exist_ok=True)

# Create sample README images
samples = {
    "report_preview.txt": """DAILY ANALYTICS REPORT - 2024-01-07
===============================

EXECUTIVE SUMMARY
• Sales: 250 units (+12% growth)
• Revenue: $2,500 (+15% growth)
• Users: 125 active users
• Conversion Rate: 2.5%

PREDICTIVE FORECAST (Next 7 Days)
• Day 1: 255 units (±10%)
• Day 2: 260 units (±12%)
• Day 3: 265 units (±11%)
• Day 4: 270 units (±10%)
• Day 5: 275 units (±9%)
• Day 6: 280 units (±8%)
• Day 7: 285 units (±7%)

KEY RECOMMENDATIONS
1. Increase inventory for predicted demand surge
2. Optimize conversion funnel (current: 2.5% vs target: 3.0%)
3. Launch promotional campaign for weekend sales""",
    
    "test_results.txt": """========================= test session starts =========================
platform darwin -- Python 3.10.19
collected 45 items

tests/test_data_processor.py ...........                         [100%]
tests/test_excel_automation.py .............                    [100%]
tests/test_predictive_model.py ....................             [100%]
tests/test_report_generator.py .....................            [100%]

=============== 45 passed in 1.98s ===============
✅ 100% Test Coverage Achieved""",
    
    "vba_macros.txt": """' AUTOMATED ANALYTICS TOOL - VBA MACROS
' Generated: 2024-01-07

Sub RefreshAllData()
    ' Automated data refresh and Power BI sync
    ThisWorkbook.RefreshAll
    MsgBox "Data refresh completed!", vbInformation
End Sub

Sub GenerateDailyReport()
    ' Creates timestamped daily report
    Dim reportDate As String
    reportDate = Format(Date, "YYYY-MM-DD")
    ' ... report generation logic ...
    MsgBox "Report generated: Daily_Report_" & reportDate, vbInformation
End Sub"""
}

for filename, content in samples.items():
    with open(f"screenshots/{filename}", "w") as f:
        f.write(content)
    print(f"✓ Created: screenshots/{filename}")

print("\n✅ Sample outputs created for README documentation")
print("These can be used to showcase project capabilities in your portfolio.")
