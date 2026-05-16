import xlwings as xw
import pandas as pd
from datetime import datetime

class ExcelExporter:
    def __init__(self, template_path=None):
        self.template_path = template_path
        self.app = None
        self.wb = None
        
    def create_report(self, df, analysis, summary_table, output_path):
        """Create formatted Excel report"""
        
        # Create a new workbook
        self.app = xw.App(visible=True)
        self.wb = self.app.books.add()
        
        # Sheet 1: Dashboard
        dashboard = self.wb.sheets[0]
        dashboard.name = "Dashboard"
        
        # Add title and timestamp
        dashboard.range("A1").value = "SALES PERFORMANCE DASHBOARD"
        dashboard.range("A1").font.bold = True
        dashboard.range("A1").font.size = 16
        dashboard.range("A2").value = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        
        # Key metrics
        metrics = [
            ["Total Revenue", f"${analysis['total_revenue']:,.2f}"],
            ["Average Order Value", f"${analysis['avg_order_value']:,.2f}"],
            ["Top Product", analysis['top_product']],
            ["Top Region", analysis['top_region']],
            ["Total Orders", analysis['total_orders']],
            ["Date Range", analysis['date_range']]
        ]
        
        dashboard.range("A4").value = "KEY METRICS"
        dashboard.range("A4").font.bold = True
        dashboard.range("A5").value = metrics
        
        # Sheet 2: Raw Data
        data_sheet = self.wb.sheets.add("Raw Data")
        data_sheet.range("A1").value = df.head(50)  # Show first 50 rows
        
        # Sheet 3: Summary Table
        summary_sheet = self.wb.sheets.add("Product-Region Summary")
        summary_sheet.range("A1").value = summary_table
        
        # Sheet 4: Monthly Trends
        monthly_sheet = self.wb.sheets.add("Monthly Trends")
        monthly_data = pd.DataFrame(list(analysis['monthly_sales'].items()), 
                                   columns=['Month', 'Sales'])
        monthly_sheet.range("A1").value = monthly_data
        
        # Auto-fit columns
        for sheet in self.wb.sheets:
            sheet.autofit()
        
        # Save the workbook
        self.wb.save(output_path)
        print(f"✅ Report saved to: {output_path}")
        
        return self.wb
    
    def close(self):
        """Close Excel gracefully"""
        if self.wb:
            self.wb.close()
        if self.app:
            self.app.quit()
