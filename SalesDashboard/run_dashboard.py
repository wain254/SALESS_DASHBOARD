import os
import sys
from datetime import datetime

# Add src folder to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from sales_analyzer import SalesAnalyzer
from excel_exporter import ExcelExporter

def main():
    print("🚀 Starting Sales Dashboard Generator...")
    
    # 1. Initialize analyzer
    analyzer = SalesAnalyzer()
    
    # 2. Generate or load data
    print("📊 Generating sample sales data...")
    df = analyzer.generate_sample_data(rows=200)
    
    # 3. Perform analysis
    print("🔍 Analyzing data...")
    analysis = analyzer.analyze_data(df)
    
    # 4. Create summary table
    summary_table = analyzer.get_summary_table(df)
    
    # 5. Export to Excel
    print("📝 Creating Excel report...")
    exporter = ExcelExporter()
    
    output_dir = os.path.join(os.path.dirname(__file__), 'output')
    os.makedirs(output_dir, exist_ok=True)
    
    output_file = os.path.join(output_dir, 
                               f"Sales_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    
    # Create the report
    wb = exporter.create_report(df, analysis, summary_table, output_file)
    
    # Keep Excel open so you can see the result
    print("\n✨ Dashboard created successfully!")
    print(f"📁 File location: {output_file}")
    print("\n💡 Excel is open with your report. Close Excel when done.")
    
    input("\nPress Enter to close Excel and exit...")
    exporter.close()
    print("👋 Done!")

if __name__ == "__main__":
    main()
