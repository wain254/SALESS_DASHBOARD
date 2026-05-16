import pandas as pd
import random
from datetime import datetime, timedelta

class SalesAnalyzer:
    def __init__(self):
        self.data = None
        
    def generate_sample_data(self, rows=100):
        """Generate sample sales data for demonstration"""
        products = ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'USB Cable']
        regions = ['North', 'South', 'East', 'West']
        
        data = {
            'Date': [datetime.today() - timedelta(days=x) for x in range(rows)],
            'Product': [random.choice(products) for _ in range(rows)],
            'Region': [random.choice(regions) for _ in range(rows)],
            'Quantity': [random.randint(1, 20) for _ in range(rows)],
            'Price': [round(random.uniform(10, 1000), 2) for _ in range(rows)]
        }
        
        df = pd.DataFrame(data)
        df['Total Sales'] = df['Quantity'] * df['Price']
        return df
    
    def analyze_data(self, df):
        """Perform key business analyses"""
        analysis = {
            'total_revenue': df['Total Sales'].sum(),
            'avg_order_value': df['Total Sales'].mean(),
            'top_product': df.groupby('Product')['Total Sales'].sum().idxmax(),
            'top_region': df.groupby('Region')['Total Sales'].sum().idxmax(),
            'total_orders': len(df),
            'date_range': f"{df['Date'].min().date()} to {df['Date'].max().date()}"
        }
        
        # Monthly summary
        df['Month'] = df['Date'].dt.strftime('%Y-%m')
        monthly_sales = df.groupby('Month')['Total Sales'].sum().to_dict()
        analysis['monthly_sales'] = monthly_sales
        
        # Product performance
        product_sales = df.groupby('Product')['Total Sales'].sum().to_dict()
        analysis['product_performance'] = product_sales
        
        return analysis
    
    def get_summary_table(self, df):
        """Create a pivot table for Excel export"""
        summary = df.groupby(['Product', 'Region']).agg({
            'Quantity': 'sum',
            'Total Sales': 'sum',
            'Price': 'mean'
        }).round(2)
        return summary
