from crypto_tracker import CryptoTracker
from datetime import datetime
import pandas as pd
from pathlib import Path

def generate_report(tracker, output_file='Crypto_Analysis_Report.pdf'):
    """Generate a PDF report with key insights"""
    try:
        # Read the latest data
        excel_data = pd.read_excel(tracker.excel_file)
        
        # Create report content using f-strings
        report_content = f"""
Cryptocurrency Market Analysis Report
Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

Executive Summary
----------------
This report provides an analysis of the top 50 cryptocurrencies by market capitalization.

Key Metrics
-----------
1. Market Overview
   - Total Number of Cryptocurrencies Analyzed: {len(excel_data)}
   - Total Market Capitalization: ${excel_data['Market Cap (USD)'].sum():,.2f}
   - Average Price: ${excel_data['Price (USD)'].mean():,.2f}
   - Median Price: ${excel_data['Price (USD)'].median():,.2f}

2. Top 5 Cryptocurrencies by Market Cap
{excel_data.head(5)[['Name', 'Symbol', 'Market Cap (USD)', 'Price (USD)']].to_string()}

3. Price Changes (24h)
   - Highest Gain: {excel_data.nlargest(1, '24h Change (%)')['Name'].iloc[0]} ({excel_data['24h Change (%)'].max():.2f}%)
   - Biggest Drop: {excel_data.nsmallest(1, '24h Change (%)')['Name'].iloc[0]} ({excel_data['24h Change (%)'].min():.2f}%)

4. Trading Volume
   - Total 24h Volume: ${excel_data['24h Volume'].sum():,.2f}
   - Average Volume per Cryptocurrency: ${excel_data['24h Volume'].mean():,.2f}

Market Insights
--------------
1. Market Concentration
   - Top 5 cryptocurrencies represent {(excel_data.head(5)['Market Cap (USD)'].sum() / excel_data['Market Cap (USD)'].sum() * 100):.2f}% of total market cap
   
2. Volatility Analysis
   - Number of cryptocurrencies with >5% price change: {len(excel_data[excel_data['24h Change (%)'].abs() > 5])}
   - Number of cryptocurrencies with >10% price change: {len(excel_data[excel_data['24h Change (%)'].abs() > 10])}

Data Source: CoinGecko API
Update Frequency: Every 5 minutes
"""
        
        # Save report to file
        with open('report.txt', 'w') as f:
            f.write(report_content)
            
        print("Report generated successfully!")
        return report_content
        
    except Exception as e:
        print(f"Error generating report: {e}")
        return None

def main():
    # Create output directory if it doesn't exist
    Path("output").mkdir(exist_ok=True)
    
    # Initialize the tracker
    tracker = CryptoTracker(
        excel_file='output/crypto_live_data.xlsx'
    )
    
    # Generate initial report
    print("Generating initial analysis report...")
    generate_report(tracker, 'output/Crypto_Analysis_Report.pdf')
    
    # Start the live tracking
    print("\nStarting live tracking...")
    tracker.run(update_interval=300)  # Update every 5 minutes

if __name__ == "__main__":
    main()