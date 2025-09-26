#!/usr/bin/env python3
"""
Example usage of the AR Analysis Tool
This script demonstrates how to use the AR analyzer programmatically
"""

from ar_analysis import ARAnalyzer
import os

def example_basic_usage():
    """Basic usage example"""
    print("=== Basic AR Analysis Example ===")
    
    # Assuming you have an input file
    input_file = "your_invoice_data.xlsx"
    
    if not os.path.exists(input_file):
        print(f"Please place your invoice data file at: {input_file}")
        return
    
    # Create analyzer
    analyzer = ARAnalyzer(input_file, "monthly_ar_report.xlsx")
    
    # Run complete analysis
    success = analyzer.run_analysis()
    
    if success:
        print("\n📈 Key Results:")
        print(f"Collectible AR: ${analyzer.metrics['collectible_ar']:,.2f}")
        print(f"Collection Rate: {analyzer.metrics['collection_rate']:.1f}%")
        print(f"Total Excluded: ${analyzer.metrics['excluded_total']:,.2f}")

def example_custom_configuration():
    """Example with custom configuration"""
    print("\n=== Custom Configuration Example ===")
    
    # Create analyzer with custom settings
    analyzer = ARAnalyzer("input.xlsx", "custom_report.xlsx")
    
    # Customize exclusion thresholds
    analyzer.wire_fee_threshold = 50  # Lower threshold
    analyzer.india_withholding_docs = ['3148', '3149']  # Multiple docs
    
    # Custom exclusion notes
    analyzer.exclusion_notes['custom_exclusion'] = 'Custom exclusion reason'
    
    print("Custom configuration applied")

def example_batch_processing():
    """Example of processing multiple files"""
    print("\n=== Batch Processing Example ===")
    
    import glob
    from datetime import datetime
    
    # Find all Excel files in current directory
    excel_files = glob.glob("*invoices*.xlsx")
    
    if not excel_files:
        print("No invoice files found matching pattern '*invoices*.xlsx'")
        return
    
    for file in excel_files:
        print(f"Processing: {file}")
        
        # Generate unique output name
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"ar_report_{timestamp}_{os.path.basename(file)}"
        
        analyzer = ARAnalyzer(file, output_file)
        analyzer.run_analysis()

def example_access_detailed_data():
    """Example of accessing detailed analysis data"""
    print("\n=== Detailed Data Access Example ===")
    
    analyzer = ARAnalyzer("input.xlsx")
    
    if not analyzer.load_data():
        return
    
    analyzer.calculate_days_past_due()
    analyzer.categorize_invoices()
    analyzer.calculate_ar_metrics()
    
    # Access detailed data
    print("\n📊 Detailed Metrics:")
    for key, value in analyzer.metrics.items():
        if isinstance(value, (int, float)):
            if 'rate' in key or 'count' in key:
                print(f"{key}: {value:,.1f}")
            else:
                print(f"{key}: ${value:,.2f}")
    
    print("\n📋 Aging Breakdown:")
    for _, row in analyzer.aging_summary.iterrows():
        pct = (row['Total Amount'] / analyzer.metrics['collectible_ar']) * 100
        print(f"{row['Aging Category']}: ${row['Total Amount']:,.2f} ({pct:.1f}%)")

def example_monthly_automation():
    """Example of monthly automated reporting"""
    print("\n=== Monthly Automation Example ===")
    
    from datetime import datetime
    
    # Generate monthly report filename
    month_year = datetime.now().strftime("%B_%Y")
    output_file = f"AR_Report_{month_year}.xlsx"
    
    # Standard monthly process
    analyzer = ARAnalyzer("current_invoices.xlsx", output_file)
    
    if analyzer.run_analysis():
        print(f"✅ Monthly AR report generated: {output_file}")
        
        # Example: Send email notification (pseudo-code)
        # send_email_notification(
        #     subject=f"AR Report - {month_year}",
        #     body=f"Collectible AR: ${analyzer.metrics['collectible_ar']:,.2f}",
        #     attachment=output_file
        # )

if __name__ == "__main__":
    """Run all examples"""
    print("🚀 AR Analysis Tool Examples\n")
    
    # Run examples (comment out as needed)
    example_basic_usage()
    example_custom_configuration() 
    example_batch_processing()
    # example_access_detailed_data()  # Requires actual data file
    example_monthly_automation()
    
    print("\n✅ Examples complete!")
    print("\nNext steps:")
    print("1. Place your invoice data file in this directory")
    print("2. Run: python ar_analysis.py your_file.xlsx")
    print("3. Customize config.ini as needed")
    print("4. Set up automation for regular reports")
