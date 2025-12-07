"""
Sample usage examples for Excel Analyzer
"""

from excel_analyzer import ExcelAnalyzer

# Example 1: Basic analysis
def example_basic_analysis(file_path):
    """Basic analysis example"""
    analyzer = ExcelAnalyzer(file_path)
    
    # Load data
    analyzer.load_data()
    
    # Get basic info
    info = analyzer.get_basic_info()
    print(f"Dataset has {info['total_rows']} rows and {info['total_columns']} columns")
    
    # Get statistical summary
    summary = analyzer.get_statistical_summary()
    
    # Generate reports
    analyzer.generate_html_report()
    analyzer.export_to_excel()


# Example 2: Detailed analysis with visualizations
def example_detailed_analysis(file_path):
    """Detailed analysis with all features"""
    analyzer = ExcelAnalyzer(file_path)
    
    # Load data from specific sheet
    analyzer.load_data(sheet_name='Sheet1')  # or use sheet index: 0
    
    # Perform all analyses
    analyzer.get_basic_info()
    analyzer.get_statistical_summary()
    analyzer.find_duplicates()
    analyzer.analyze_missing_data()
    analyzer.get_correlation_matrix()
    
    # Generate visualizations
    analyzer.generate_visualizations(output_dir='my_reports')
    
    # Generate custom reports
    analyzer.generate_html_report(output_file='my_reports/custom_report.html')
    analyzer.export_to_excel(output_file='my_reports/custom_analysis.xlsx')


# Example 3: Quick analysis
def quick_analysis(file_path):
    """Quick analysis for fast insights"""
    analyzer = ExcelAnalyzer(file_path)
    analyzer.load_data()
    
    # Just get the key metrics
    info = analyzer.get_basic_info()
    duplicates = analyzer.find_duplicates()
    missing = analyzer.analyze_missing_data()
    
    print(f"\n📊 Quick Summary:")
    print(f"Total Records: {info['total_rows']}")
    print(f"Duplicates: {duplicates['total_duplicates']}")
    print(f"Missing Values: {missing['total_missing']}")


if __name__ == "__main__":
    # Replace with your Excel file path
    excel_file = "your_data.xlsx"
    
    # Run the example you want
    # example_basic_analysis(excel_file)
    # example_detailed_analysis(excel_file)
    # quick_analysis(excel_file)
    
    print("Update the excel_file variable with your file path and uncomment the example you want to run.")
