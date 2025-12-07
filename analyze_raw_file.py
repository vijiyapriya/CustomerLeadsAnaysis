"""
Quick analysis script for Raw File-LS-Full Data.xlsx
"""

from excel_analyzer import ExcelAnalyzer

# Your Excel file path
file_path = r"C:\Users\karul\Downloads\Raw File-LS-Full Data.xlsx"

print("=" * 60)
print("Excel Data Analyzer - Raw File-LS-Full Data")
print("=" * 60)

# Create analyzer instance
analyzer = ExcelAnalyzer(file_path)

# Load data
if analyzer.load_data():
    print("\nPerforming comprehensive analysis...")
    
    # Basic information
    print("\n📋 Basic Information:")
    info = analyzer.get_basic_info()
    print(f"  - Rows: {info['total_rows']}")
    print(f"  - Columns: {info['total_columns']}")
    print(f"  - Memory: {info['memory_usage']:.2f} MB")
    print(f"  - Columns: {', '.join(info['column_names'][:5])}{'...' if len(info['column_names']) > 5 else ''}")
    
    # Statistical summary
    print("\n📊 Statistical Summary:")
    analyzer.get_statistical_summary()
    print("  ✓ Calculated statistics for all columns")
    
    # Duplicates
    print("\n🔍 Checking for duplicates:")
    dup_info = analyzer.find_duplicates()
    print(f"  - Duplicate rows: {dup_info['total_duplicates']} ({dup_info['duplicate_percentage']:.2f}%)")
    
    # Missing data
    print("\n⚠️  Missing Data Analysis:")
    missing_info = analyzer.analyze_missing_data()
    print(f"  - Total missing values: {missing_info['total_missing']}")
    
    # Show columns with missing data
    if missing_info['total_missing'] > 0:
        print("\n  Columns with missing data:")
        for col, count in missing_info['missing_by_column'].items():
            if count > 0:
                pct = missing_info['missing_percentage'][col]
                print(f"    - {col}: {count} ({pct:.2f}%)")
    
    # Correlation
    print("\n🔗 Correlation Analysis:")
    analyzer.get_correlation_matrix()
    print("  ✓ Correlation matrix calculated")
    
    # Generate visualizations
    print("\n📈 Generating visualizations...")
    analyzer.generate_visualizations()
    
    # Generate reports
    print("\n📄 Generating reports...")
    analyzer.generate_html_report()
    analyzer.export_to_excel()
    
    print("\n" + "=" * 60)
    print("✓ Analysis complete!")
    print("=" * 60)
    print("\nGenerated files in 'reports' folder:")
    print("  - analysis_report.html (Open in browser)")
    print("  - analysis_summary.xlsx")
    print("  - correlation_heatmap.png")
    print("  - distributions.png")
    print("  - missing_data_heatmap.png (if missing data exists)")
    print("=" * 60)
else:
    print("\n✗ Failed to load the Excel file. Please check the file path.")
