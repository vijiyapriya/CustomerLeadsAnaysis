# Excel Data Analyzer and Report Generator

A comprehensive Python program for analyzing Excel files and generating customized reports with visualizations.

## Features

- 📊 **Statistical Analysis**: Comprehensive statistics for numerical and categorical data
- 📈 **Visualizations**: Correlation heatmaps, distribution plots, missing data analysis
- 📄 **HTML Reports**: Beautiful, interactive HTML reports
- 📑 **Excel Export**: Multi-sheet Excel reports with analysis results
- 🔍 **Data Quality Checks**: Duplicate detection, missing value analysis
- 📉 **Correlation Analysis**: Identify relationships between numerical variables

## Installation

1. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

### Method 1: Interactive Mode

Run the main program and follow the prompts:
```bash
python excel_analyzer.py
```

### Method 2: Programmatic Usage

```python
from excel_analyzer import ExcelAnalyzer

# Create analyzer instance
analyzer = ExcelAnalyzer('your_data.xlsx')

# Load data
analyzer.load_data()

# Perform analysis
analyzer.get_basic_info()
analyzer.get_statistical_summary()
analyzer.find_duplicates()
analyzer.analyze_missing_data()
analyzer.get_correlation_matrix()

# Generate reports
analyzer.generate_visualizations()
analyzer.generate_html_report()
analyzer.export_to_excel()
```

### Method 3: Use Sample Examples

Check `sample_usage.py` for various usage examples:
```python
python sample_usage.py
```

## Output

The program generates the following outputs in the `reports/` folder:

1. **HTML Report** (`analysis_report.html`): Interactive web-based report
2. **Excel Report** (`analysis_summary.xlsx`): Multi-sheet Excel workbook
3. **Visualizations**:
   - `correlation_heatmap.png`: Correlation matrix heatmap
   - `distributions.png`: Distribution plots for numerical columns
   - `missing_data_heatmap.png`: Missing data pattern visualization

## Analysis Components

### 1. Basic Information
- Total rows and columns
- Column names and data types
- Memory usage
- Missing value counts

### 2. Statistical Summary
- **Numerical columns**: Mean, median, std, min, max, quartiles
- **Categorical columns**: Unique values, most frequent values

### 3. Data Quality
- Duplicate row detection
- Missing data analysis
- Missing value percentages by column

### 4. Correlation Analysis
- Correlation matrix for numerical variables
- Heatmap visualization

### 5. Visualizations
- Missing data patterns
- Correlation heatmaps
- Distribution histograms

## Customization

You can customize the analysis by modifying parameters:

```python
# Load specific sheet
analyzer.load_data(sheet_name='Sheet2')

# Custom output directory
analyzer.generate_visualizations(output_dir='custom_output')

# Custom report file names
analyzer.generate_html_report(output_file='reports/my_report.html')
analyzer.export_to_excel(output_file='reports/my_analysis.xlsx')
```

## Requirements

- Python 3.7+
- pandas
- numpy
- matplotlib
- seaborn
- openpyxl

## Example Output Structure

```
reports/
├── analysis_report.html          # Main HTML report
├── analysis_summary.xlsx         # Excel summary
├── correlation_heatmap.png       # Correlation visualization
├── distributions.png             # Distribution plots
└── missing_data_heatmap.png     # Missing data visualization
```

## Tips

- Ensure your Excel file is not open in another program before running the analysis
- For large files (>100MB), analysis may take a few minutes
- The HTML report is interactive and best viewed in a modern web browser
- Missing data visualizations only generate if there are missing values

## Troubleshooting

**Issue**: `FileNotFoundError`
- **Solution**: Verify the Excel file path is correct and the file exists

**Issue**: `PermissionError`
- **Solution**: Close the Excel file if it's open in another program

**Issue**: Import errors
- **Solution**: Install all requirements: `pip install -r requirements.txt`

## License

Free to use and modify for your needs.
