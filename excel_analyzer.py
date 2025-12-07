"""
Excel Data Analyzer and Report Generator
This program reads Excel files, performs data analysis, and generates customized reports.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os


class ExcelAnalyzer:
    """Class to analyze Excel data and generate reports"""
    
    def __init__(self, file_path):
        """
        Initialize the analyzer with an Excel file
        
        Args:
            file_path (str): Path to the Excel file
        """
        self.file_path = file_path
        self.df = None
        self.report = {}
        
    def load_data(self, sheet_name=0):
        """
        Load data from Excel file
        
        Args:
            sheet_name: Sheet name or index to load (default: 0)
        """
        try:
            self.df = pd.read_excel(self.file_path, sheet_name=sheet_name)
            print(f"✓ Successfully loaded data from {self.file_path}")
            print(f"  Shape: {self.df.shape[0]} rows × {self.df.shape[1]} columns")
            return True
        except Exception as e:
            print(f"✗ Error loading file: {e}")
            return False
    
    def get_basic_info(self):
        """Get basic information about the dataset"""
        if self.df is None:
            print("No data loaded. Please load data first.")
            return
        
        info = {
            'total_rows': len(self.df),
            'total_columns': len(self.df.columns),
            'column_names': list(self.df.columns),
            'data_types': self.df.dtypes.to_dict(),
            'missing_values': self.df.isnull().sum().to_dict(),
            'memory_usage': self.df.memory_usage(deep=True).sum() / 1024**2  # MB
        }
        
        self.report['basic_info'] = info
        return info
    
    def get_statistical_summary(self):
        """Get statistical summary of numerical columns"""
        if self.df is None:
            print("No data loaded. Please load data first.")
            return
        
        # Numerical columns summary
        numeric_summary = self.df.describe().to_dict()
        
        # Categorical columns summary
        categorical_cols = self.df.select_dtypes(include=['object']).columns
        categorical_summary = {}
        for col in categorical_cols:
            categorical_summary[col] = {
                'unique_values': self.df[col].nunique(),
                'top_value': self.df[col].mode()[0] if len(self.df[col].mode()) > 0 else None,
                'frequency': self.df[col].value_counts().to_dict()
            }
        
        self.report['numerical_summary'] = numeric_summary
        self.report['categorical_summary'] = categorical_summary
        
        return {
            'numerical': numeric_summary,
            'categorical': categorical_summary
        }
    
    def find_duplicates(self):
        """Find duplicate rows in the dataset"""
        if self.df is None:
            print("No data loaded. Please load data first.")
            return
        
        duplicates = self.df[self.df.duplicated()]
        duplicate_info = {
            'total_duplicates': len(duplicates),
            'duplicate_percentage': (len(duplicates) / len(self.df)) * 100
        }
        
        self.report['duplicates'] = duplicate_info
        return duplicate_info
    
    def analyze_missing_data(self):
        """Analyze missing data patterns"""
        if self.df is None:
            print("No data loaded. Please load data first.")
            return
        
        missing_data = {
            'total_missing': self.df.isnull().sum().sum(),
            'missing_by_column': self.df.isnull().sum().to_dict(),
            'missing_percentage': (self.df.isnull().sum() / len(self.df) * 100).to_dict()
        }
        
        self.report['missing_data'] = missing_data
        return missing_data
    
    def get_correlation_matrix(self):
        """Calculate correlation matrix for numerical columns"""
        if self.df is None:
            print("No data loaded. Please load data first.")
            return
        
        numeric_df = self.df.select_dtypes(include=[np.number])
        if len(numeric_df.columns) > 1:
            correlation = numeric_df.corr().to_dict()
            self.report['correlation'] = correlation
            return correlation
        else:
            print("Not enough numerical columns for correlation analysis.")
            return None
    
    def generate_visualizations(self, output_dir='reports'):
        """Generate visualization charts"""
        if self.df is None:
            print("No data loaded. Please load data first.")
            return
        
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        # Set style
        sns.set_style("whitegrid")
        
        # 1. Missing data heatmap
        if self.df.isnull().sum().sum() > 0:
            plt.figure(figsize=(12, 6))
            sns.heatmap(self.df.isnull(), cbar=True, yticklabels=False, cmap='viridis')
            plt.title('Missing Data Heatmap')
            plt.tight_layout()
            plt.savefig(f'{output_dir}/missing_data_heatmap.png', dpi=300)
            plt.close()
        
        # 2. Correlation heatmap for numerical columns
        numeric_df = self.df.select_dtypes(include=[np.number])
        if len(numeric_df.columns) > 1:
            plt.figure(figsize=(10, 8))
            sns.heatmap(numeric_df.corr(), annot=True, cmap='coolwarm', center=0, fmt='.2f')
            plt.title('Correlation Matrix')
            plt.tight_layout()
            plt.savefig(f'{output_dir}/correlation_heatmap.png', dpi=300)
            plt.close()
        
        # 3. Distribution plots for numerical columns
        numeric_cols = numeric_df.columns[:6]  # Limit to first 6 columns
        if len(numeric_cols) > 0:
            fig, axes = plt.subplots(2, 3, figsize=(15, 10))
            axes = axes.flatten()
            
            for idx, col in enumerate(numeric_cols):
                if idx < len(axes):
                    self.df[col].hist(bins=30, ax=axes[idx], edgecolor='black')
                    axes[idx].set_title(f'Distribution of {col}')
                    axes[idx].set_xlabel(col)
                    axes[idx].set_ylabel('Frequency')
            
            # Hide unused subplots
            for idx in range(len(numeric_cols), len(axes)):
                axes[idx].set_visible(False)
            
            plt.tight_layout()
            plt.savefig(f'{output_dir}/distributions.png', dpi=300)
            plt.close()
        
        print(f"✓ Visualizations saved to {output_dir}/ directory")
    
    def generate_html_report(self, output_file='reports/analysis_report.html'):
        """Generate an HTML report with all analysis results"""
        if self.df is None:
            print("No data loaded. Please load data first.")
            return
        
        # Ensure reports directory exists
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Excel Data Analysis Report</title>
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    margin: 20px;
                    background-color: #f5f5f5;
                }}
                .container {{
                    max-width: 1200px;
                    margin: 0 auto;
                    background-color: white;
                    padding: 30px;
                    box-shadow: 0 0 10px rgba(0,0,0,0.1);
                }}
                h1 {{
                    color: #2c3e50;
                    border-bottom: 3px solid #3498db;
                    padding-bottom: 10px;
                }}
                h2 {{
                    color: #34495e;
                    margin-top: 30px;
                    border-left: 4px solid #3498db;
                    padding-left: 10px;
                }}
                table {{
                    border-collapse: collapse;
                    width: 100%;
                    margin: 20px 0;
                }}
                th, td {{
                    border: 1px solid #ddd;
                    padding: 12px;
                    text-align: left;
                }}
                th {{
                    background-color: #3498db;
                    color: white;
                }}
                tr:nth-child(even) {{
                    background-color: #f2f2f2;
                }}
                .metric {{
                    background-color: #ecf0f1;
                    padding: 15px;
                    margin: 10px 0;
                    border-radius: 5px;
                }}
                .metric-value {{
                    font-size: 24px;
                    font-weight: bold;
                    color: #2980b9;
                }}
                img {{
                    max-width: 100%;
                    height: auto;
                    margin: 20px 0;
                    border: 1px solid #ddd;
                    border-radius: 5px;
                }}
                .footer {{
                    margin-top: 40px;
                    text-align: center;
                    color: #7f8c8d;
                    font-size: 12px;
                }}
            </style>
        </head>
        <body>
            <div class="container">
                <h1>📊 Excel Data Analysis Report</h1>
                <p><strong>Generated on:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
                <p><strong>File:</strong> {self.file_path}</p>
                
                <h2>1. Dataset Overview</h2>
                <div class="metric">
                    <p>Total Rows: <span class="metric-value">{len(self.df)}</span></p>
                    <p>Total Columns: <span class="metric-value">{len(self.df.columns)}</span></p>
                    <p>Memory Usage: <span class="metric-value">{self.df.memory_usage(deep=True).sum() / 1024**2:.2f} MB</span></p>
                </div>
                
                <h3>Column Information</h3>
                <table>
                    <tr>
                        <th>Column Name</th>
                        <th>Data Type</th>
                        <th>Missing Values</th>
                        <th>Missing %</th>
                    </tr>
        """
        
        # Add column information
        for col in self.df.columns:
            missing = self.df[col].isnull().sum()
            missing_pct = (missing / len(self.df)) * 100
            html_content += f"""
                    <tr>
                        <td>{col}</td>
                        <td>{self.df[col].dtype}</td>
                        <td>{missing}</td>
                        <td>{missing_pct:.2f}%</td>
                    </tr>
            """
        
        html_content += """
                </table>
                
                <h2>2. Statistical Summary</h2>
        """
        
        # Add numerical summary
        numeric_df = self.df.select_dtypes(include=[np.number])
        if len(numeric_df.columns) > 0:
            html_content += "<h3>Numerical Columns</h3>"
            html_content += numeric_df.describe().to_html()
        
        # Add categorical summary
        categorical_df = self.df.select_dtypes(include=['object'])
        if len(categorical_df.columns) > 0:
            html_content += "<h3>Categorical Columns</h3><table><tr><th>Column</th><th>Unique Values</th><th>Most Frequent</th></tr>"
            for col in categorical_df.columns:
                unique = self.df[col].nunique()
                top = self.df[col].mode()[0] if len(self.df[col].mode()) > 0 else "N/A"
                html_content += f"<tr><td>{col}</td><td>{unique}</td><td>{top}</td></tr>"
            html_content += "</table>"
        
        # Add data quality section
        duplicates = len(self.df[self.df.duplicated()])
        html_content += f"""
                <h2>3. Data Quality</h2>
                <div class="metric">
                    <p>Duplicate Rows: <span class="metric-value">{duplicates}</span> ({(duplicates/len(self.df)*100):.2f}%)</p>
                    <p>Total Missing Values: <span class="metric-value">{self.df.isnull().sum().sum()}</span></p>
                </div>
        """
        
        # Add visualizations if they exist
        html_content += """
                <h2>4. Visualizations</h2>
        """
        
        viz_files = ['correlation_heatmap.png', 'distributions.png', 'missing_data_heatmap.png']
        for viz_file in viz_files:
            viz_path = f'reports/{viz_file}'
            if os.path.exists(viz_path):
                html_content += f'<img src="{viz_file}" alt="{viz_file}">'
        
        # Add data preview
        html_content += f"""
                <h2>5. Data Preview (First 10 Rows)</h2>
                {self.df.head(10).to_html()}
                
                <div class="footer">
                    <p>Report generated by Excel Analyzer Tool</p>
                </div>
            </div>
        </body>
        </html>
        """
        
        # Write to file
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"✓ HTML report saved to {output_file}")
        return output_file
    
    def export_to_excel(self, output_file='reports/analysis_summary.xlsx'):
        """Export analysis results to Excel with multiple sheets"""
        if self.df is None:
            print("No data loaded. Please load data first.")
            return
        
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Original data
            self.df.to_excel(writer, sheet_name='Original Data', index=False)
            
            # Statistical summary
            numeric_df = self.df.select_dtypes(include=[np.number])
            if len(numeric_df.columns) > 0:
                numeric_df.describe().to_excel(writer, sheet_name='Statistical Summary')
            
            # Missing data analysis
            missing_df = pd.DataFrame({
                'Column': self.df.columns,
                'Missing Count': self.df.isnull().sum().values,
                'Missing %': (self.df.isnull().sum() / len(self.df) * 100).values
            })
            missing_df.to_excel(writer, sheet_name='Missing Data', index=False)
            
            # Correlation matrix
            if len(numeric_df.columns) > 1:
                numeric_df.corr().to_excel(writer, sheet_name='Correlation Matrix')
        
        print(f"✓ Excel report saved to {output_file}")
        return output_file


def main():
    """Main function to demonstrate usage"""
    print("=" * 60)
    print("Excel Data Analyzer and Report Generator")
    print("=" * 60)
    
    # Example usage
    file_path = input("\nEnter the path to your Excel file: ").strip()
    
    if not os.path.exists(file_path):
        print(f"✗ File not found: {file_path}")
        return
    
    # Create analyzer instance
    analyzer = ExcelAnalyzer(file_path)
    
    # Load data
    if not analyzer.load_data():
        return
    
    print("\nPerforming analysis...")
    
    # Basic information
    print("\n📋 Basic Information:")
    info = analyzer.get_basic_info()
    print(f"  - Rows: {info['total_rows']}")
    print(f"  - Columns: {info['total_columns']}")
    print(f"  - Memory: {info['memory_usage']:.2f} MB")
    
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
    print("✓ Analysis complete! Check the 'reports' folder for outputs.")
    print("=" * 60)


if __name__ == "__main__":
    main()
