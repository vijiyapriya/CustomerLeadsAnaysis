"""
Active Leads Analysis
Finds and analyzes all active leads from the dataset
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file
file_path = r"C:\Users\karul\Downloads\Raw File-LS-Full Data.xlsx"

print("=" * 70)
print("Active Leads Analysis")
print("=" * 70)

# Load data
print("\nLoading data...")
df = pd.read_excel(file_path)
print(f"✓ Loaded {len(df):,} rows and {len(df.columns)} columns")

# Check for Lead Stage column
print("\n" + "=" * 70)
print("Lead Stage Analysis")
print("=" * 70)

if 'Lead Stage' in df.columns:
    print(f"\nTotal records: {len(df):,}")
    print(f"Records with Lead Stage data: {df['Lead Stage'].notna().sum():,}")
    print(f"Records with missing Lead Stage: {df['Lead Stage'].isna().sum():,}")
    
    # Get unique Lead Stage values
    print(f"\nUnique Lead Stage values: {df['Lead Stage'].nunique()}")
    
    # Show all Lead Stage values
    print("\n📊 Lead Stage Distribution:")
    print("-" * 70)
    stage_counts = df['Lead Stage'].value_counts(dropna=False)
    print(f"{'Lead Stage':<40} {'Count':>12} {'Percentage':>12}")
    print("-" * 70)
    for stage, count in stage_counts.items():
        pct = (count / len(df)) * 100
        stage_name = str(stage) if pd.notna(stage) else "Missing/Unknown"
        print(f"{stage_name:<40} {count:>12,} {pct:>11.2f}%")
    
    # Filter for Active leads
    print("\n" + "=" * 70)
    print("Filtering Active Leads")
    print("=" * 70)
    
    # Look for "Active" in Lead Stage (case-insensitive)
    active_mask = df['Lead Stage'].astype(str).str.contains('active', case=False, na=False)
    active_leads = df[active_mask].copy()
    
    print(f"\n✓ Found {len(active_leads):,} Active leads")
    
    if len(active_leads) > 0:
        # Show unique active stages
        print("\n📋 Active Lead Stage Types:")
        print("-" * 70)
        active_stages = active_leads['Lead Stage'].value_counts()
        for stage, count in active_stages.items():
            print(f"  {str(stage):<50}: {count:>8,}")
        
        # Analyze Active Leads by Country
        print("\n" + "=" * 70)
        print("Active Leads by Country")
        print("=" * 70)
        
        country_counts = active_leads['Country'].value_counts(dropna=False)
        
        print(f"\nTotal active leads: {len(active_leads):,}")
        print(f"Countries represented: {active_leads['Country'].nunique()}")
        print(f"Records with missing country: {active_leads['Country'].isna().sum():,}")
        
        print("\n📍 Top 20 Countries with Active Leads:")
        print("-" * 70)
        print(f"{'Country':<35} {'Count':>10} {'Percentage':>12}")
        print("-" * 70)
        
        for country, count in country_counts.head(20).items():
            pct = (count / len(active_leads)) * 100
            country_name = str(country) if pd.notna(country) else "Missing/Unknown"
            print(f"{country_name:<35} {count:>10,} {pct:>11.2f}%")
        
        # Analyze Active Leads by Industry
        if 'Industry Vertical' in active_leads.columns:
            print("\n" + "=" * 70)
            print("Active Leads by Industry Vertical")
            print("=" * 70)
            
            industry_counts = active_leads['Industry Vertical'].value_counts(dropna=False)
            
            print(f"\n📊 Top 15 Industries with Active Leads:")
            print("-" * 70)
            print(f"{'Industry Vertical':<35} {'Count':>10} {'Percentage':>12}")
            print("-" * 70)
            
            for industry, count in industry_counts.head(15).items():
                pct = (count / len(active_leads)) * 100
                industry_name = str(industry) if pd.notna(industry) else "Missing/Unknown"
                print(f"{industry_name:<35} {count:>10,} {pct:>11.2f}%")
        
        # Analyze by Lead Source
        if 'Lead Source' in active_leads.columns:
            print("\n" + "=" * 70)
            print("Active Leads by Lead Source")
            print("=" * 70)
            
            source_counts = active_leads['Lead Source'].value_counts(dropna=False)
            
            print(f"\n📌 Lead Sources for Active Leads:")
            print("-" * 70)
            print(f"{'Lead Source':<35} {'Count':>10} {'Percentage':>12}")
            print("-" * 70)
            
            for source, count in source_counts.head(15).items():
                pct = (count / len(active_leads)) * 100
                source_name = str(source) if pd.notna(source) else "Missing/Unknown"
                print(f"{source_name:<35} {count:>10,} {pct:>11.2f}%")
        
        # Analyze by Company Size
        if 'Company size' in active_leads.columns:
            print("\n" + "=" * 70)
            print("Active Leads by Company Size")
            print("=" * 70)
            
            size_counts = active_leads['Company size'].value_counts(dropna=False)
            
            print(f"\n🏢 Company Size Distribution:")
            print("-" * 70)
            print(f"{'Company Size':<35} {'Count':>10} {'Percentage':>12}")
            print("-" * 70)
            
            for size, count in size_counts.items():
                pct = (count / len(active_leads)) * 100
                size_name = str(size) if pd.notna(size) else "Missing/Unknown"
                print(f"{size_name:<35} {count:>10,} {pct:>11.2f}%")
        
        # Analyze by Last Activity
        if 'Last Activity' in active_leads.columns:
            print("\n" + "=" * 70)
            print("Active Leads - Last Activity")
            print("=" * 70)
            
            activity_counts = active_leads['Last Activity'].value_counts(dropna=False)
            
            print(f"\n🔔 Top 10 Last Activities for Active Leads:")
            print("-" * 70)
            print(f"{'Last Activity':<35} {'Count':>10} {'Percentage':>12}")
            print("-" * 70)
            
            for activity, count in activity_counts.head(10).items():
                pct = (count / len(active_leads)) * 100
                activity_name = str(activity) if pd.notna(activity) else "Missing/Unknown"
                print(f"{activity_name:<35} {count:>10,} {pct:>11.2f}%")
        
        # Export to Excel
        print("\n" + "=" * 70)
        print("Exporting Results")
        print("=" * 70)
        
        with pd.ExcelWriter('reports/active_leads_analysis.xlsx', engine='openpyxl') as writer:
            # Sheet 1: All Active Leads
            active_leads.to_excel(writer, sheet_name='Active Leads', index=False)
            
            # Sheet 2: Summary by Country
            country_summary = pd.DataFrame({
                'Country': country_counts.index,
                'Count': country_counts.values,
                'Percentage': (country_counts.values / len(active_leads) * 100).round(2)
            })
            country_summary.to_excel(writer, sheet_name='By Country', index=False)
            
            # Sheet 3: Summary by Industry
            if 'Industry Vertical' in active_leads.columns:
                industry_summary = pd.DataFrame({
                    'Industry Vertical': industry_counts.index,
                    'Count': industry_counts.values,
                    'Percentage': (industry_counts.values / len(active_leads) * 100).round(2)
                })
                industry_summary.to_excel(writer, sheet_name='By Industry', index=False)
            
            # Sheet 4: Summary by Lead Source
            if 'Lead Source' in active_leads.columns:
                source_summary = pd.DataFrame({
                    'Lead Source': source_counts.index,
                    'Count': source_counts.values,
                    'Percentage': (source_counts.values / len(active_leads) * 100).round(2)
                })
                source_summary.to_excel(writer, sheet_name='By Lead Source', index=False)
            
            # Sheet 5: Summary by Company Size
            if 'Company size' in active_leads.columns:
                size_summary = pd.DataFrame({
                    'Company Size': size_counts.index,
                    'Count': size_counts.values,
                    'Percentage': (size_counts.values / len(active_leads) * 100).round(2)
                })
                size_summary.to_excel(writer, sheet_name='By Company Size', index=False)
            
            # Sheet 6: Summary Statistics
            summary_stats = pd.DataFrame({
                'Metric': [
                    'Total Active Leads',
                    'Countries Represented',
                    'Industries Represented',
                    'Lead Sources',
                    'Active % of Total Dataset'
                ],
                'Value': [
                    len(active_leads),
                    active_leads['Country'].nunique(),
                    active_leads['Industry Vertical'].nunique() if 'Industry Vertical' in active_leads.columns else 0,
                    active_leads['Lead Source'].nunique() if 'Lead Source' in active_leads.columns else 0,
                    f"{(len(active_leads) / len(df) * 100):.2f}%"
                ]
            })
            summary_stats.to_excel(writer, sheet_name='Summary', index=False)
        
        print("✓ Exported to: reports/active_leads_analysis.xlsx")
        
        # Create visualizations
        print("\n📊 Creating visualizations...")
        
        # 1. Active Leads by Country (Top 15)
        plt.figure(figsize=(14, 8))
        top_countries = country_counts.head(15)
        colors = sns.color_palette("viridis", len(top_countries))
        bars = plt.barh(range(len(top_countries)), top_countries.values, color=colors)
        plt.yticks(range(len(top_countries)), [str(c) if pd.notna(c) else 'Unknown' for c in top_countries.index])
        plt.xlabel('Number of Active Leads', fontsize=12, fontweight='bold')
        plt.ylabel('Country', fontsize=12, fontweight='bold')
        plt.title('Top 15 Countries - Active Leads', fontsize=14, fontweight='bold', pad=20)
        plt.gca().invert_yaxis()
        
        for i, (bar, value) in enumerate(zip(bars, top_countries.values)):
            plt.text(value, i, f' {value:,}', va='center', fontsize=10)
        
        plt.tight_layout()
        plt.savefig('reports/active_leads_by_country.png', dpi=300, bbox_inches='tight')
        plt.close()
        print("✓ Saved: reports/active_leads_by_country.png")
        
        # 2. Active Leads by Industry (if available)
        if 'Industry Vertical' in active_leads.columns and len(industry_counts) > 0:
            plt.figure(figsize=(14, 8))
            top_industries = industry_counts.head(12)
            colors = sns.color_palette("coolwarm", len(top_industries))
            bars = plt.barh(range(len(top_industries)), top_industries.values, color=colors)
            plt.yticks(range(len(top_industries)), [str(i)[:40] if pd.notna(i) else 'Unknown' for i in top_industries.index])
            plt.xlabel('Number of Active Leads', fontsize=12, fontweight='bold')
            plt.ylabel('Industry Vertical', fontsize=12, fontweight='bold')
            plt.title('Top 12 Industries - Active Leads', fontsize=14, fontweight='bold', pad=20)
            plt.gca().invert_yaxis()
            
            for i, (bar, value) in enumerate(zip(bars, top_industries.values)):
                plt.text(value, i, f' {value:,}', va='center', fontsize=10)
            
            plt.tight_layout()
            plt.savefig('reports/active_leads_by_industry.png', dpi=300, bbox_inches='tight')
            plt.close()
            print("✓ Saved: reports/active_leads_by_industry.png")
        
        # 3. Country Distribution Pie Chart
        plt.figure(figsize=(12, 8))
        top_10_countries = country_counts.head(10)
        others_count = country_counts[10:].sum()
        
        if others_count > 0:
            plot_data = pd.concat([top_10_countries, pd.Series({'Others': others_count})])
        else:
            plot_data = top_10_countries
        
        colors = sns.color_palette("Set3", len(plot_data))
        wedges, texts, autotexts = plt.pie(plot_data.values, 
                                           labels=[str(c) if pd.notna(c) else 'Unknown' for c in plot_data.index],
                                           autopct='%1.1f%%', colors=colors, startangle=90)
        
        for text in texts:
            text.set_fontsize(10)
            text.set_fontweight('bold')
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontsize(9)
            autotext.set_fontweight('bold')
        
        plt.title('Active Leads Distribution by Country (Top 10)', fontsize=14, fontweight='bold', pad=20)
        plt.tight_layout()
        plt.savefig('reports/active_leads_country_pie.png', dpi=300, bbox_inches='tight')
        plt.close()
        print("✓ Saved: reports/active_leads_country_pie.png")
        
    else:
        print("\n⚠️  No Active leads found")
        print("\nShowing all Lead Stage values for reference.")
        
else:
    print("\n✗ 'Lead Stage' column not found in the dataset")
    print(f"\nAvailable columns: {', '.join(df.columns)}")

print("\n" + "=" * 70)
print("✓ Analysis Complete!")
print("=" * 70)
print("\nGenerated files:")
print("  - reports/active_leads_analysis.xlsx")
print("  - reports/active_leads_by_country.png")
print("  - reports/active_leads_by_industry.png")
print("  - reports/active_leads_country_pie.png")
print("=" * 70)
