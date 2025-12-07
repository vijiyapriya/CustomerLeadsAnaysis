"""
Role Analysis: HR, IT, Finance Leads, CEO, CFO by Country
Analyzes specific roles and provides country-wise breakdown
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file (using the final updated file)
file_path = r"reports/Raw_File_LS_Updated_Regions_Final.xlsx"

print("=" * 70)
print("Role Analysis: HR, IT, Finance Leads, CEO, CFO by Country")
print("=" * 70)

# Load data
print("\nLoading data...")
df = pd.read_excel(file_path)
print(f"✓ Loaded {len(df):,} rows and {len(df.columns)} columns")

# Check Role column
print("\n" + "=" * 70)
print("Role Column Analysis")
print("=" * 70)

if 'Role' in df.columns:
    print(f"\nTotal records: {len(df):,}")
    print(f"Records with Role data: {df['Role'].notna().sum():,}")
    print(f"Records with missing Role: {df['Role'].isna().sum():,}")
    print(f"Unique Role values: {df['Role'].nunique()}")
    
    # Show sample of role values
    print("\n📊 Sample Role Values (Top 20):")
    print("-" * 70)
    role_counts = df['Role'].value_counts(dropna=False)
    for role, count in role_counts.head(20).items():
        role_name = str(role)[:50] if pd.notna(role) else "Missing"
        print(f"  {role_name:<50}: {count:>8,}")
    
    # Define search terms for each role category
    role_categories = {
        'HR Leads': ['hr', 'human resource', 'human capital', 'people', 'talent'],
        'IT Leads': ['it ', 'information technology', 'technology', 'tech ', 'cio', 'chief information'],
        'Finance Leads': ['finance', 'financial', 'cfo', 'chief financial'],
        'CEO': ['ceo', 'chief executive'],
        'CFO': ['cfo', 'chief financial officer']
    }
    
    print("\n" + "=" * 70)
    print("Filtering by Role Categories")
    print("=" * 70)
    
    # Create masks for each category
    role_results = {}
    
    for category, keywords in role_categories.items():
        # Create a mask that checks if any keyword is in the role (case-insensitive)
        mask = pd.Series([False] * len(df), index=df.index)
        
        for keyword in keywords:
            keyword_mask = df['Role'].astype(str).str.contains(keyword, case=False, na=False)
            mask = mask | keyword_mask
        
        filtered_df = df[mask].copy()
        role_results[category] = filtered_df
        
        print(f"\n✓ {category}: {len(filtered_df):,} records found")
        
        if len(filtered_df) > 0:
            # Show sample roles found
            sample_roles = filtered_df['Role'].value_counts().head(10)
            print(f"  Top roles in this category:")
            for role, count in sample_roles.items():
                role_name = str(role)[:45] if pd.notna(role) else "Missing"
                print(f"    • {role_name:<45}: {count:>6,}")
    
    # Analyze by Country
    print("\n" + "=" * 70)
    print("Country-wise Analysis")
    print("=" * 70)
    
    # Create summary dataframe
    country_summary = []
    
    for category, filtered_df in role_results.items():
        if len(filtered_df) > 0:
            country_counts = filtered_df['Country'].value_counts(dropna=False)
            
            print(f"\n📍 {category} by Country (Top 20):")
            print("-" * 70)
            print(f"{'Country':<35} {'Count':>10} {'%':>8}")
            print("-" * 70)
            
            for country, count in country_counts.head(20).items():
                pct = (count / len(filtered_df)) * 100
                country_name = str(country) if pd.notna(country) else "Missing/Unknown"
                print(f"{country_name:<35} {count:>10,} {pct:>7.2f}%")
                
                country_summary.append({
                    'Role Category': category,
                    'Country': country_name,
                    'Count': count,
                    'Percentage': round(pct, 2)
                })
    
    # Create pivot table for all roles by country
    print("\n" + "=" * 70)
    print("Combined Country Summary (Top 30 Countries)")
    print("=" * 70)
    
    # Get top countries across all categories
    all_countries = {}
    for category, filtered_df in role_results.items():
        if len(filtered_df) > 0:
            for country in filtered_df['Country'].dropna().unique():
                if country not in all_countries:
                    all_countries[country] = 0
                all_countries[country] += len(filtered_df[filtered_df['Country'] == country])
    
    # Sort and get top 30
    top_countries = sorted(all_countries.items(), key=lambda x: x[1], reverse=True)[:30]
    
    # Create pivot table
    print(f"\n{'Country':<25} {'HR':>8} {'IT':>8} {'Finance':>8} {'CEO':>8} {'CFO':>8} {'Total':>10}")
    print("-" * 90)
    
    pivot_data = []
    for country, _ in top_countries:
        row = {'Country': country}
        total = 0
        for category, filtered_df in role_results.items():
            count = len(filtered_df[filtered_df['Country'] == country])
            row[category] = count
            total += count
        row['Total'] = total
        pivot_data.append(row)
        
        print(f"{country:<25} "
              f"{row.get('HR Leads', 0):>8,} "
              f"{row.get('IT Leads', 0):>8,} "
              f"{row.get('Finance Leads', 0):>8,} "
              f"{row.get('CEO', 0):>8,} "
              f"{row.get('CFO', 0):>8,} "
              f"{row['Total']:>10,}")
    
    # Calculate totals
    total_row = {'Country': 'TOTAL'}
    grand_total = 0
    for category, filtered_df in role_results.items():
        total_row[category] = len(filtered_df)
        grand_total += len(filtered_df)
    total_row['Total'] = grand_total
    
    print("-" * 90)
    print(f"{'TOTAL':<25} "
          f"{total_row.get('HR Leads', 0):>8,} "
          f"{total_row.get('IT Leads', 0):>8,} "
          f"{total_row.get('Finance Leads', 0):>8,} "
          f"{total_row.get('CEO', 0):>8,} "
          f"{total_row.get('CFO', 0):>8,} "
          f"{total_row['Total']:>10,}")
    
    # Export to Excel
    print("\n" + "=" * 70)
    print("Exporting Results")
    print("=" * 70)
    
    with pd.ExcelWriter('reports/role_analysis_by_country.xlsx', engine='openpyxl') as writer:
        # Sheet 1: Overall Summary
        summary_df = pd.DataFrame([
            {'Role Category': cat, 'Total Count': len(df_filtered)}
            for cat, df_filtered in role_results.items()
        ])
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Sheet 2: Country-wise breakdown
        country_summary_df = pd.DataFrame(country_summary)
        country_summary_df.to_excel(writer, sheet_name='By Country', index=False)
        
        # Sheet 3: Pivot table
        pivot_df = pd.DataFrame(pivot_data)
        pivot_df.to_excel(writer, sheet_name='Country Pivot', index=False)
        
        # Sheet 4-8: Individual role category details
        for category, filtered_df in role_results.items():
            if len(filtered_df) > 0:
                sheet_name = category.replace(' ', '_')[:31]  # Excel sheet name limit
                filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Sheet: HR Leads by Country
        if len(role_results['HR Leads']) > 0:
            hr_by_country = role_results['HR Leads']['Country'].value_counts().reset_index()
            hr_by_country.columns = ['Country', 'Count']
            hr_by_country.to_excel(writer, sheet_name='HR by Country', index=False)
        
        # Sheet: IT Leads by Country
        if len(role_results['IT Leads']) > 0:
            it_by_country = role_results['IT Leads']['Country'].value_counts().reset_index()
            it_by_country.columns = ['Country', 'Count']
            it_by_country.to_excel(writer, sheet_name='IT by Country', index=False)
        
        # Sheet: Finance Leads by Country
        if len(role_results['Finance Leads']) > 0:
            fin_by_country = role_results['Finance Leads']['Country'].value_counts().reset_index()
            fin_by_country.columns = ['Country', 'Count']
            fin_by_country.to_excel(writer, sheet_name='Finance by Country', index=False)
        
        # Sheet: CEO by Country
        if len(role_results['CEO']) > 0:
            ceo_by_country = role_results['CEO']['Country'].value_counts().reset_index()
            ceo_by_country.columns = ['Country', 'Count']
            ceo_by_country.to_excel(writer, sheet_name='CEO by Country', index=False)
        
        # Sheet: CFO by Country
        if len(role_results['CFO']) > 0:
            cfo_by_country = role_results['CFO']['Country'].value_counts().reset_index()
            cfo_by_country.columns = ['Country', 'Count']
            cfo_by_country.to_excel(writer, sheet_name='CFO by Country', index=False)
    
    print("✓ Exported to: reports/role_analysis_by_country.xlsx")
    
    # Create visualizations
    print("\n📊 Creating visualizations...")
    
    # 1. Bar chart - Total by Role Category
    plt.figure(figsize=(12, 6))
    categories = list(role_results.keys())
    counts = [len(df) for df in role_results.values()]
    colors = sns.color_palette("viridis", len(categories))
    bars = plt.bar(categories, counts, color=colors)
    plt.xlabel('Role Category', fontsize=12, fontweight='bold')
    plt.ylabel('Count', fontsize=12, fontweight='bold')
    plt.title('Total Count by Role Category', fontsize=14, fontweight='bold', pad=20)
    plt.xticks(rotation=15, ha='right')
    
    # Add value labels
    for bar, value in zip(bars, counts):
        height = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2., height,
                f'{int(value):,}', ha='center', va='bottom', fontsize=10)
    
    plt.tight_layout()
    plt.savefig('reports/role_category_totals.png', dpi=300, bbox_inches='tight')
    plt.close()
    print("✓ Saved: reports/role_category_totals.png")
    
    # 2. Stacked bar chart - Top 15 countries
    if pivot_data:
        plt.figure(figsize=(14, 8))
        top_15_data = pivot_data[:15]
        countries_plot = [d['Country'][:25] for d in top_15_data]
        
        hr_counts = [d.get('HR Leads', 0) for d in top_15_data]
        it_counts = [d.get('IT Leads', 0) for d in top_15_data]
        fin_counts = [d.get('Finance Leads', 0) for d in top_15_data]
        ceo_counts = [d.get('CEO', 0) for d in top_15_data]
        cfo_counts = [d.get('CFO', 0) for d in top_15_data]
        
        x = range(len(countries_plot))
        width = 0.15
        
        plt.bar([i - 2*width for i in x], hr_counts, width, label='HR Leads', color='#FF6B6B')
        plt.bar([i - width for i in x], it_counts, width, label='IT Leads', color='#4ECDC4')
        plt.bar(x, fin_counts, width, label='Finance Leads', color='#45B7D1')
        plt.bar([i + width for i in x], ceo_counts, width, label='CEO', color='#FFA07A')
        plt.bar([i + 2*width for i in x], cfo_counts, width, label='CFO', color='#98D8C8')
        
        plt.xlabel('Country', fontsize=12, fontweight='bold')
        plt.ylabel('Count', fontsize=12, fontweight='bold')
        plt.title('Role Distribution by Country (Top 15)', fontsize=14, fontweight='bold', pad=20)
        plt.xticks(x, countries_plot, rotation=45, ha='right')
        plt.legend(loc='upper right')
        plt.tight_layout()
        plt.savefig('reports/roles_by_country_top15.png', dpi=300, bbox_inches='tight')
        plt.close()
        print("✓ Saved: reports/roles_by_country_top15.png")
    
    # 3. Individual pie charts for each role
    fig, axes = plt.subplots(2, 3, figsize=(18, 12))
    axes = axes.flatten()
    
    for idx, (category, filtered_df) in enumerate(role_results.items()):
        if len(filtered_df) > 0 and idx < len(axes):
            country_counts = filtered_df['Country'].value_counts().head(10)
            others = filtered_df['Country'].value_counts()[10:].sum()
            
            if others > 0:
                plot_data = pd.concat([country_counts, pd.Series({'Others': others})])
            else:
                plot_data = country_counts
            
            colors = sns.color_palette("Set3", len(plot_data))
            wedges, texts, autotexts = axes[idx].pie(plot_data.values,
                                                      labels=[str(c)[:20] if pd.notna(c) else 'Unknown' for c in plot_data.index],
                                                      autopct='%1.1f%%',
                                                      colors=colors,
                                                      startangle=90)
            
            for text in texts:
                text.set_fontsize(8)
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontsize(7)
                autotext.set_fontweight('bold')
            
            axes[idx].set_title(f'{category} (Total: {len(filtered_df):,})', fontsize=11, fontweight='bold')
    
    # Hide unused subplot
    if len(role_results) < len(axes):
        axes[-1].set_visible(False)
    
    plt.tight_layout()
    plt.savefig('reports/role_country_distribution_pies.png', dpi=300, bbox_inches='tight')
    plt.close()
    print("✓ Saved: reports/role_country_distribution_pies.png")
    
else:
    print("\n✗ 'Role' column not found in the dataset")
    print(f"\nAvailable columns: {', '.join(df.columns)}")

print("\n" + "=" * 70)
print("✓ Analysis Complete!")
print("=" * 70)

print("\n📊 Summary:")
for category, filtered_df in role_results.items():
    print(f"  • {category}: {len(filtered_df):,} records")

print("\n📁 Generated files:")
print("  - reports/role_analysis_by_country.xlsx (Multi-sheet workbook)")
print("  - reports/role_category_totals.png")
print("  - reports/roles_by_country_top15.png")
print("  - reports/role_country_distribution_pies.png")
print("=" * 70)
