"""
Custom Analysis: Email Bounced Status by Country
Analyzes Last Activity for email bounced status and counts by country
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file
file_path = r"C:\Users\karul\Downloads\Raw File-LS-Full Data.xlsx"

print("=" * 70)
print("Email Bounced Status Analysis by Country")
print("=" * 70)

# Load data
print("\nLoading data...")
df = pd.read_excel(file_path)
print(f"✓ Loaded {len(df):,} rows and {len(df.columns)} columns")

# Check the Last Activity column
print("\n" + "=" * 70)
print("Last Activity Analysis")
print("=" * 70)

# Get unique values in Last Activity
if 'Last Activity' in df.columns:
    print(f"\nTotal records: {len(df):,}")
    print(f"Records with Last Activity data: {df['Last Activity'].notna().sum():,}")
    print(f"Records with missing Last Activity: {df['Last Activity'].isna().sum():,}")
    
    # Find all unique Last Activity values
    print(f"\nUnique Last Activity values: {df['Last Activity'].nunique()}")
    
    # Show value counts for Last Activity
    print("\n📊 Last Activity Value Counts:")
    print("-" * 70)
    activity_counts = df['Last Activity'].value_counts(dropna=False)
    for activity, count in activity_counts.head(20).items():
        pct = (count / len(df)) * 100
        print(f"  {str(activity)[:50]:50s}: {count:>8,} ({pct:>6.2f}%)")
    
    if len(activity_counts) > 20:
        print(f"  ... and {len(activity_counts) - 20} more unique values")
    
    # Filter for email bounced activities
    print("\n" + "=" * 70)
    print("Filtering for Email Bounced Status")
    print("=" * 70)
    
    # Search for bounced-related activities (case-insensitive)
    bounced_mask = df['Last Activity'].astype(str).str.contains('bounce', case=False, na=False)
    bounced_df = df[bounced_mask].copy()
    
    print(f"\n✓ Found {len(bounced_df):,} records with 'bounce' in Last Activity")
    
    if len(bounced_df) > 0:
        # Show unique bounced activities
        print("\n📧 Email Bounced Activity Types:")
        print("-" * 70)
        bounced_activities = bounced_df['Last Activity'].value_counts()
        for activity, count in bounced_activities.items():
            print(f"  {str(activity)[:50]:50s}: {count:>8,}")
        
        # Count by Country
        print("\n" + "=" * 70)
        print("Email Bounced Count by Country")
        print("=" * 70)
        
        country_counts = bounced_df['Country'].value_counts(dropna=False)
        
        print(f"\nTotal bounced emails: {len(bounced_df):,}")
        print(f"Countries represented: {bounced_df['Country'].nunique()}")
        print(f"Records with missing country: {bounced_df['Country'].isna().sum():,}")
        
        print("\n📍 Top Countries with Email Bounced:")
        print("-" * 70)
        print(f"{'Country':<30} {'Count':>10} {'Percentage':>12}")
        print("-" * 70)
        
        for country, count in country_counts.head(30).items():
            pct = (count / len(bounced_df)) * 100
            country_name = str(country) if pd.notna(country) else "Missing/Unknown"
            print(f"{country_name:<30} {count:>10,} {pct:>11.2f}%")
        
        if len(country_counts) > 30:
            remaining = len(country_counts) - 30
            remaining_count = country_counts[30:].sum()
            pct = (remaining_count / len(bounced_df)) * 100
            print(f"{'... Other countries':<30} {remaining_count:>10,} {pct:>11.2f}%")
        
        # Export to Excel
        print("\n" + "=" * 70)
        print("Exporting Results")
        print("=" * 70)
        
        # Create detailed report
        with pd.ExcelWriter('reports/email_bounced_analysis.xlsx', engine='openpyxl') as writer:
            # Sheet 1: Summary by Country
            country_summary = pd.DataFrame({
                'Country': country_counts.index,
                'Bounced Count': country_counts.values,
                'Percentage': (country_counts.values / len(bounced_df) * 100).round(2)
            })
            country_summary.to_excel(writer, sheet_name='Bounced by Country', index=False)
            
            # Sheet 2: Bounced Activity Types
            activity_summary = pd.DataFrame({
                'Last Activity': bounced_activities.index,
                'Count': bounced_activities.values,
                'Percentage': (bounced_activities.values / len(bounced_df) * 100).round(2)
            })
            activity_summary.to_excel(writer, sheet_name='Activity Types', index=False)
            
            # Sheet 3: Detailed bounced records
            bounced_df.to_excel(writer, sheet_name='Bounced Records', index=False)
            
            # Sheet 4: Country + Activity breakdown
            country_activity = bounced_df.groupby(['Country', 'Last Activity']).size().reset_index(name='Count')
            country_activity = country_activity.sort_values('Count', ascending=False)
            country_activity.to_excel(writer, sheet_name='Country + Activity', index=False)
        
        print("✓ Exported to: reports/email_bounced_analysis.xlsx")
        
        # Create visualizations
        print("\n📊 Creating visualizations...")
        
        # 1. Bar chart of top countries
        plt.figure(figsize=(14, 8))
        top_countries = country_counts.head(15)
        colors = sns.color_palette("viridis", len(top_countries))
        bars = plt.barh(range(len(top_countries)), top_countries.values, color=colors)
        plt.yticks(range(len(top_countries)), [str(c) if pd.notna(c) else 'Unknown' for c in top_countries.index])
        plt.xlabel('Number of Bounced Emails', fontsize=12, fontweight='bold')
        plt.ylabel('Country', fontsize=12, fontweight='bold')
        plt.title('Top 15 Countries with Email Bounced Activity', fontsize=14, fontweight='bold', pad=20)
        plt.gca().invert_yaxis()
        
        # Add value labels on bars
        for i, (bar, value) in enumerate(zip(bars, top_countries.values)):
            plt.text(value, i, f' {value:,}', va='center', fontsize=10)
        
        plt.tight_layout()
        plt.savefig('reports/bounced_by_country.png', dpi=300, bbox_inches='tight')
        plt.close()
        print("✓ Saved: reports/bounced_by_country.png")
        
        # 2. Pie chart for top countries
        plt.figure(figsize=(12, 8))
        top_10_countries = country_counts.head(10)
        others_count = country_counts[10:].sum()
        
        if others_count > 0:
            plot_data = pd.concat([top_10_countries, pd.Series({'Others': others_count})])
        else:
            plot_data = top_10_countries
        
        colors = sns.color_palette("Set3", len(plot_data))
        wedges, texts, autotexts = plt.pie(plot_data.values, labels=[str(c) if pd.notna(c) else 'Unknown' for c in plot_data.index],
                                           autopct='%1.1f%%', colors=colors, startangle=90)
        
        # Improve text
        for text in texts:
            text.set_fontsize(10)
            text.set_fontweight('bold')
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontsize(9)
            autotext.set_fontweight('bold')
        
        plt.title('Email Bounced Distribution by Country (Top 10)', fontsize=14, fontweight='bold', pad=20)
        plt.tight_layout()
        plt.savefig('reports/bounced_country_pie.png', dpi=300, bbox_inches='tight')
        plt.close()
        print("✓ Saved: reports/bounced_country_pie.png")
        
        # 3. Activity type distribution
        plt.figure(figsize=(12, 6))
        activity_plot = bounced_activities.head(10)
        colors = sns.color_palette("coolwarm", len(activity_plot))
        bars = plt.bar(range(len(activity_plot)), activity_plot.values, color=colors)
        plt.xticks(range(len(activity_plot)), [str(a)[:30] for a in activity_plot.index], rotation=45, ha='right')
        plt.ylabel('Count', fontsize=12, fontweight='bold')
        plt.xlabel('Activity Type', fontsize=12, fontweight='bold')
        plt.title('Email Bounced Activity Types', fontsize=14, fontweight='bold', pad=20)
        
        # Add value labels
        for bar, value in zip(bars, activity_plot.values):
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height,
                    f'{int(value):,}', ha='center', va='bottom', fontsize=9)
        
        plt.tight_layout()
        plt.savefig('reports/bounced_activity_types.png', dpi=300, bbox_inches='tight')
        plt.close()
        print("✓ Saved: reports/bounced_activity_types.png")
        
    else:
        print("\n⚠️  No records found with 'bounce' in Last Activity")
        print("\nLet me show you all unique Last Activity values to help identify the correct filter...")
        
else:
    print("\n✗ 'Last Activity' column not found in the dataset")
    print(f"\nAvailable columns: {', '.join(df.columns)}")

print("\n" + "=" * 70)
print("✓ Analysis Complete!")
print("=" * 70)
print("\nGenerated files:")
print("  - reports/email_bounced_analysis.xlsx")
print("  - reports/bounced_by_country.png")
print("  - reports/bounced_country_pie.png")
print("  - reports/bounced_activity_types.png")
print("=" * 70)
