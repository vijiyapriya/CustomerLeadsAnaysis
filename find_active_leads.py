"""
Active Leads Analysis - Comprehensive
Analyzes leads that are still in active stages (not Won, Lost, or Disqualified)
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

print("\n" + "=" * 70)
print("Lead Stage Distribution")
print("=" * 70)

# Show all lead stages
stage_counts = df['Lead Stage'].value_counts(dropna=False)
print(f"\n📊 All Lead Stages:")
print("-" * 70)
print(f"{'Lead Stage':<40} {'Count':>12} {'Percentage':>12}")
print("-" * 70)
for stage, count in stage_counts.items():
    pct = (count / len(df)) * 100
    stage_name = str(stage) if pd.notna(stage) else "Missing/Unknown"
    print(f"{stage_name:<40} {count:>12,} {pct:>11.2f}%")

# Define inactive stages
inactive_stages = ['Disqualified', 'Lost', 'Won', 'Closure - Customer', 'Closure']

print("\n" + "=" * 70)
print("Defining Active Leads")
print("=" * 70)
print(f"\nInactive stages (excluded): {', '.join(inactive_stages)}")
print("Active leads = All other stages (in sales pipeline)")

# Filter for active leads (not in inactive stages)
active_leads = df[~df['Lead Stage'].isin(inactive_stages)].copy()

print(f"\n✓ Found {len(active_leads):,} Active leads ({(len(active_leads)/len(df)*100):.2f}% of total)")

# Show active lead stages breakdown
print("\n📋 Active Lead Stages Breakdown:")
print("-" * 70)
active_stage_counts = active_leads['Lead Stage'].value_counts(dropna=False)
print(f"{'Lead Stage':<40} {'Count':>12} {'Percentage':>12}")
print("-" * 70)
for stage, count in active_stage_counts.items():
    pct = (count / len(active_leads)) * 100
    stage_name = str(stage) if pd.notna(stage) else "Missing/Unknown"
    print(f"{stage_name:<40} {count:>12,} {pct:>11.2f}%")

# Analyze Active Leads by Country
print("\n" + "=" * 70)
print("Active Leads by Country")
print("=" * 70)

country_counts = active_leads['Country'].value_counts(dropna=False)

print(f"\nTotal active leads: {len(active_leads):,}")
print(f"Countries represented: {active_leads['Country'].nunique()}")
print(f"Records with missing country: {active_leads['Country'].isna().sum():,}")

print("\n📍 Top 25 Countries with Active Leads:")
print("-" * 70)
print(f"{'Country':<35} {'Count':>10} {'Percentage':>12}")
print("-" * 70)

for country, count in country_counts.head(25).items():
    pct = (count / len(active_leads)) * 100
    country_name = str(country) if pd.notna(country) else "Missing/Unknown"
    print(f"{country_name:<35} {count:>10,} {pct:>11.2f}%")

# Analyze Active Leads by Industry
if 'Industry Vertical' in active_leads.columns:
    print("\n" + "=" * 70)
    print("Active Leads by Industry Vertical")
    print("=" * 70)
    
    industry_counts = active_leads['Industry Vertical'].value_counts(dropna=False)
    
    print(f"\n📊 Top 20 Industries with Active Leads:")
    print("-" * 70)
    print(f"{'Industry Vertical':<40} {'Count':>10} {'%':>8}")
    print("-" * 70)
    
    for industry, count in industry_counts.head(20).items():
        pct = (count / len(active_leads)) * 100
        industry_name = str(industry)[:38] if pd.notna(industry) else "Missing"
        print(f"{industry_name:<40} {count:>10,} {pct:>7.2f}%")

# Analyze by Lead Source
if 'Lead Source' in active_leads.columns:
    print("\n" + "=" * 70)
    print("Active Leads by Lead Source")
    print("=" * 70)
    
    source_counts = active_leads['Lead Source'].value_counts(dropna=False)
    
    print(f"\n📌 Lead Sources for Active Leads:")
    print("-" * 70)
    print(f"{'Lead Source':<40} {'Count':>10} {'%':>8}")
    print("-" * 70)
    
    for source, count in source_counts.head(15).items():
        pct = (count / len(active_leads)) * 100
        source_name = str(source)[:38] if pd.notna(source) else "Missing"
        print(f"{source_name:<40} {count:>10,} {pct:>7.2f}%")

# Analyze by Company Size
if 'Company size' in active_leads.columns:
    print("\n" + "=" * 70)
    print("Active Leads by Company Size")
    print("=" * 70)
    
    size_counts = active_leads['Company size'].value_counts(dropna=False)
    
    print(f"\n🏢 Company Size Distribution:")
    print("-" * 70)
    print(f"{'Company Size':<40} {'Count':>10} {'%':>8}")
    print("-" * 70)
    
    for size, count in size_counts.items():
        pct = (count / len(active_leads)) * 100
        size_name = str(size)[:38] if pd.notna(size) else "Missing"
        print(f"{size_name:<40} {count:>10,} {pct:>7.2f}%")

# Analyze by Last Activity
if 'Last Activity' in active_leads.columns:
    print("\n" + "=" * 70)
    print("Active Leads - Last Activity")
    print("=" * 70)
    
    activity_counts = active_leads['Last Activity'].value_counts(dropna=False)
    
    print(f"\n🔔 Top 15 Last Activities for Active Leads:")
    print("-" * 70)
    print(f"{'Last Activity':<40} {'Count':>10} {'%':>8}")
    print("-" * 70)
    
    for activity, count in activity_counts.head(15).items():
        pct = (count / len(active_leads)) * 100
        activity_name = str(activity)[:38] if pd.notna(activity) else "Missing"
        print(f"{activity_name:<40} {count:>10,} {pct:>7.2f}%")

# Analyze by Region
if 'Region Specific' in active_leads.columns:
    print("\n" + "=" * 70)
    print("Active Leads by Region")
    print("=" * 70)
    
    region_counts = active_leads['Region Specific'].value_counts(dropna=False)
    
    print(f"\n🌍 Regional Distribution:")
    print("-" * 70)
    print(f"{'Region':<40} {'Count':>10} {'%':>8}")
    print("-" * 70)
    
    for region, count in region_counts.items():
        pct = (count / len(active_leads)) * 100
        region_name = str(region)[:38] if pd.notna(region) else "Missing"
        print(f"{region_name:<40} {count:>10,} {pct:>7.2f}%")

# Export to Excel
print("\n" + "=" * 70)
print("Exporting Results")
print("=" * 70)

with pd.ExcelWriter('reports/active_leads_comprehensive.xlsx', engine='openpyxl') as writer:
    # Sheet 1: All Active Leads
    active_leads.to_excel(writer, sheet_name='Active Leads', index=False)
    
    # Sheet 2: Summary Statistics
    summary_stats = pd.DataFrame({
        'Metric': [
            'Total Records in Dataset',
            'Total Active Leads',
            'Active % of Total',
            'Inactive Leads (Won/Lost/Disqualified)',
            'Countries Represented',
            'Industries Represented',
            'Lead Sources',
            'Leads with Email',
            'Leads with Phone/Mobile'
        ],
        'Value': [
            f"{len(df):,}",
            f"{len(active_leads):,}",
            f"{(len(active_leads) / len(df) * 100):.2f}%",
            f"{len(df) - len(active_leads):,}",
            f"{active_leads['Country'].nunique():,}",
            f"{active_leads['Industry Vertical'].nunique():,}" if 'Industry Vertical' in active_leads.columns else 'N/A',
            f"{active_leads['Lead Source'].nunique():,}" if 'Lead Source' in active_leads.columns else 'N/A',
            f"{active_leads['Email'].notna().sum():,}" if 'Email' in active_leads.columns else 'N/A',
            f"{(active_leads['Phone Number'].notna() | active_leads['Mobile Number'].notna()).sum():,}" if 'Phone Number' in active_leads.columns else 'N/A'
        ]
    })
    summary_stats.to_excel(writer, sheet_name='Summary', index=False)
    
    # Sheet 3: Lead Stage Breakdown
    stage_summary = pd.DataFrame({
        'Lead Stage': active_stage_counts.index,
        'Count': active_stage_counts.values,
        'Percentage': (active_stage_counts.values / len(active_leads) * 100).round(2)
    })
    stage_summary.to_excel(writer, sheet_name='By Lead Stage', index=False)
    
    # Sheet 4: By Country
    country_summary = pd.DataFrame({
        'Country': country_counts.index,
        'Count': country_counts.values,
        'Percentage': (country_counts.values / len(active_leads) * 100).round(2)
    })
    country_summary.to_excel(writer, sheet_name='By Country', index=False)
    
    # Sheet 5: By Industry
    if 'Industry Vertical' in active_leads.columns:
        industry_summary = pd.DataFrame({
            'Industry Vertical': industry_counts.index,
            'Count': industry_counts.values,
            'Percentage': (industry_counts.values / len(active_leads) * 100).round(2)
        })
        industry_summary.to_excel(writer, sheet_name='By Industry', index=False)
    
    # Sheet 6: By Lead Source
    if 'Lead Source' in active_leads.columns:
        source_summary = pd.DataFrame({
            'Lead Source': source_counts.index,
            'Count': source_counts.values,
            'Percentage': (source_counts.values / len(active_leads) * 100).round(2)
        })
        source_summary.to_excel(writer, sheet_name='By Lead Source', index=False)
    
    # Sheet 7: By Company Size
    if 'Company size' in active_leads.columns:
        size_summary = pd.DataFrame({
            'Company Size': size_counts.index,
            'Count': size_counts.values,
            'Percentage': (size_counts.values / len(active_leads) * 100).round(2)
        })
        size_summary.to_excel(writer, sheet_name='By Company Size', index=False)
    
    # Sheet 8: By Region
    if 'Region Specific' in active_leads.columns:
        region_summary = pd.DataFrame({
            'Region': region_counts.index,
            'Count': region_counts.values,
            'Percentage': (region_counts.values / len(active_leads) * 100).round(2)
        })
        region_summary.to_excel(writer, sheet_name='By Region', index=False)

print("✓ Exported to: reports/active_leads_comprehensive.xlsx")

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
plt.title(f'Top 15 Countries - Active Leads (Total: {len(active_leads):,})', fontsize=14, fontweight='bold', pad=20)
plt.gca().invert_yaxis()

for i, (bar, value) in enumerate(zip(bars, top_countries.values)):
    plt.text(value, i, f' {value:,}', va='center', fontsize=10)

plt.tight_layout()
plt.savefig('reports/active_leads_by_country.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: reports/active_leads_by_country.png")

# 2. Active Leads Stage Distribution
plt.figure(figsize=(12, 8))
colors = sns.color_palette("Set2", len(active_stage_counts))
wedges, texts, autotexts = plt.pie(active_stage_counts.values, 
                                   labels=[str(s)[:30] if pd.notna(s) else 'Unknown' for s in active_stage_counts.index],
                                   autopct='%1.1f%%', colors=colors, startangle=90)

for text in texts:
    text.set_fontsize(9)
    text.set_fontweight('bold')
for autotext in autotexts:
    autotext.set_color('white')
    autotext.set_fontsize(9)
    autotext.set_fontweight('bold')

plt.title(f'Active Leads by Stage (Total: {len(active_leads):,})', fontsize=14, fontweight='bold', pad=20)
plt.tight_layout()
plt.savefig('reports/active_leads_by_stage.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: reports/active_leads_by_stage.png")

# 3. Active Leads by Industry (if available)
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

# 4. Company Size Distribution
if 'Company size' in active_leads.columns and len(size_counts) > 0:
    plt.figure(figsize=(12, 7))
    colors = sns.color_palette("Spectral", len(size_counts))
    wedges, texts, autotexts = plt.pie(size_counts.values,
                                       labels=[str(s)[:25] if pd.notna(s) else 'Unknown' for s in size_counts.index],
                                       autopct='%1.1f%%', colors=colors, startangle=45)
    
    for text in texts:
        text.set_fontsize(9)
        text.set_fontweight('bold')
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontsize(9)
        autotext.set_fontweight('bold')
    
    plt.title('Active Leads by Company Size', fontsize=14, fontweight='bold', pad=20)
    plt.tight_layout()
    plt.savefig('reports/active_leads_by_company_size.png', dpi=300, bbox_inches='tight')
    plt.close()
    print("✓ Saved: reports/active_leads_by_company_size.png")

print("\n" + "=" * 70)
print("✓ Analysis Complete!")
print("=" * 70)
print(f"\n📊 Key Insights:")
print(f"  • Total Active Leads: {len(active_leads):,}")
print(f"  • Active Rate: {(len(active_leads)/len(df)*100):.2f}%")
print(f"  • Top Country: {country_counts.index[0]} ({country_counts.values[0]:,} leads)")
if 'Industry Vertical' in active_leads.columns:
    print(f"  • Top Industry: {industry_counts.index[0]} ({industry_counts.values[0]:,} leads)")

print("\n📁 Generated files:")
print("  - reports/active_leads_comprehensive.xlsx (8 sheets)")
print("  - reports/active_leads_by_country.png")
print("  - reports/active_leads_by_stage.png")
print("  - reports/active_leads_by_industry.png")
print("  - reports/active_leads_by_company_size.png")
print("=" * 70)
