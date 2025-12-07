"""
Update Region Specific to ME for Middle Eastern Countries
Updates Saudi Arabia, Bahrain, Oman, Kuwait, UAE, and Qatar to ME region
"""

import pandas as pd
from datetime import datetime

# Load the Excel file
file_path = r"C:\Users\karul\Downloads\Raw File-LS-Full Data.xlsx"

print("=" * 70)
print("Updating Region Specific for Middle Eastern Countries")
print("=" * 70)

# Load data
print("\nLoading data...")
df = pd.read_excel(file_path)
print(f"✓ Loaded {len(df):,} rows and {len(df.columns)} columns")

# Define ME countries
me_countries = [
    'Saudi Arabia',
    'Bahrain',
    'Oman',
    'Kuwait',
    'United Arab Emirates',
    'Qatar'
]

print("\n" + "=" * 70)
print("Middle Eastern Countries to Update")
print("=" * 70)
print("\nCountries that should have Region Specific = 'ME':")
for country in me_countries:
    print(f"  • {country}")

# Check current state
print("\n" + "=" * 70)
print("Current State Analysis")
print("=" * 70)

# Count records for each ME country
print("\n📊 Current records for ME countries:")
print("-" * 70)
print(f"{'Country':<30} {'Total':>10} {'Already ME':>12} {'To Update':>12}")
print("-" * 70)

total_me_records = 0
total_already_me = 0
total_to_update = 0

for country in me_countries:
    country_mask = df['Country'] == country
    country_count = country_mask.sum()
    
    already_me = ((df['Country'] == country) & (df['Region Specific'] == 'ME')).sum()
    to_update = country_count - already_me
    
    total_me_records += country_count
    total_already_me += already_me
    total_to_update += to_update
    
    print(f"{country:<30} {country_count:>10,} {already_me:>12,} {to_update:>12,}")

print("-" * 70)
print(f"{'TOTAL':<30} {total_me_records:>10,} {total_already_me:>12,} {total_to_update:>12,}")

# Update the Region Specific field
print("\n" + "=" * 70)
print("Updating Records")
print("=" * 70)

# Create a backup column to track changes
df['Region Specific (Before)'] = df['Region Specific'].copy()

# Update Region Specific for ME countries
for country in me_countries:
    mask = df['Country'] == country
    df.loc[mask, 'Region Specific'] = 'ME'

# Verify updates
print("\n✓ Updates applied!")

# Show updated counts
print("\n📊 Verification - Region Specific after update:")
print("-" * 70)
print(f"{'Country':<30} {'Total':>10} {'Now ME':>12} {'Success':>10}")
print("-" * 70)

for country in me_countries:
    country_count = (df['Country'] == country).sum()
    now_me = ((df['Country'] == country) & (df['Region Specific'] == 'ME')).sum()
    success = '✓' if country_count == now_me else '✗'
    print(f"{country:<30} {country_count:>10,} {now_me:>12,} {success:>10}")

# Show overall Region Specific distribution
print("\n" + "=" * 70)
print("Updated Region Specific Distribution")
print("=" * 70)

region_counts = df['Region Specific'].value_counts(dropna=False)
print(f"\n📍 All Regions:")
print("-" * 70)
print(f"{'Region':<30} {'Count':>12} {'Percentage':>12}")
print("-" * 70)

for region, count in region_counts.items():
    pct = (count / len(df)) * 100
    region_name = str(region) if pd.notna(region) else "Missing/Unknown"
    print(f"{region_name:<30} {count:>12,} {pct:>11.2f}%")

# Show what changed
print("\n" + "=" * 70)
print("Changes Summary")
print("=" * 70)

changes_mask = df['Region Specific (Before)'] != df['Region Specific']
changes_df = df[changes_mask][['Country', 'Region Specific (Before)', 'Region Specific']].copy()

print(f"\nTotal records changed: {len(changes_df):,}")

if len(changes_df) > 0:
    print("\n📝 Changes by previous region:")
    print("-" * 70)
    change_summary = changes_df.groupby('Region Specific (Before)').size().sort_values(ascending=False)
    
    for old_region, count in change_summary.items():
        old_region_name = str(old_region) if pd.notna(old_region) else "Missing/Unknown"
        print(f"  {old_region_name:<30} → ME: {count:>8,} records")

# Export updated data
print("\n" + "=" * 70)
print("Exporting Updated Data")
print("=" * 70)

# Remove the backup column before exporting
df_export = df.drop(columns=['Region Specific (Before)'])

# Export to new file
output_file = 'reports/Raw_File_LS_Updated_Regions.xlsx'
df_export.to_excel(output_file, index=False, engine='openpyxl')
print(f"✓ Exported updated data to: {output_file}")

# Export change log
with pd.ExcelWriter('reports/region_update_log.xlsx', engine='openpyxl') as writer:
    # Sheet 1: Summary
    summary_data = {
        'Metric': [
            'Total Records in Dataset',
            'ME Country Records',
            'Records Updated',
            'Records Already Correct',
            'Update Date'
        ],
        'Value': [
            f"{len(df):,}",
            f"{total_me_records:,}",
            f"{total_to_update:,}",
            f"{total_already_me:,}",
            datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        ]
    }
    pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
    
    # Sheet 2: Country Breakdown
    country_breakdown = []
    for country in me_countries:
        country_count = (df_export['Country'] == country).sum()
        now_me = ((df_export['Country'] == country) & (df_export['Region Specific'] == 'ME')).sum()
        country_breakdown.append({
            'Country': country,
            'Total Records': country_count,
            'Region = ME': now_me,
            'Success': '✓' if country_count == now_me else '✗'
        })
    pd.DataFrame(country_breakdown).to_excel(writer, sheet_name='Country Breakdown', index=False)
    
    # Sheet 3: Changes Detail
    if len(changes_df) > 0:
        changes_df.to_excel(writer, sheet_name='Changes Detail', index=False)
    
    # Sheet 4: Region Distribution
    region_dist = pd.DataFrame({
        'Region': region_counts.index,
        'Count': region_counts.values,
        'Percentage': (region_counts.values / len(df) * 100).round(2)
    })
    region_dist.to_excel(writer, sheet_name='Region Distribution', index=False)

print(f"✓ Exported change log to: reports/region_update_log.xlsx")

print("\n" + "=" * 70)
print("✓ Update Complete!")
print("=" * 70)

print("\n📊 Summary:")
print(f"  • Total ME country records: {total_me_records:,}")
print(f"  • Records updated: {total_to_update:,}")
print(f"  • Records already correct: {total_already_me:,}")
print(f"  • ME region total: {region_counts.get('ME', 0):,}")

print("\n📁 Generated files:")
print("  - reports/Raw_File_LS_Updated_Regions.xlsx (Updated dataset)")
print("  - reports/region_update_log.xlsx (Change log)")
print("=" * 70)
