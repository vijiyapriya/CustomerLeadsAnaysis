"""
Regenerate Active Leads Pie Chart with Better Clarity
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load data
file_path = r"reports/Raw_File_LS_Updated_Regions_Final.xlsx"
df = pd.read_excel(file_path)

# Filter for active leads
inactive_stages = ['Disqualified', 'Lost', 'Won', 'Closure - Customer', 'Closure']
active_leads = df[~df['Lead Stage'].isin(inactive_stages)].copy()

# Get stage counts
stage_counts = active_leads['Lead Stage'].value_counts()

# Create improved pie chart
plt.figure(figsize=(14, 10))
colors = sns.color_palette("Set2", len(stage_counts))

# Create pie chart with improved settings
wedges, texts, autotexts = plt.pie(stage_counts.values,
                                    labels=[str(s)[:30] if pd.notna(s) else 'Unknown' for s in stage_counts.index],
                                    autopct='%1.1f%%',
                                    colors=colors,
                                    startangle=90,
                                    textprops={'fontsize': 14, 'weight': 'bold'},
                                    pctdistance=0.85)

# Improve text visibility
for text in texts:
    text.set_fontsize(16)
    text.set_fontweight('bold')
    text.set_color('black')

for autotext in autotexts:
    autotext.set_color('white')
    autotext.set_fontsize(14)
    autotext.set_fontweight('bold')

plt.title(f'Active Leads by Stage (Total: {len(active_leads):,})', 
          fontsize=20, fontweight='bold', pad=30, color='#1f4e79')

plt.tight_layout()
plt.savefig('reports/active_leads_by_stage.png', dpi=300, bbox_inches='tight', facecolor='white')
plt.close()

print("✓ Regenerated active_leads_by_stage.png with improved clarity")

# Also create a horizontal bar chart alternative
plt.figure(figsize=(12, 8))
stage_df = stage_counts.reset_index()
stage_df.columns = ['Stage', 'Count']
stage_df['Percentage'] = (stage_df['Count'] / len(active_leads) * 100).round(1)

colors_bar = sns.color_palette("viridis", len(stage_df))
bars = plt.barh(range(len(stage_df)), stage_df['Count'], color=colors_bar)
plt.yticks(range(len(stage_df)), stage_df['Stage'])
plt.xlabel('Number of Leads', fontsize=14, fontweight='bold')
plt.ylabel('Lead Stage', fontsize=14, fontweight='bold')
plt.title(f'Active Leads by Stage (Total: {len(active_leads):,})', 
          fontsize=16, fontweight='bold', pad=20)
plt.gca().invert_yaxis()

# Add value labels
for i, (bar, row) in enumerate(zip(bars, stage_df.itertuples())):
    plt.text(row.Count, i, f' {row.Count:,} ({row.Percentage}%)', 
             va='center', fontsize=11, fontweight='bold')

plt.tight_layout()
plt.savefig('reports/active_leads_by_stage_bar.png', dpi=300, bbox_inches='tight')
plt.close()

print("✓ Created alternative bar chart: active_leads_by_stage_bar.png")
