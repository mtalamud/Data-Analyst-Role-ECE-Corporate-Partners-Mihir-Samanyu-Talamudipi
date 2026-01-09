# =====================================================
# Corporate Partnerships Analysis - Visualization Script
# =====================================================

# --- 1. Setup & Imports ---
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
from IPython.display import display

# Ensure figures folder exists
os.makedirs("figures", exist_ok=True)

# Load Excel file
file_path = "Corporate Partners_Dataset.xlsx"

partners_df = pd.read_excel(file_path, sheet_name="partners_table")
tiers_df = pd.read_excel(file_path, sheet_name="tiers_table")

# Display previews
display(partners_df.head())
display(tiers_df.head())

# =====================================================
# --- 2. Visualization 1: Partner Count per University ---
# =====================================================

# Count partners by university
partner_count = partners_df.groupby("University")["Partner Company"].count().sort_values(ascending=False)

plt.figure(figsize=(12, 6))
sns.barplot(x=partner_count.values, y=partner_count.index, palette="crest")
plt.title("Number of Corporate Partners per University", fontsize=14, fontweight="bold")
plt.xlabel("Number of Partners")
plt.ylabel("University")
plt.tight_layout()
plt.savefig("figures/partner_count_per_university.png", dpi=300)
plt.show()

# =====================================================
# --- 3. Visualization 2: Distribution of Technology Sectors ---
# =====================================================

# Count sectors and group smaller ones
sector_count = partners_df["Technology Sector"].value_counts()
top_sectors = sector_count[:8]
other_total = sector_count[8:].sum()

# Combine into one series
sector_summary = top_sectors.copy()
sector_summary["Other"] = other_total

# Create donut chart with clean colors
colors = sns.color_palette("crest", len(sector_summary))

fig, ax = plt.subplots(figsize=(8, 6))
wedges, texts, autotexts = ax.pie(
    sector_summary,
    autopct=lambda p: f"{p:.1f}%" if p > 3 else "",
    startangle=140,
    colors=colors,
    wedgeprops={'linewidth': 1, 'edgecolor': 'white'},
    textprops={'fontsize': 10, 'color': 'black'}
)

# Draw center circle for donut effect
centre_circle = plt.Circle((0, 0), 0.70, fc="white")
fig.gca().add_artist(centre_circle)

ax.set_title("Top Partner Technology Sectors", fontsize=14, fontweight="bold")

# Add legend outside
plt.legend(
    sector_summary.index,
    title="Technology Sector",
    bbox_to_anchor=(1.05, 1),
    loc="upper left"
)

plt.tight_layout()
plt.savefig("figures/sector_distribution.png", dpi=300, bbox_inches="tight")
plt.show()

# =====================================================
# --- 4. Visualization 3: Tier Comparison — Purdue vs Stanford ---
# =====================================================

# --- Purdue vs Stanford: Highest Tier Annual Fees ---
tier_fee_data = {
    'University': ['Purdue', 'Stanford'],
    'Highest Tier': ['Gold', 'Explorer'],
    'Annual Fee (USD)': [25000, 35000]  # <-- replace with actual if available
}

tier_fee_df = pd.DataFrame(tier_fee_data)

# Create figure
plt.figure(figsize=(7, 5))
sns.barplot(
    data=tier_fee_df,
    x='University',
    y='Annual Fee (USD)',
    hue='Highest Tier',
    palette=["#000000", "#c2303f"]  # black for Purdue, deep red for Stanford
)

# Chart styling
plt.title("Highest Tier Annual Fee: Purdue vs Stanford", fontsize=14, fontweight='bold')
plt.ylabel("Annual Fee (USD)")
plt.xlabel("")
plt.grid(axis='y', linestyle='--', alpha=0.6)
plt.legend(title="Highest Tier")
plt.tight_layout()

# Save and show
plt.savefig("viz4_highest_tier_fee_comparison.png", bbox_inches='tight', dpi=300)
plt.show()
plt.close()

# =====================================================
# --- 5. Visualization 4: Heatmap of Partner Distribution ---
# =====================================================

# Create pivot table
heatmap_data = partners_df.pivot_table(
    index="University",
    columns="Technology Sector",
    values="Partner Company",
    aggfunc="count",
    fill_value=0
)

plt.figure(figsize=(14, 7))
sns.heatmap(heatmap_data, annot=True, fmt="d", cmap="Blues")
plt.title("Heatmap: Partner Distribution by University and Sector", fontsize=14, fontweight="bold")
plt.xlabel("Technology Sector")
plt.ylabel("University")
plt.tight_layout()
plt.savefig("figures/heatmap_partner_distribution.png", dpi=300)
plt.show()

# =====================================================
# --- 6. Summary ---
# =====================================================

print("""
Summary of Visual Insights:
1. The bar chart shows which universities maintain the most extensive corporate partnerships.
2. The pie chart reveals which technology sectors dominate across the top ECE programs.
3. The Purdue vs Stanford tier fee comparison provides insight into how two universities structure their partnership levels.
4. The heatmap highlights which universities engage most heavily with specific sectors.
""")
