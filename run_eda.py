# run_eda.py
# Python script to perform Exploratory Data Analysis and generate charts for Geldium.

import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Set paths
base_dir = "C:/Users/USER/.gemini/antigravity/scratch/geldium_credit_risk"
csv_path = f"{base_dir}/geldium_customer_data.csv"
plots_dir = f"{base_dir}/plots"

# Create plots folder if it doesn't exist
os.makedirs(plots_dir, exist_ok=True)

# Load data
df = pd.read_csv(csv_path)
print(f"Data loaded: {df.shape[0]} rows, {df.shape[1]} columns")

# Set visualization style
sns.set_theme(style="whitegrid")
plt.rcParams.update({
    'font.family': 'sans-serif',
    'font.sans-serif': ['DejaVu Sans', 'Arial', 'Helvetica'],
    'axes.edgecolor': '#cccccc',
    'axes.linewidth': 0.8,
    'xtick.color': '#333333',
    'ytick.color': '#333333',
    'text.color': '#333333',
    'axes.labelcolor': '#111111',
    'figure.titlesize': 14,
    'axes.titlesize': 12,
    'axes.labelsize': 10,
    'xtick.labelsize': 9,
    'ytick.labelsize': 9
})

# Define colors
c_navy = "#1B365D"
c_grey = "#5C768D"
c_gold = "#D99B00"
c_light_grey = "#EAEAEA"
palette_binary = {0: c_grey, 1: c_gold, "0": c_grey, "1": c_gold}

# --- 1. DATA CLEANING & PREPROCESSING ---
print("\n--- Running Data Cleaning & Preprocessing ---")

# A. Missing Value Handling
# Record missing counts
missing_income = df['Annual_Income'].isna().sum()
missing_score = df['Credit_Score'].isna().sum()
print(f"Missing values before cleaning: Annual_Income={missing_income}, Credit_Score={missing_score}")

# Impute with medians
income_median = df['Annual_Income'].median()
score_median = df['Credit_Score'].median()

df['Annual_Income_Cleaned'] = df['Annual_Income'].fillna(income_median)
df['Credit_Score_Cleaned'] = df['Credit_Score'].fillna(score_median)

# B. Outlier Treatment
# Outliers in income: manually injected extremely large values (e.g. > $500,000)
# Outliers in credit score: FICO score of 9999 (system anomaly)
print("Identifying outliers...")
income_outliers = df[df['Annual_Income_Cleaned'] > 500000]
score_outliers = df[df['Credit_Score_Cleaned'] > 850]

print(f"Found {len(income_outliers)} income outliers (> $500k) and {len(score_outliers)} credit score outliers (> 850)")

# Treat outliers: Cap income at the 99th percentile of non-outliers, and FICO score at 850 (or replace 9999 with median)
income_cap = df.loc[df['Annual_Income_Cleaned'] <= 500000, 'Annual_Income_Cleaned'].quantile(0.99)
df['Annual_Income_Cleaned'] = np.where(df['Annual_Income_Cleaned'] > 500000, income_cap, df['Annual_Income_Cleaned'])
df['Credit_Score_Cleaned'] = np.where(df['Credit_Score_Cleaned'] > 850, score_median, df['Credit_Score_Cleaned'])

print(f"Cleaned stats: Max Income = ${df['Annual_Income_Cleaned'].max():,.2f}, Max Credit Score = {df['Credit_Score_Cleaned'].max()}")

# --- 2. EXPLORATORY DATA ANALYSIS & CHART GENERATION ---
print("\n--- Generating Plots ---")

# Plot 1: Credit Score Distribution by Delinquency Status
plt.figure(figsize=(7, 4.5))
ax = sns.kdeplot(data=df, x='Credit_Score_Cleaned', hue='Delinquency_Flag', fill=True, common_norm=False, palette=palette_binary, alpha=0.4, linewidth=2)
plt.title("Credit Score (FICO) Distribution by Delinquency Status", fontsize=13, weight='bold', color=c_navy, pad=15)
plt.xlabel("FICO Credit Score", fontsize=10, labelpad=10)
plt.ylabel("Density", fontsize=10, labelpad=10)
plt.xlim(300, 850)
legend = ax.get_legend()
if legend:
    legend.set_title("Status")
    new_labels = ['Current (0)', 'Delinquent (1)']
    for t, l in zip(legend.texts, new_labels):
        t.set_text(l)
sns.despine()
plt.tight_layout()
plt.savefig(f"{plots_dir}/dist_credit_score.png", dpi=300)
plt.close()
print("Saved dist_credit_score.png")

# Plot 2: Credit Utilization by Delinquency Status
plt.figure(figsize=(7, 4.5))
ax = sns.boxplot(data=df, x='Delinquency_Flag', y='Credit_Utilization', palette=palette_binary, width=0.4, linewidth=1.5)
plt.title("Credit Utilization Ratio by Delinquency Status", fontsize=13, weight='bold', color=c_navy, pad=15)
plt.xlabel("Delinquency Flag (0 = Current, 1 = Delinquent)", fontsize=10, labelpad=10)
plt.ylabel("Credit Card Utilization Ratio", fontsize=10, labelpad=10)
plt.xticks([0, 1], ['Current (0)', 'Delinquent (1)'])
sns.despine()
plt.tight_layout()
plt.savefig(f"{plots_dir}/dist_utilization.png", dpi=300)
plt.close()
print("Saved dist_utilization.png")

# Plot 3: Late Payments vs Delinquency Rate
plt.figure(figsize=(7, 4.5))
# Calculate delinquency rate for each late payment count
delinq_by_late = df.groupby('Payment_History')['Delinquency_Flag'].agg(['mean', 'count']).reset_index()
# Filter to counts with at least 5 observations to avoid noise
delinq_by_late = delinq_by_late[delinq_by_late['count'] > 5]

ax = sns.barplot(data=delinq_by_late, x='Payment_History', y='mean', color=c_navy, alpha=0.85)
plt.title("Delinquency Rate by Number of Historic Late Payments", fontsize=13, weight='bold', color=c_navy, pad=15)
plt.xlabel("Late Payments (Past 12 Months)", fontsize=10, labelpad=10)
plt.ylabel("Delinquency Rate (Proportion)", fontsize=10, labelpad=10)
# Formatter for y-axis as percentage
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _: '{:.0%}'.format(y)))

# Annotate bars with values
for p in ax.patches:
    ax.annotate(f"{p.get_height():.1%}", (p.get_x() + p.get_width() / 2., p.get_height() + 0.02),
                ha='center', va='center', xytext=(0, 5), textcoords='offset points', fontsize=8.5, weight='bold', color='#444444')

plt.ylim(0, 1.1)
sns.despine()
plt.tight_layout()
plt.savefig(f"{plots_dir}/late_payments_by_delinquency.png", dpi=300)
plt.close()
print("Saved late_payments_by_delinquency.png")

# Plot 4: Correlation Matrix Heatmap
plt.figure(figsize=(7.5, 6.5))
# Keep numerical variables
num_cols = ['Age', 'Annual_Income_Cleaned', 'Credit_Score_Cleaned', 'Credit_Limit', 'Credit_Utilization', 'Existing_Debt', 'Payment_History', 'Delinquency_Flag']
corr_matrix = df[num_cols].corr()

# Rename columns for cleaner visualization
corr_matrix.columns = ['Age', 'Annual Income', 'FICO Score', 'Credit Limit', 'Credit Util.', 'Total Debt', 'Late Payments', 'Delinquency']
corr_matrix.index = corr_matrix.columns

mask = np.triu(np.ones_like(corr_matrix, dtype=bool))
cmap = sns.diverging_palette(230, 20, as_cmap=True)

sns.heatmap(corr_matrix, mask=mask, cmap=cmap, vmax=0.8, center=0, annot=True, fmt=".2f", square=True, linewidths=.5, cbar_kws={"shrink": .8})
plt.title("Correlation Matrix of Key Credit Risk Metrics", fontsize=13, weight='bold', color=c_navy, pad=20)
plt.tight_layout()
plt.savefig(f"{plots_dir}/correlation_matrix.png", dpi=300)
plt.close()
print("Saved correlation_matrix.png")

# Cleaned data export
cleaned_csv_path = f"{base_dir}/geldium_customer_data_cleaned.csv"
df.to_csv(cleaned_csv_path, index=False)
print(f"Cleaned dataset exported to {cleaned_csv_path}")

print("EDA and plot generation completed successfully.")
