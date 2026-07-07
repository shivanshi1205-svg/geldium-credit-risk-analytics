# generate_dataset.py
# Python script to generate a realistic credit risk dataset for Geldium.

import numpy as np
import pandas as pd

# Set random seed for reproducibility
np.random.seed(42)

n_samples = 2500

# 1. Base Customer ID
customer_ids = [f"GELD-{10000 + i}" for i in range(n_samples)]

# 2. Age (Normally distributed around 38, bounded between 18 and 70)
age = np.random.normal(loc=38, scale=12, size=n_samples).astype(int)
age = np.clip(age, 18, 70)

# 3. Employment Status
employment_options = ["Employed", "Self-Employed", "Unemployed", "Student"]
employment_weights = [0.70, 0.15, 0.10, 0.05]
employment_status = np.random.choice(employment_options, size=n_samples, p=employment_weights)

# 4. Annual Income (Log-normal distribution, skewed right)
# Median income around $55,000, with range $15,000 to $250,000
raw_income = np.random.lognormal(mean=10.9, sigma=0.45, size=n_samples)
# Adjust income based on age (younger/older have lower on average) and employment
age_factor = np.where(age < 25, 0.5, np.where(age > 60, 0.75, 1.0))
emp_factor = np.where(employment_status == "Unemployed", 0.15, 
                     np.where(employment_status == "Student", 0.20, 1.0))
annual_income = (raw_income * age_factor * emp_factor).astype(int)
annual_income = np.clip(annual_income, 12000, 280000)

# 5. Credit Score (FICO range 300 to 850)
# Base score depends on age, employment and income
base_score = 620 + (age - 35) * 1.5 + (annual_income - 60000) / 1000
base_score = np.where(employment_status == "Unemployed", base_score - 40, base_score)
base_score = np.where(employment_status == "Student", base_score - 15, base_score)
# Add significant variation/noise
credit_score = np.random.normal(loc=base_score, scale=50).astype(int)
credit_score = np.clip(credit_score, 300, 850)

# 6. Credit Limit (Correlated with income and credit score)
credit_limit = (annual_income * 0.15) + (credit_score - 500) * 18 + np.random.normal(0, 1500, n_samples)
credit_limit = np.clip(credit_limit, 500, 45000).astype(int)

# 7. Credit Utilization (0.0 to 1.0, negatively correlated with FICO score)
# Lower credit scores are forced to have higher credit utilization
util_base = 0.85 - (credit_score - 300) / 750
credit_utilization = np.clip(np.random.beta(a=2, b=5, size=n_samples) + util_base, 0.0, 1.15)
# If credit limit is very low, utilization is higher
credit_utilization = np.where(credit_limit < 2000, credit_utilization + 0.1, credit_utilization)
credit_utilization = np.clip(credit_utilization, 0.0, 1.1)

# 8. Existing Debt (Car loans, student loans, other credit, based on income)
dti_ratio = np.random.lognormal(mean=-1.6, sigma=0.4, size=n_samples) # DTI ratio mean around 20%
existing_debt = (annual_income * dti_ratio).astype(int)
existing_debt = np.where(employment_status == "Student", existing_debt + np.random.randint(15000, 45000, n_samples), existing_debt) # Student loans
existing_debt = np.clip(existing_debt, 0, 150000)

# 9. Payment History (Number of 30+ days late payments in last 12 months)
# Strongly correlated with low credit score and high utilization
late_prob_base = 0.05 + 0.4 * credit_utilization + (700 - credit_score) / 1000
late_prob_base = np.clip(late_prob_base, 0.01, 0.95)
payment_history = np.zeros(n_samples, dtype=int)
for i in range(n_samples):
    # Number of late payments drawn from a Poisson distribution
    payment_history[i] = np.random.poisson(lam=late_prob_base[i] * 2.5)
payment_history = np.clip(payment_history, 0, 8)

# 10. Delinquency Flag (Target Variable: 1 = Delinquent, 0 = Current)
# Delinquency probability modeled via a logistic regression formula:
# High utilization, late payments, low FICO, and high debt-to-income drive delinquency
dti = existing_debt / np.maximum(annual_income, 10000)
z = -6.2 + 4.2 * credit_utilization + 1.1 * payment_history - 0.006 * (credit_score - 600) + 0.5 * dti
prob_delinquency = 1 / (1 + np.exp(-z))
# Add random variation to delinquency flag to simulate noise and prevent perfect separation
delinquency_flag = (np.random.uniform(0, 1, n_samples) < prob_delinquency).astype(int)

# Create DataFrame
df = pd.DataFrame({
    "Customer_ID": customer_ids,
    "Age": age,
    "Employment_Status": employment_status,
    "Annual_Income": annual_income,
    "Credit_Score": credit_score,
    "Credit_Limit": credit_limit,
    "Credit_Utilization": credit_utilization,
    "Existing_Debt": existing_debt,
    "Payment_History": payment_history,
    "Delinquency_Flag": delinquency_flag
})

# Inject simulated missing values to make it realistic (e.g. 2% missing in Annual_Income, 2.5% in Credit_Score)
missing_income_idx = np.random.choice(n_samples, size=int(n_samples * 0.02), replace=False)
missing_score_idx = np.random.choice(n_samples, size=int(n_samples * 0.025), replace=False)

df.loc[missing_income_idx, "Annual_Income"] = np.nan
df.loc[missing_score_idx, "Credit_Score"] = np.nan

# Inject extreme outliers to make outlier handling realistic
# e.g., a few customers with 0 FICO score (system error) or $2M income (data entry error)
outlier_income_idx = np.random.choice(n_samples, size=5, replace=False)
df.loc[outlier_income_idx, "Annual_Income"] = df.loc[outlier_income_idx, "Annual_Income"] * 10

outlier_score_idx = np.random.choice(n_samples, size=3, replace=False)
df.loc[outlier_score_idx, "Credit_Score"] = 9999 # System anomaly score

# Export to CSV
csv_path = "C:/Users/USER/.gemini/antigravity/scratch/geldium_credit_risk/geldium_customer_data.csv"
df.to_csv(csv_path, index=False)

# Print statistics for verification
print(f"Dataset generated successfully with {df.shape[0]} rows and {df.shape[1]} columns.")
print(f"Delinquency rate: {df['Delinquency_Flag'].mean():.2%}")
print(f"Missing values: Income={df['Annual_Income'].isna().sum()}, Credit Score={df['Credit_Score'].isna().sum()}")
print(f"Delinquent count: {df['Delinquency_Flag'].sum()} / {len(df)}")
