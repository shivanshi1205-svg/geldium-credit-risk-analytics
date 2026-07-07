# train_models.py
# Python script to train and evaluate the Logistic Regression baseline and XGBoost models.

import os
import joblib
import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler, OneHotEncoder
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import GradientBoostingClassifier # Use scikit-learn's GBM as a robust alternative to XGBoost if xgboost library is not present, or try importing xgboost
from sklearn.metrics import classification_report, roc_auc_score, precision_recall_curve

# Set paths
base_dir = "C:/Users/USER/.gemini/antigravity/scratch/geldium_credit_risk"
csv_path = f"{base_dir}/geldium_customer_data_cleaned.csv"
model_dir = f"{base_dir}/models"
os.makedirs(model_dir, exist_ok=True)

# Load cleaned dataset
df = pd.read_csv(csv_path)
print(f"Loaded dataset: {df.shape[0]} rows, {df.shape[1]} columns")

# Define features and target variable
target = 'Delinquency_Flag'
numeric_features = ['Age', 'Annual_Income_Cleaned', 'Credit_Score_Cleaned', 'Credit_Limit', 'Credit_Utilization', 'Existing_Debt', 'Payment_History']
categorical_features = ['Employment_Status']

X = df[numeric_features + categorical_features]
y = df[target]

# Split data: 70% Train, 30% Test, stratified by delinquency flag
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.30, random_state=42, stratify=y)
print(f"Training set size: {X_train.shape[0]} rows")
print(f"Testing set size: {X_test.shape[0]} rows")

# Preprocessing pipeline
preprocessor = ColumnTransformer(
    transformers=[
        ('num', StandardScaler(), numeric_features),
        ('cat', OneHotEncoder(drop='first'), categorical_features)
    ])

# -------------------------------------------------------------
# 1. BASELINE MODEL: LOGISTIC REGRESSION
# -------------------------------------------------------------
print("\n--- Training Logistic Regression Baseline ---")
log_reg_pipeline = Pipeline(steps=[
    ('preprocessor', preprocessor),
    ('classifier', LogisticRegression(random_state=42, class_weight='balanced'))
])

log_reg_pipeline.fit(X_train, y_train)

# Evaluate Baseline
y_pred_lr = log_reg_pipeline.predict(X_test)
y_prob_lr = log_reg_pipeline.predict_proba(X_test)[:, 1]

print("Logistic Regression Metrics:")
print(classification_report(y_test, y_pred_lr))
print(f"ROC-AUC Score: {roc_auc_score(y_test, y_prob_lr):.4f}")

# -------------------------------------------------------------
# 2. PRODUCTION MODEL: XGBOOST / GRADIENT BOOSTING CHAMPION
# -------------------------------------------------------------
print("\n--- Training Production Champion (Gradient Boosting Classifier) ---")

# Try to import xgboost; if not installed, fall back to scikit-learn's GradientBoostingClassifier
try:
    import xgboost as xgb
    champion_classifier = xgb.XGBClassifier(
        random_state=42,
        use_label_encoder=False,
        eval_metric='logloss',
        scale_pos_weight=(len(y_train) - y_train.sum()) / y_train.sum() # handle class imbalance
    )
    print("Using XGBoost Classifier...")
except ImportError:
    champion_classifier = GradientBoostingClassifier(random_state=42, n_estimators=100, learning_rate=0.1)
    print("XGBoost library not found, using Scikit-Learn's GradientBoostingClassifier...")

champion_pipeline = Pipeline(steps=[
    ('preprocessor', preprocessor),
    ('classifier', champion_classifier)
])

champion_pipeline.fit(X_train, y_train)

# Evaluate Production Model
y_pred_ch = champion_pipeline.predict(X_test)
y_prob_ch = champion_pipeline.predict_proba(X_test)[:, 1]

print("Production Champion Metrics:")
print(classification_report(y_test, y_pred_ch))
print(f"ROC-AUC Score: {roc_auc_score(y_test, y_prob_ch):.4f}")

# -------------------------------------------------------------
# 3. OPTIMIZING THRESHOLD FOR RECALL (CREDIT RISK DECISIONING)
# -------------------------------------------------------------
print("\n--- Optimizing Decision Threshold for Default Detection (Recall Priority) ---")
# Get precision and recall curve values
precisions, recalls, thresholds = precision_recall_curve(y_test, y_prob_ch)

# Find the threshold that yields at least 85% Recall while maximizing Precision
desired_recall = 0.85
idx = np.where(recalls >= desired_recall)[0][-1]
opt_threshold = thresholds[idx]

print(f"Standard Decision Threshold: 0.50")
print(f"Optimized Decision Threshold (Recall >= {desired_recall:.0%}): {opt_threshold:.4f}")

# Apply optimized threshold
y_pred_opt = (y_prob_ch >= opt_threshold).astype(int)
print(f"\nClassification Report with Optimized Threshold ({opt_threshold:.2f}):")
print(classification_report(y_test, y_pred_opt))

# -------------------------------------------------------------
# 4. SAVE MODELS AND PREPROCESSORS TO DISK
# -------------------------------------------------------------
joblib.dump(log_reg_pipeline, f"{model_dir}/baseline_logistic_regression.joblib")
joblib.dump(champion_pipeline, f"{model_dir}/production_gradient_boosting.joblib")

# Save threshold metadata
threshold_df = pd.DataFrame({"optimal_threshold": [opt_threshold]})
threshold_df.to_csv(f"{model_dir}/model_metadata.csv", index=False)

print(f"\nModels successfully saved to {model_dir}")
