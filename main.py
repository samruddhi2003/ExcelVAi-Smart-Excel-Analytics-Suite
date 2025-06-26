import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

# Paths
INPUT_FILE = "data/sample_data.xlsx"
OUTPUT_DIR = "outputs"
CLEANED_FILE = os.path.join(OUTPUT_DIR, "cleaned_data.xlsx")

# Create outputs directory if it doesn't exist
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Load Excel File
print("🔄 Loading Excel file...")
try:
    df = pd.read_excel(INPUT_FILE)
    print("✅ File loaded successfully!")
except Exception as e:
    print(f"❌ Error loading file: {e}")
    exit()

# Display basic info
print("\n📊 Data Summary:")
print(df.info())

# Drop rows with all NaNs
df.dropna(how='all', inplace=True)

# Fill missing numeric columns with median
numeric_cols = df.select_dtypes(include=['number']).columns
df[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].median())

# Fill missing object columns with mode
for col in df.select_dtypes(include=['object']).columns:
    if df[col].isnull().sum() > 0:
        df[col] = df[col].fillna(df[col].mode()[0])

# Summary Statistics
summary = df.describe(include='all')
print("\n📋 Summary Statistics:")
print(summary)

# Save summary as Excel
summary.to_excel(os.path.join(OUTPUT_DIR, "summary_report.xlsx"))

# Visualization: Correlation Heatmap
plt.figure(figsize=(10, 6))
sns.heatmap(df.corr(numeric_only=True), annot=True, cmap="coolwarm")
plt.title("Correlation Heatmap")
plt.tight_layout()
plt.savefig(os.path.join(OUTPUT_DIR, "correlation_heatmap.png"))
plt.close()

# Save cleaned data
df.to_excel(CLEANED_FILE, index=False)
print(f"\n✅ Cleaned data and summary exported to '{OUTPUT_DIR}' folder.")
