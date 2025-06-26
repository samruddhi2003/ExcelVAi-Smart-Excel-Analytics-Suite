import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
from fpdf import FPDF
import re

# 📂 Ensure output directory exists
output_dir = "output"
os.makedirs(output_dir, exist_ok=True)

# Step 5: Load Excel file
file_path = 'data/sample_data.xlsx'
df = pd.read_excel(file_path)

# Step 5a: Preview
print("📄 First 5 Rows of Data:")
print(df.head())

print("\nℹ️ Dataset Info:")
print(df.info())

# ✅ Step 6: Clean the Data
print("\n🧼 Cleaning Data...")
df_cleaned = df.dropna().drop_duplicates()
print(f"✅ Data cleaned. Rows after cleaning: {len(df_cleaned)}")

# ✅ Step 8: Summarize the cleaned data
print("\n📊 Summary Statistics:")
print(df_cleaned.describe())

print("\n🔢 Column-wise Null Value Count:")
print(df_cleaned.isnull().sum())

print("\n📐 Data Types:")
print(df_cleaned.dtypes)

sns.set_theme(style="whitegrid")

# ✅ Visualizations
# Replace placeholders with actual column names in your Excel
if 'CategoryColumn' in df_cleaned.columns:
    plt.figure(figsize=(8, 5))
    sns.countplot(data=df_cleaned, x='CategoryColumn')
    plt.title("Count of Categories")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(f"{output_dir}/category_count.png")
    plt.show()

if 'NumericColumn' in df_cleaned.columns:
    plt.figure(figsize=(8, 5))
    sns.histplot(df_cleaned['NumericColumn'], kde=True)
    plt.title("Distribution of NumericColumn")
    plt.tight_layout()
    plt.savefig(f"{output_dir}/numeric_distribution.png")
    plt.show()

# ✅ Correlation heatmap
plt.figure(figsize=(10, 7))
sns.heatmap(df_cleaned.corr(numeric_only=True), annot=True, cmap='coolwarm', fmt=".2f")
plt.title("Correlation Matrix")
plt.tight_layout()
plt.savefig(f"{output_dir}/correlation_matrix.png")
plt.show()

# ✅ Export Cleaned Data
cleaned_file_path = f"{output_dir}/cleaned_data.xlsx"
df_cleaned.to_excel(cleaned_file_path, index=False)
print(f"📤 Cleaned data exported to: {cleaned_file_path}")

# ✅ Create Summary Report
summary_report = f"""
Data Summary Report
-----------------------
Total Rows (Original): {len(df)}
Total Rows (Cleaned): {len(df_cleaned)}

Columns:
{', '.join(df.columns)}

Missing Values per Column:
{df.isnull().sum().to_string()}

Basic Stats:
{df.describe().to_string()}
"""

# Clean summary of emojis or unsupported characters for PDF
def remove_non_latin1(text):
    return re.sub(r'[^\x00-\xFF]', '', text)

# ✅ Export Summary as TXT
summary_path = f"{output_dir}/summary_report.txt"
with open(summary_path, "w", encoding="utf-8") as f:
    f.write(summary_report)

print(f"📄 Summary TXT exported to: {summary_path}")

# ✅ Export as PDF
class PDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 14)
        self.cell(0, 10, "Excel Data Summary Report", ln=True, align="C")

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Page {self.page_no()}", align="C")

    def chapter_body(self, text):
        self.set_font("Arial", "", 12)
        self.multi_cell(0, 10, text)

pdf = PDF()
pdf.add_page()

with open(summary_path, "r", encoding="utf-8") as f:
    cleaned_text = remove_non_latin1(f.read())  # Remove unsupported characters
    pdf.chapter_body(cleaned_text)

pdf_path = os.path.join(output_dir, "summary_report.pdf")
pdf.output(pdf_path)

# Optional: Open PDF automatically (works on Windows only)
try:
    os.startfile(pdf_path)
except AttributeError:
    print(f"📄 PDF saved at: {pdf_path} (open manually)")

print(f"📄 Summary PDF exported to: {pdf_path}")
