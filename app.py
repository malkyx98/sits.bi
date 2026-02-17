# ==============================================
# 🌐 AI-Powered Universal Excel Analyzer Dashboard
# ==============================================

# Install required libraries
!pip install pandas numpy matplotlib seaborn plotly openpyxl xlrd --quiet

# -----------------------------
# IMPORTS
# -----------------------------
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
from IPython.display import display, HTML
from google.colab import files

# -----------------------------
# FILE UPLOAD
# -----------------------------
uploaded = files.upload()
file_name = list(uploaded.keys())[0]
df = pd.read_excel(file_name)
print(f"✅ Loaded file: {file_name} | Rows: {df.shape[0]} | Columns: {df.shape[1]}")

# -----------------------------
# AUTOMATIC COLUMN DETECTION
# -----------------------------
def detect_columns(df):
    numeric, categorical, date, id_cols, status = [], [], [], [], []
    for col in df.columns:
        s = df[col].dropna()
        if pd.api.types.is_numeric_dtype(s):
            numeric.append(col)
        elif pd.api.types.is_datetime64_any_dtype(s):
            date.append(col)
        elif any(k in col.lower() for k in ['date','time']):
            try:
                df[col] = pd.to_datetime(df[col], errors='coerce')
                date.append(col)
            except:
                categorical.append(col)
        elif any(k in col.lower() for k in ['id','name','ref','user']):
            id_cols.append(col)
        elif any(k in col.lower() for k in ['status','sla','done','closed','pending']):
            status.append(col)
        else:
            categorical.append(col)
    return numeric, categorical, date, id_cols, status

numeric_cols, categorical_cols, date_cols, id_cols, status_cols = detect_columns(df)

# -----------------------------
# BEAUTIFUL DASHBOARD HEADER
# -----------------------------
display(HTML("<h1 style='color:#00BFFF;text-align:center;'>🤖 AI-Powered Universal Excel Analyzer</h1>"))
display(HTML(f"<h4 style='color:#FFD700;text-align:center;'>Loaded file: {file_name} | Rows: {df.shape[0]} | Columns: {df.shape[1]}</h4>"))

# -----------------------------
# DISPLAY COLUMN TYPES
# -----------------------------
display(HTML("<h3 style='color:#32CD32;'>🔍 Detected Column Roles:</h3>"))
display(pd.DataFrame({
    "Numeric Columns": [numeric_cols],
    "Categorical Columns": [categorical_cols],
    "Date Columns": [date_cols],
    "ID/Name Columns": [id_cols],
    "Status Columns": [status_cols]
}))

# -----------------------------
# NUMERIC KPIs
# -----------------------------
if numeric_cols:
    display(HTML("<h3 style='color:#FF69B4;'>📊 Numeric KPIs:</h3>"))
    kpi_table = pd.DataFrame(columns=["Column","Total","Mean","Max","Min","Std"])
    for col in numeric_cols:
        kpi_table = pd.concat([kpi_table, pd.DataFrame({
            "Column":[col],
            "Total":[df[col].sum()],
            "Mean":[df[col].mean()],
            "Max":[df[col].max()],
            "Min":[df[col].min()],
            "Std":[df[col].std()]
        })])
    display(kpi_table)
else:
    display(HTML("<h4 style='color:#FF4500;'>No numeric columns detected.</h4>"))

# -----------------------------
# CORRELATION HEATMAP
# -----------------------------
if len(numeric_cols) >= 2:
    display(HTML("<h3 style='color:#1E90FF;'>🧊 Numeric Correlation Heatmap:</h3>"))
    corr = df[numeric_cols].corr()
    plt.figure(figsize=(10,6))
    sns.heatmap(corr, annot=True, fmt=".2f", cmap='coolwarm', cbar_kws={'label': 'Correlation'})
    plt.show()

# -----------------------------
# CATEGORICAL ANALYSIS
# -----------------------------
if categorical_cols:
    display(HTML("<h3 style='color:#8A2BE2;'>🏷️ Categorical Analysis:</h3>"))
    for col in categorical_cols:
        counts = df[col].value_counts().head(10)
        display(HTML(f"<b>Top 10 values for '{col}':</b>"))
        display(counts)
        fig = px.bar(counts, x=counts.index, y=counts.values, 
                     labels={'x':col,'y':'Count'},
                     title=f"Top 10 {col} Values",
                     color=counts.values, color_continuous_scale='Viridis')
        fig.show()

# -----------------------------
# DATE TREND ANALYSIS
# -----------------------------
if date_cols:
    display(HTML("<h3 style='color:#FF8C00;'>📆 Date Trend Analysis:</h3>"))
    for col in date_cols:
        df[col+'_dt'] = pd.to_datetime(df[col], errors='coerce')
        trend = df.groupby(df[col+'_dt'].dt.to_period('M')).size().rename_axis('Month').reset_index(name='Count')
        if trend.shape[0] > 0:
            fig = px.line(trend, x='Month', y='Count', markers=True, 
                          title=f"Monthly Trend for {col}", color_discrete_sequence=['#00CED1'])
            fig.show()

# -----------------------------
# SAVE PROCESSED REPORT
# -----------------------------
output_file = "AI_Excel_Analysis_Report.xlsx"
df.to_excel(output_file, index=False)
print(f"\n✅ Processed report saved as {output_file}")
files.download(output_file)

