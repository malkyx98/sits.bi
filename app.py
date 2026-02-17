import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
from io import BytesIO

# -------------------------------
# Page config
# -------------------------------
st.set_page_config(
    page_title="🤖 AI Global Excel Analyzer",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------
# CSS for colorful system-engineer theme
# -------------------------------
st.markdown("""
<style>
    body {
        background-color: #0f1b2b;
        color: #ffffff;
    }
    .stButton>button {
        background-color: #ff6f61;
        color: white;
    }
    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {
        color: #00d4ff;
    }
    .stDataFrame, .stTable {
        background-color: #1e2a40;
        color: #ffffff;
    }
</style>
""", unsafe_allow_html=True)

# -------------------------------
# Title
# -------------------------------
st.title("🤖 AI-Powered Global Excel Analyzer")
st.markdown("Upload **ANY Excel file** → Get instant AI-driven KPIs, visualizations & insights")

# -------------------------------
# File upload
# -------------------------------
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
if not uploaded_file:
    st.info("📂 Please upload an Excel file to analyze.")
    st.stop()

df = pd.read_excel(uploaded_file)
st.write(f"Uploaded `{uploaded_file.name}`: Rows = {df.shape[0]}, Columns = {df.shape[1]}")

# -------------------------------
# Auto-detect columns
# -------------------------------
numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
date_cols = df.select_dtypes(include='datetime').columns.tolist()
categorical_cols = df.select_dtypes(include='object').columns.tolist()
id_cols = [c for c in df.columns if c not in numeric_cols + categorical_cols + date_cols]

# Try to parse any object column as date
for col in df.columns:
    if col not in numeric_cols + date_cols:
        try:
            df[col] = pd.to_datetime(df[col])
            date_cols.append(col)
        except:
            continue

st.subheader("🔍 Detected Columns")
st.markdown(f"**Numeric:** {numeric_cols if numeric_cols else 'None'}")
st.markdown(f"**Categorical:** {categorical_cols if categorical_cols else 'None'}")
st.markdown(f"**Date:** {date_cols if date_cols else 'None'}")
st.markdown(f"**ID / Name:** {id_cols if id_cols else 'None'}")

# -------------------------------
# KPI summary
# -------------------------------
st.subheader("📊 KPIs for Numeric Columns")
if numeric_cols:
    kpi_df = df[numeric_cols].describe().T
    st.dataframe(kpi_df)
else:
    st.info("No numeric columns detected for KPIs.")

# -------------------------------
# Categorical counts & top values
# -------------------------------
st.subheader("📈 Categorical Value Counts")
for col in categorical_cols[:5]:  # limit to first 5 columns
    counts = df[col].value_counts().head(10)
    st.markdown(f"**{col} - Top Values**")
    st.bar_chart(counts)

# -------------------------------
# Heatmaps
# -------------------------------
if numeric_cols:
    st.subheader("🔥 Numeric Correlation Heatmap")
    corr = df[numeric_cols].corr()
    plt.figure(figsize=(10,6))
    sns.heatmap(corr, annot=True, cmap="coolwarm", fmt=".2f")
    st.pyplot(plt.gcf())

# -------------------------------
# Distribution plots for numeric
# -------------------------------
if numeric_cols:
    st.subheader("📊 Numeric Distributions")
    for col in numeric_cols:
        fig, ax = plt.subplots()
        sns.histplot(df[col].dropna(), kde=True, color="#00d4ff")
        ax.set_title(f"{col} Distribution", color="white")
        ax.set_facecolor("#1e2a40")
        fig.patch.set_facecolor('#0f1b2b')
        st.pyplot(fig)

# -------------------------------
# Trend analysis for date columns
# -------------------------------
if date_cols and numeric_cols:
    st.subheader("📈 Trends Over Time")
    for date_col in date_cols[:3]:  # limit to 3 date columns
        for num_col in numeric_cols[:3]:  # limit to 3 numeric columns
            trend_df = df.groupby(date_col)[num_col].sum().reset_index()
            fig = px.line(trend_df, x=date_col, y=num_col, title=f"{num_col} Trend over {date_col}")
            st.plotly_chart(fig, use_container_width=True)

# -------------------------------
# Download processed Excel
# -------------------------------
st.subheader("📥 Download Processed Excel")
output = BytesIO()
with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False, sheet_name='Processed_Data')
st.download_button("Download Excel", output.getvalue(), f"{uploaded_file.name.split('.')[0]}_processed.xlsx")

