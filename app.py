# app.py
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px

try:
    from pptx import Presentation
    pptx_available = True
except ImportError:
    pptx_available = False

# ------------------------------
# PAGE CONFIG
# ------------------------------
st.set_page_config(page_title="AI-Powered Global Excel Analyzer", layout="wide")
st.title("🤖 AI-Powered Global Excel Analyzer")
st.markdown("Upload ANY Excel file → Get instant AI-driven KPIs, trends, and reports")

# ------------------------------
# FILE UPLOAD
# ------------------------------
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success(f"File uploaded! Rows: {df.shape[0]}, Columns: {df.shape[1]}")

    # ------------------------------
    # DETECT COLUMN TYPES
    # ------------------------------
    numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
    date_cols = df.select_dtypes(include='datetime').columns.tolist()

    # Try to convert object columns to dates
    for col in df.select_dtypes(include='object').columns:
        try:
            df[col] = pd.to_datetime(df[col])
            date_cols.append(col)
        except:
            continue

    numeric_cols = [c for c in df.columns if c in numeric_cols]
    date_cols = list(set(date_cols))
    categorical_cols = [c for c in df.columns if c not in numeric_cols + date_cols]

    id_cols = [c for c in categorical_cols if 'name' in c.lower() or 'id' in c.lower()]
    status_cols = [c for c in categorical_cols if 'status' in c.lower()]

    st.subheader("Detected Column Roles")
    st.markdown(f"🔢 Numeric: {numeric_cols if numeric_cols else 'None'}")
    st.markdown(f"🏷️ Categorical: {categorical_cols if categorical_cols else 'None'}")
    st.markdown(f"📅 Date: {date_cols if date_cols else 'None'}")
    st.markdown(f"🆔 ID/Name: {id_cols if id_cols else 'None'}")
    st.markdown(f"✅ Status: {status_cols if status_cols else 'None'}")

    # ------------------------------
    # AUTOMATIC KPI TABLES
    # ------------------------------
    st.subheader("📊 Automatic KPI & Summary Tables")
    if numeric_cols:
        kpi_df = df[numeric_cols].agg(['count','sum','mean','max','min']).transpose()
        kpi_df = kpi_df.rename(columns={'count':'Total','mean':'Average','max':'Max','min':'Min'})
        st.dataframe(kpi_df)
    else:
        st.info("No numeric columns detected.")

    # ------------------------------
    # AUTOMATIC VISUALIZATIONS
    # ------------------------------
    st.subheader("📈 Automatic Visualizations")
    
    # Safe categorical plots (skip if too many unique values)
    for col in categorical_cols[:5]:
        counts = df[col].value_counts().head(20)
        if counts.empty or counts.shape[0] <= 1:
            st.warning(f"Cannot plot '{col}' (too few values or empty)")
        else:
            fig = px.bar(counts, x=counts.index, y=counts.values, labels={'x': col, 'y':'Count'}, title=f"Top values in {col}")
            st.plotly_chart(fig, use_container_width=True)

    # Numeric trends over dates
    if numeric_cols and date_cols:
        date_col = st.selectbox("Select Date Column for Trend", date_cols)
        numeric_col = st.selectbox("Select Numeric Column for Trend", numeric_cols)
        trend_df = df.groupby(date_col)[numeric_col].sum().reset_index()
        fig = px.line(trend_df, x=date_col, y=numeric_col, title=f"{numeric_col} Trend over {date_col}")
        st.plotly_chart(fig, use_container_width=True)

    # ------------------------------
    # DOWNLOAD PROCESSED DATA
    # ------------------------------
    st.subheader("📥 Download Processed Reports")
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Raw_Data")
        if numeric_cols:
            kpi_df.to_excel(writer, sheet_name="KPI_Summary")
    st.download_button("Download Excel Report", excel_buffer.getvalue(), "AI_Excel_Report.xlsx")

    # ------------------------------
    # OPTIONAL POWERPOINT EXPORT
    # ------------------------------
    if pptx_available:
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = "AI-Powered Excel Report"

        max_rows = min(df.shape[0]+1, 30)
        max_cols = min(df.shape[1], 10)

        table_shape = slide.shapes.add_table(max_rows, max_cols, 0.5, 1.5, 9, 5)
        table = table_shape.table

        for j in range(max_cols):
            table.cell(0,j).text = str(df.columns[j])

        for i in range(1, max_rows):
            for j in range(max_cols):
                if i-1 < df.shape[0]:
                    table.cell(i,j).text = str(df.iloc[i-1,j])
                else:
                    table.cell(i,j).text = ""

        ppt_buffer = BytesIO()
        prs.save(ppt_buffer)
        ppt_buffer.seek(0)
        st.download_button("Download PowerPoint Preview", ppt_buffer, "AI_Excel_Report.pptx")
    else:
        st.info("Install `python-pptx` to enable PowerPoint export.")
else:
    st.info("📂 Please upload an Excel file to start analysis.")


