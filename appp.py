import streamlit as st
import pandas as pd
import plotly.express as px
import os
from fpdf import FPDF
import base64

st.set_page_config(page_title="Performance Scorecard", layout="wide")

FILE_NAME = 'For appp.xlsx'

COLUMN_MAP = {
    "ENGG": "Product Design", "ENGG2": "Proto Build+DVP&R", "ENGG3": "SE and EQ", 
    "ENGG4": "Skill Development", "ENGG5": "Knowledge management", "ENGG6": "PRM development", 
    "ENGG7": "Innovation Culture", "ENGG8": "ENGG Overall Score",
    "PE": "Process Design", "PE2": "Facilities Tools & Gauges", "PE3": "Proto Build and PVP&R", 
    "PE4": "Process Qualification", "PE5": "Skill Development (PE)", "PE6": "Knowledge Management (PE)", 
    "PE7": "Process Innovation", "PE8": "Productivity", "PE9": "PE Overall Score",
    "NPQE": "Program Quality", "NPQE2": "Warranty Data", "NPQE3": "Builds Quality", 
    "NPQE4": "Tooled-up Parts", "NPQE5": "Customer Audit", "NPQE6": "MSA & Operators", 
    "NPQE7": "FTR Handover", "NPQE8": "NPQE Overall Score",
    "NPC": "Project Definition", "NPC2": "Project Planning", "NPC3": "Project Execution", 
    "NPC4": "Flawless Launch", "NPC5": "Program Learnings", "NPC6": "NPC Overall Score",
    "CD": "Supplier Review", "CD2": "Supplier FTG", "CD3": "Proto Build (CD)", 
    "CD4": "PPAP Readiness", "CD5": "CD Overall Score"
}

def load_data():
    if not os.path.exists(FILE_NAME):
        st.error(f"File '{FILE_NAME}' not found!")
        return None
    df = pd.read_excel(FILE_NAME, header=0)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def generate_pdf(df_filtered, bu, dept):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(200, 10, txt="Business Performance Report", ln=True, align='C')
    pdf.set_font("Arial", '', 12)
    pdf.cell(200, 10, txt=f"Business Unit: {bu}", ln=True)
    pdf.cell(200, 10, txt=f"Department: {dept}", ln=True)
    pdf.ln(10)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(40, 10, "Year", border=1)
    pdf.cell(100, 10, "Metric", border=1)
    pdf.cell(30, 10, "Score", border=1)
    pdf.ln()
    pdf.set_font("Arial", '', 10)
    for index, row in df_filtered.iterrows():
        year = str(row['Assessment Year'])
        for col in df_filtered.columns:
            if col.startswith(dept):
                metric_name = COLUMN_MAP.get(col, col)
                score = str(row[col])
                pdf.cell(40, 10, year, border=1)
                pdf.cell(100, 10, metric_name[:50], border=1)
                pdf.cell(30, 10, score, border=1)
                pdf.ln()
    return pdf.output(dest='S').encode('latin-1')

df = load_data()

if df is not None:
    st.sidebar.title("App Navigation")
    menu = st.sidebar.radio("Go to", ["Dashboard", "Data Entry"])

    if menu == "Dashboard":
        st.title("📊 Performance Dashboard")
        col1, col2, col3 = st.columns(3)
        with col1:
            # FIX: Filter out NaN from Business list
            bu_list = [bu for bu in df['Business'].unique() if pd.notna(bu)]
            selected_bu = st.selectbox("1. Select Business Unit", bu_list)
        with col2:
            dept_prefixes = ["ENGG", "PE", "NPQE", "NPC", "CD"]
            selected_dept = st.selectbox("2. Select Department", dept_prefixes)
        with col3:
            relevant_cols = [c for c in df.columns if c.startswith(selected_dept)]
            display_options = {c: COLUMN_MAP.get(c, c) for c in relevant_cols}
            selected_col_key = st.selectbox("3. Select Section", options=list(display_options.keys()), format_func=lambda x: display_options[x])

        chart_data = df[df['Business'] == selected_bu].copy()
        
        if not chart_data.empty:
            # FIX: Correct sorting for years (FY 23-24, etc.)
            chart_data['sort_key'] = chart_data['Assessment Year'].str.extract('(\d+)').astype(float)
            chart_data = chart_data.sort_values("sort_key")
            
            fig = px.line(chart_data, x="Assessment Year", y=selected_col_key,
                          title=f"{display_options[selected_col_key]} Trend: {selected_bu}",
                          markers=True, color_discrete_sequence=["#00CC96"])
            fig.update_layout(yaxis_range=[0, 5])
            st.plotly_chart(fig, use_container_width=True)
            chart_data = chart_data.drop(columns=['sort_key'])

            st.divider()
            st.subheader("📥 Export Report")
            ec1, ec2, ec3 = st.columns(3)
            with ec1:
                if st.button("Generate PDF Summary"):
                    pdf_data = generate_pdf(chart_data, selected_bu, selected_dept)
                    st.download_button(label="Click to Download PDF", data=pdf_data, file_name=f"Report_{selected_bu}.pdf", mime="application/pdf")
            with ec2:
                csv = chart_data.to_csv(index=False).encode('utf-8')
                st.download_button(label="Download Data as CSV", data=csv, file_name=f"Data_{selected_bu}.csv", mime='text/csv')
            with ec3:
                st.info("💡 To save full dashboard: Press Ctrl + P and 'Save as PDF'.")

    elif menu == "Data Entry":
        st.title("📝 Data Entry Form")
        with st.form("entry_form"):
            e_col1, e_col2 = st.columns(2)
            # FIX: Filter out NaN from Business list
            bu_list_entry = [bu for bu in df['Business'].unique() if pd.notna(bu)]
            b_unit = e_col1.selectbox("Business Unit", bu_list_entry)
            year = e_col2.text_input("Year (e.g., FY 25-26)")
            st.divider()
            entry_data = {"Business": b_unit, "Assessment Year": year}
            for dept in ["ENGG", "PE", "NPQE", "NPC", "CD"]:
                with st.expander(f"Enter Scores for {dept}"):
                    cols = st.columns(3)
                    relevant_fields = [c for c in df.columns if c.startswith(dept)]
                    for i, field in enumerate(relevant_fields):
                        label = COLUMN_MAP.get(field, field)
                        val = cols[i % 3].number_input(f"{label}", 0.0, 5.0, 0.0, key=f"input_{field}")
                        entry_data[field] = val
            
            if st.form_submit_button("Save Entry to Excel"):
                if not year: st.error("Please enter a Year.")
                else:
                    new_row = pd.DataFrame([entry_data])[df.columns]
                    updated_df = pd.concat([df, new_row], ignore_index=True)
                    updated_df.to_excel(FILE_NAME, index=False)
                    st.success("Success! Excel file updated.")
                    st.balloons()
