import os
import io
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from openai import OpenAI
from datetime import datetime

# Initialize Streamlit
st.set_page_config(page_title="Collision Report Generator", layout="centered")
st.title("🚦 Collision Analysis Report Generator")
st.markdown("Upload an Excel file to generate a detailed Word report.")

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

if uploaded_file:
    with st.spinner("Generating report. Please wait..."):
        df = pd.read_excel(uploaded_file)
        st.success("File read successfully.")

        doc = Document()
        doc.add_heading("Collision Analysis Report", 0)
        doc.add_paragraph("Prepared automatically by Mobility Edge Solution").alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        doc.add_page_break()

        section_count = 1

        def add_section(title, chart_data, chart_type="bar", prompt_level="basic"):
            global section_count
            if len(chart_data) < 2:
                return

            img_stream = io.BytesIO()
            plt.figure(figsize=(6, 4))
            if chart_type == "pie" and len(chart_data) <= 6:
                plt.pie(chart_data, labels=chart_data.index, autopct='%1.1f%%')
            else:
                chart_data.plot(kind="bar", stacked=True if isinstance(chart_data, pd.DataFrame) else False)
                plt.xticks(rotation=45, ha='right')
            plt.title(title)
            plt.tight_layout()
            plt.savefig(img_stream, format="png")
            plt.close()
            img_stream.seek(0)

            prompt = f"You are a traffic safety analyst. The chart below shows accident distribution titled '{title}'. Summarize key findings and highlight any safety-critical patterns.\n\nData:\n{chart_data.head(10).to_string()}"

            try:
                response = client.chat.completions.create(
                    model="gpt-4o",
                    messages=[{"role": "user", "content": prompt}],
                    max_tokens=300
                )
                summary = response.choices[0].message.content.strip()
            except Exception as e:
                summary = f"[GPT Error: {e}]"

            doc.add_heading(f"Section {section_count}: {title}", level=1)
            doc.add_picture(img_stream, width=Inches(5.5))
            caption = doc.add_paragraph(f"Figure {section_count}: {title}")
            caption.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            caption.runs[0].italic = True
            para = doc.add_paragraph(summary)
            para.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            para.runs[0].font.size = Pt(11)
            doc.add_page_break()
            section_count += 1

        # SECTION LOGIC
        if 'Classification Of Accident' in df.columns:
            add_section("Accident Severity Distribution", df['Classification Of Accident'].value_counts(), chart_type="pie")

        if 'Classification Of Accident' in df.columns and 'Location' in df.columns:
            grouped = df.groupby(['Location', 'Classification Of Accident']).size().unstack(fill_value=0)
            if not grouped.empty:
                add_section("Severity by Location", grouped.sum(axis=1).sort_values(ascending=False).head(10), chart_type="bar")

        for time_col in ['Accident Year', 'Accident Month', 'Accident Day']:
            if time_col in df.columns:
                add_section(f"Accidents by {time_col}", df[time_col].value_counts().sort_index(), chart_type="bar")

        for env_col in ['Light', 'Environment Condition 1', 'Environment Condition 2']:
            if env_col in df.columns:
                add_section(f"{env_col} Distribution", df[env_col].value_counts(), chart_type="bar")

        for impact_col in ['Initial Impact Type', 'Impact Location']:
            if impact_col in df.columns:
                add_section(f"{impact_col} Analysis", df[impact_col].value_counts(), chart_type="pie")

        for driver_col in ['Apparent Driver 1 Action', 'Apparent Driver 2 Action', 'Driver 1 Condition', 'Driver 2 Condition']:
            if driver_col in df.columns:
                add_section(f"{driver_col} Trends", df[driver_col].value_counts(), chart_type="bar")

        if 'Accident Date' in df.columns and 'Classification Of Accident' in df.columns:
            try:
                df['Accident Date'] = pd.to_datetime(df['Accident Date'], errors='coerce')
                df['Day of Week'] = df['Accident Date'].dt.day_name()
                weekday_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
                weekday_grouped = df.groupby(['Day of Week', 'Classification Of Accident']).size().unstack(fill_value=0).reindex(weekday_order)
                add_section("Accident Type by Day of Week", weekday_grouped, chart_type="bar")

                df['Day Type'] = df['Accident Date'].dt.dayofweek.apply(lambda x: 'Weekend' if x >= 5 else 'Weekday')
                daytype_grouped = df.groupby(['Day Type', 'Classification Of Accident']).size().unstack(fill_value=0)
                add_section("Accident Type by Weekday vs Weekend", daytype_grouped, chart_type="bar")
            except Exception as e:
                st.warning(f"Could not process date-based charts: {e}")

        if 'Accident Time' in df.columns and 'Classification Of Accident' in df.columns:
            try:
                def classify_period(t):
                    if pd.isnull(t): return 'Unknown'
                    try:
                        hour = int(str(t).split(':')[0])
                        if 6 <= hour < 12: return 'Morning'
                        elif 12 <= hour < 17: return 'Afternoon'
                        elif 17 <= hour < 21: return 'Evening'
                        else: return 'Night'
                    except:
                        return 'Unknown'

                df['Time Period'] = df['Accident Time'].apply(classify_period)
                period_order = ['Morning', 'Afternoon', 'Evening', 'Night', 'Unknown']
                time_grouped = df.groupby(['Time Period', 'Classification Of Accident']).size().unstack(fill_value=0).reindex(period_order)
                add_section("Accident Type by Time of Day", time_grouped, chart_type="bar")
            except Exception as e:
                st.warning(f"Could not process time of day: {e}")

        doc.add_heading(f"Section {section_count}: Spatial Distribution of Accidents", level=1)
        doc.add_paragraph("[Map-based XY scatter plot will be implemented here in future versions.]")
        doc.add_page_break()
        section_count += 1

        doc.add_heading(f"Section {section_count}: Collision Type Diagrams", level=1)
        doc.add_paragraph("[Custom collision type diagrams will be rendered based on type and geometry data in future versions.]")
        doc.add_page_break()

        output_path = "collision_report.docx"
        doc.save(output_path)
        st.success("✅ Report is ready!")
        with open(output_path, "rb") as f:
            st.download_button("📥 Download Report", f, file_name="collision_report.docx")
else:
    st.info("Please upload an Excel file to begin.")
