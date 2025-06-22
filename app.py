import os
import io
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from openai import OpenAI

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
            nonlocal section_count
            if len(chart_data) < 2:
                return

            # Generate chart
            img_stream = io.BytesIO()
            plt.figure(figsize=(6, 4))
            if chart_type == "pie" and len(chart_data) <= 6:
                plt.pie(chart_data, labels=chart_data.index, autopct='%1.1f%%')
            else:
                chart_data.plot(kind="bar")
                plt.xticks(rotation=45, ha='right')
            plt.title(title)
            plt.tight_layout()
            plt.savefig(img_stream, format="png")
            plt.close()
            img_stream.seek(0)

            # Prompt logic
            base_prompt = f"You are a road safety analyst. Write a short professional summary of this accident chart titled '{title}'.\n\nData: {chart_data.to_string()}\n\nHighlight the most common types and any interesting patterns."
            enhanced_prompt = f"You are a road safety expert analyzing a chart titled '{title}'.\n\nData Summary:\n{chart_data.to_string()}\n\nProvide a professional summary highlighting major risks, patterns, and safety-critical findings."
            advanced_prompt = f"You are a transportation safety specialist analyzing a comparative chart titled '{title}'.\n\nData Table:\n{chart_data.to_string()}\n\nSummarize significant trends, focusing on VRU risks, severity levels, or systemic concerns."

            prompt = base_prompt
            if prompt_level == "enhanced":
                prompt = enhanced_prompt
            elif prompt_level == "advanced":
                prompt = advanced_prompt

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
            add_section("Accident Severity Distribution", df['Classification Of Accident'].value_counts(), chart_type="pie", prompt_level="basic")

        if 'Classification Of Accident' in df.columns and 'Location' in df.columns:
            grouped = df.groupby(['Location', 'Classification Of Accident']).size().unstack(fill_value=0)
            if not grouped.empty:
                add_section("Severity by Location", grouped.sum(axis=1).sort_values(ascending=False).head(10), chart_type="bar", prompt_level="advanced")

        for time_col in ['Accident Year', 'Accident Month', 'Accident Day', 'Accident Time']:
            if time_col in df.columns:
                add_section(f"Accidents by {time_col}", df[time_col].value_counts().sort_index(), chart_type="bar", prompt_level="enhanced")

        for env_col in ['Light', 'Environment Condition 1', 'Environment Condition 2']:
            if env_col in df.columns:
                add_section(f"{env_col} Distribution", df[env_col].value_counts(), chart_type="bar", prompt_level="enhanced")

        for impact_col in ['Initial Impact Type', 'Impact Location']:
            if impact_col in df.columns:
                add_section(f"{impact_col} Analysis", df[impact_col].value_counts(), chart_type="pie", prompt_level="basic")

        for driver_col in ['Apparent Driver 1 Action', 'Apparent Driver 2 Action', 'Driver 1 Condition', 'Driver 2 Condition']:
            if driver_col in df.columns:
                add_section(f"{driver_col} Trends", df[driver_col].value_counts(), chart_type="bar", prompt_level="enhanced")

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
