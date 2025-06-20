import os
import io
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from openai import OpenAI

# Set page layout and title
st.set_page_config(page_title="Collision Report Generator", layout="centered")
st.title("🚦 Collision Analysis Report Generator")
st.markdown("Upload your Excel accident data file to generate a Word report with visual charts and AI-generated descriptions.")

# Load OpenAI key from Streamlit secrets
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# File uploader
uploaded_file = st.file_uploader("📂 Upload your Excel file", type=["xlsx"])

# Confirm app is running
st.write("✅ App is ready. Waiting for upload...")

if uploaded_file:
    with st.spinner("⏳ Generating your report... please wait..."):
        df = pd.read_excel(uploaded_file)
        st.success("✅ File uploaded and read successfully.")

        # Step 1: Detect usable columns
        excluded_cols = ['Latitude', 'Longitude', 'X-Coordinate', 'Y-Coordinate']
        categorical_cols = [
            col for col in df.columns
            if df[col].dtype == 'object'
            and col not in excluded_cols
            and 2 <= df[col].nunique() <= 15
        ]

        st.write("🔎 Columns being analyzed:", categorical_cols)

        if not categorical_cols:
            st.warning("⚠️ No usable categorical columns found in the dataset.")
        else:
            # Step 2: Generate Word report
            doc = Document()
            doc.add_heading("Collision Analysis Report", 0)
            doc.add_paragraph("Generated automatically by Mobility Edge Solution.")
            doc.add_page_break()

            figure_count = 1

            for col in categorical_cols[:5]:  # Limit to 5 for demo/testing
                value_counts = df[col].value_counts()
                if len(value_counts) < 2:
                    continue

                # Create chart in memory
                img_stream = io.BytesIO()
                plt.figure(figsize=(6, 4))
                if len(value_counts) <= 5:
                    plt.pie(value_counts, labels=value_counts.index, autopct='%1.1f%%')
                else:
                    value_counts.plot(kind='bar')
                    plt.xticks(rotation=45, ha='right')
                plt.title(col)
                plt.tight_layout()
                plt.savefig(img_stream, format='png')
                plt.close()
                img_stream.seek(0)

                # GPT prompt
                title = col
                data = value_counts.to_string()
                prompt = (
                    f"You are a road safety analyst. Write a short professional summary "
                    f"of this accident chart titled '{title}'.\n\n"
                    f"Data: {data}\n\n"
                    f"Highlight the most common types and any interesting patterns."
                )

                try:
                    response = client.chat.completions.create(
                        model="gpt-4o",
                        messages=[{"role": "user", "content": prompt}],
                        max_tokens=250
                    )
                    summary = response.choices[0].message.content.strip()
                except Exception as e:
                    summary = f"[GPT error: {e}]"
                    st.warning(summary)

                # Add to Word report
                doc.add_heading(f"{figure_count}. {col}", level=1)
                doc.add_picture(img_stream, width=Inches(5.5))

                caption = doc.add_paragraph(f"Figure {figure_count}: {col} Distribution")
                caption.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                caption.runs[0].italic = True

                para = doc.add_paragraph(summary)
                para.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                para.runs[0].font.size = Pt(11)

                doc.add_page_break()
                figure_count += 1

            # Save and provide download link
            output_path = "collision_report.docx"
            doc.save(output_path)
            st.success("✅ Report is ready!")

            with open(output_path, "rb") as f:
                st.download_button("📥 Download Word Report", f, file_name="collision_report.docx")

else:
    st.info("📄 Please upload an Excel file to get started.")
