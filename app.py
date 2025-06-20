import os
import io
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from openai import OpenAI

# Set page settings
st.set_page_config(page_title="Collision Report Generator", layout="centered")

# Load OpenAI API key from Streamlit Secrets
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# App title
st.title("🚦 Collision Analysis Report Generator")
st.write("Upload an Excel file containing accident data to generate your report.")

st.markdown("""
<style>
    .report-title {
        font-size: 22px;
        font-weight: 600;
        color: #2c3e50;
    }
    .caption {
        font-size: 16px;
        color: gray;
        margin-bottom: 20px;
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="report-title">Mobility Edge Solution – Collision Analysis Tool</div>', unsafe_allow_html=True)
st.markdown('<div class="caption">Upload your data and receive a formatted collision summary report powered by AI.</div>', unsafe_allow_html=True)


# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    with st.spinner("⏳ Generating your report... please wait..."):
        df = pd.read_excel(uploaded_file)
        ...
        # (Keep all report logic here)

    st.success("✅ Report is ready!")

        # Select relevant categorical columns
        excluded_cols = ['Latitude', 'Longitude', 'X-Coordinate', 'Y-Coordinate']
        categorical_cols = [
            col for col in df.columns
            if df[col].dtype == 'object'
            and col not in excluded_cols
            and 2 <= df[col].nunique() <= 15
        ]

        if not categorical_cols:
            st.warning("⚠️ No suitable categorical columns found for analysis.")
        else:
            # Start building Word report
            doc = Document()
            doc.add_heading("Collision Analysis Report", 0)
            doc.add_paragraph("Generated automatically by Mobility Edge Solution")
            doc.add_page_break()

            figure_count = 1

            for col in categorical_cols:
                value_counts = df[col].value_counts()
                if len(value_counts) < 2:
                    continue

                # Chart creation (in memory)
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

                # Generate GPT summary
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
                    summary = f"[⚠️ GPT error: {e}]"

                # Add section to Word report
                doc.add_heading(f"{figure_count}. {col}", level=1)
                doc.add_picture(img_stream, width=Inches(5.5))

                caption = doc.add_paragraph(f"Figure {figure_count}: {col} Distribution")
                caption.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                caption.runs[0].italic = True

                paragraph = doc.add_paragraph(summary)
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                paragraph.runs[0].font.size = Pt(11)

                doc.add_page_break()
                figure_count += 1

            # Save the Word report
            report_path = "collision_report.docx"
            doc.save(report_path)
            st.success("✅ Report generated!")

            # Download button
            with open(report_path, "rb") as f:
                st.download_button("📥 Download Report", f, file_name="collision_report.docx")

    except Exception as e:
        st.error(f"❌ Error processing the file: {e}")
else:
    st.info("📂 Please upload an Excel file to begin.")
