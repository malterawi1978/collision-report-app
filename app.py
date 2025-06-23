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
from PIL import Image, ImageDraw, ImageFont
from staticmap import StaticMap, CircleMarker

st.set_page_config(
    page_title="Collisio – Automated Collision Report Generator",
    page_icon="🚦",
    layout="centered"
)

logo = Image.open("Collisio_Logo.png")
st.image(logo, width=200)

st.title("🤖 Collisio")
st.markdown("### Automated Collision Report Generator")
st.markdown("Upload your traffic accident data to generate a smart report with charts and insights powered by AI.")

st.markdown("**Need help formatting your accident data?**")
st.markdown("Download our ready-made Excel template to ensure your data is structured correctly before upload.")

with open("collision_template.xlsx", "rb") as f:
    st.download_button(
        label="📅 Download Excel Template",
        data=f,
        file_name="collision_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

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

        def add_section(title, chart_data, chart_type="bar"):
            nonlocal section_count
            if len(chart_data) < 2:
                return

            img_stream = io.BytesIO()
            plt.figure(figsize=(6, 4))
            if chart_type == "pie" and len(chart_data) <= 6:
                plt.pie(chart_data, labels=chart_data.index, autopct='%1.1f%%')
            else:
                ax = chart_data.plot(kind="bar", stacked=isinstance(chart_data, pd.DataFrame))
                plt.xticks(rotation=45, ha='right')
                ax.set_ylabel("Number of Accidents")
            plt.title(title)
            plt.tight_layout()
            plt.savefig(img_stream, format="png")
            plt.close()
            img_stream.seek(0)

            prompt = (
                f"You are a road safety analyst. Write a short professional summary "
                f"of this accident chart titled '{title}'.\n\n"
                f"Data: {chart_data.head(10).to_string()}\n\n"
                f"Highlight the most common types and any interesting patterns."
            )
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

        if 'Classification Of Accident' in df.columns:
            add_section("Accident Severity Distribution", df['Classification Of Accident'].value_counts(), chart_type="pie")

        if 'Classification Of Accident' in df.columns and 'Location' in df.columns:
            grouped = df.groupby(['Location', 'Classification Of Accident']).size().unstack(fill_value=0)
            if not grouped.empty:
                add_section("Severity by Location", grouped.sum(axis=1).sort_values(ascending=False).head(10), chart_type="bar")

        for time_col in ['Accident Year', 'Accident Day']:
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

                df['Accident Month'] = df['Accident Date'].dt.month_name()
                month_order = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
                month_grouped = df.groupby(['Accident Month', 'Classification Of Accident']).size().unstack(fill_value=0).reindex(month_order)
                add_section("Accident Type by Month", month_grouped, chart_type="bar")
            except Exception as e:
                st.warning(f"Could not process date-based charts: {e}")

        if 'Accident Time' in df.columns and 'Classification Of Accident' in df.columns:
            try:
                def clean_time_string(t):
                    if pd.isnull(t): return None
                    t = str(t).lower().strip().replace('pm', '').replace('am', '')
                    try:
                        dt = datetime.strptime(t.strip(), '%H:%M:%S')
                    except:
                        try:
                            dt = datetime.strptime(t.strip(), '%I:%M:%S')
                        except:
                            return None
                    return dt.strftime('%H:%M')

                df['Cleaned Time'] = df['Accident Time'].apply(clean_time_string)

                def classify_period(t):
                    try:
                        if t is None:
                            return 'Unknown'
                        hour = int(t.split(':')[0])
                        if 6 <= hour < 12: return 'Morning'
                        elif 12 <= hour < 17: return 'Afternoon'
                        elif 17 <= hour < 21: return 'Evening'
                        else: return 'Night'
                    except:
                        return 'Unknown'

                df['Time Period'] = df['Cleaned Time'].apply(classify_period)
                period_order = ['Morning', 'Afternoon', 'Evening', 'Night']
                df_filtered = df[df['Time Period'].isin(period_order)]
                time_grouped = df_filtered.groupby(['Time Period', 'Classification Of Accident']).size().unstack(fill_value=0).reindex(period_order)
                add_section("Accident Type by Time of Day", time_grouped, chart_type="bar")
            except Exception as e:
                st.warning(f"Could not process time of day: {e}")

        try:
            if 'Latitude' in df.columns and 'Longitude' in df.columns and 'Classification Of Accident' in df.columns:
    type_counts = df['Classification Of Accident'].value_counts()
    base_colors = ['#1f77b4', '#ff7f0e', '#2ca02c']
    legend_items = {label: base_colors[i % len(base_colors)] for i, label in enumerate(type_counts.index.tolist())}

    map_df = df[['Latitude', 'Longitude', 'Classification Of Accident']].dropna()
    map_df = map_df[(map_df['Latitude'].between(-90, 90)) & (map_df['Longitude'].between(-180, 180))]

    if not map_df.empty:
        smap = StaticMap(800, 600)

        for _, row in map_df.iterrows():
            acc_type_label = str(row['Classification Of Accident']).strip()
            color = legend_items.get(acc_type_label, '#0000ff')
            marker = CircleMarker((row['Longitude'], row['Latitude']), color, 10)
            smap.add_marker(marker)

        image = smap.render()
        overlay = Image.new('RGBA', image.size, (255, 255, 255, 40))
        image = Image.alpha_composite(image.convert("RGBA"), overlay)

        draw = ImageDraw.Draw(image)
        font = ImageFont.load_default()
        legend_x, legend_y = 20, image.size[1] - 100
        draw.rectangle([legend_x - 10, legend_y - 10, legend_x + 200, legend_y + 20 + 18 * len(legend_items)], fill="white", outline="gray")
        draw.text((legend_x, legend_y), "Accident Type:", fill="black", font=font)
        y_offset = 15
        for label, color in legend_items.items():
            draw.ellipse([legend_x, legend_y + y_offset, legend_x + 10, legend_y + y_offset + 10], fill=color)
            draw.text((legend_x + 15, legend_y + y_offset - 2), label, fill="black", font=font)
            y_offset += 18

        map_path = "static_map.png"
        image.convert("RGB").save(map_path)

        doc.add_heading(f"Section {section_count}: Spatial Distribution of Accidents", level=1)
        doc.add_picture(map_path, width=Inches(5.5))
        caption = doc.add_paragraph(f"Figure {section_count}: Map showing accident locations colored by type with legend.")
        caption.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        caption.runs[0].italic = True
        doc.add_page_break()
        section_count += 1
    else:
        st.warning("Latitude/Longitude data is present but empty or out of bounds.")
        section_count += 1
                    section_count += 1
                else:
                    st.warning("Latitude/Longitude data is present but empty or out of bounds.")
        except Exception as e:
            st.warning(f"Could not generate static map: {e}")

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
