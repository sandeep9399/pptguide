# Apollo Slide Visual Guide App (v2)
# Upload a PPT → Get back slide-by-slide layout, visual guidance, and downloadable design suggestions

import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
import io
import re
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE

st.set_page_config(page_title="PPT Visual Guide Generator", layout="wide")
st.title("🎯 Apollo PPT Visual Design Guide")
st.markdown("Upload your PowerPoint presentation below. We'll analyze and generate layout, visual suggestions, and export-ready guides.")

uploaded_file = st.file_uploader("Upload your .pptx file", type=["pptx"])

def clean_illegal_chars(text):
    ILLEGAL_CHARACTERS_RE = re.compile(r"[\000-\010\013\014\016-\037]")
    return ILLEGAL_CHARACTERS_RE.sub("", text)

def suggest_design_elements(full_text):
    if "who" in full_text:
        return {
            "Suggested Layout": "Two-column: WHO quote on left, image on right",
            "Typography": "Poppins Bold 36pt title, Segoe UI 22pt body",
            "Color Theme": "Blue-Grey healthcare theme",
            "Icon Style": "Line icons (globe, stethoscope)",
            "Animation": "Fade-in for text, fly-in for image"
        }
    elif "components" in full_text:
        return {
            "Suggested Layout": "4-quadrant grid",
            "Typography": "Arial Rounded 28pt bold titles, 20pt text",
            "Color Theme": "Color blocks: blue, green, orange, violet",
            "Icon Style": "Health category icons (🧠 ❤️ 🧘‍♀️ 👥)",
            "Animation": "Sequential fade-in"
        }
    elif "india" in full_text:
        return {
            "Suggested Layout": "Data + map overlay",
            "Typography": "Segoe UI Bold 30pt, Regular 20pt",
            "Color Theme": "Warm tones + India map overlay",
            "Icon Style": "Flat infographic symbols",
            "Animation": "Bar chart build-up"
        }
    else:
        return {
            "Suggested Layout": "Title + Image Right",
            "Typography": "Calibri Light 32pt title, 20pt body",
            "Color Theme": "Light pastel healthcare theme",
            "Icon Style": "Simple flat icons",
            "Animation": "Appear on click"
        }

def analyze_ppt(ppt_file):
    prs = Presentation(ppt_file)
    data = []

    for i, slide in enumerate(prs.slides, start=1):
        block_texts = [
            clean_illegal_chars(shape.text.strip()) for shape in slide.shapes
            if shape.has_text_frame and shape.text.strip()
        ]
        full_text = " ".join(block_texts).lower()
        split_required = len(block_texts) > 3 or any(k in full_text for k in ["components", "diseases", "determinants", "definition"])

        part_count = 2 if split_required else 1
        for idx in range(part_count):
            part = f"Part {idx+1}" if part_count == 2 else "Full Slide"
            design = suggest_design_elements(full_text)

            data.append({
                "Slide Number": i,
                "Slide Part": part,
                "Block Title": block_texts[0][:60] + "..." if block_texts else "Untitled",
                "Content Alignment": design["Suggested Layout"],
                "Suggested Visual Style": design["Icon Style"],
                "Designer Note": "Split this into multiple sections." if part != "Full Slide" else "All content fits in one slide.",
                "Typography": design["Typography"],
                "Color Theme": design["Color Theme"],
                "Animation": design["Animation"]
            })

    return pd.DataFrame(data)

if uploaded_file:
    df = analyze_ppt(uploaded_file)
    st.success("✅ Analysis complete!")
    st.dataframe(df, use_container_width=True)

    # Excel export
    towrite = io.BytesIO()
    df.to_excel(towrite, index=False, engine='openpyxl')
    towrite.seek(0)

    st.download_button(
        label="📥 Download Excel",
        data=towrite,
        file_name="Slide_Visual_Design_Guide.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # PPT mockup (optional generation)
    output_ppt = Presentation()
    title_slide_layout = output_ppt.slide_layouts[1]

    for _, row in df.iterrows():
        slide = output_ppt.slides.add_slide(title_slide_layout)
        slide.shapes.title.text = f"Slide {int(row['Slide Number'])} - {row['Slide Part']}"
        content = slide.placeholders[1].text_frame
        content.clear()
        for field in ["Block Title", "Content Alignment", "Suggested Visual Style", "Typography", "Color Theme", "Animation"]:
            p = content.add_paragraph()
            p.text = f"{field}: {row[field]}"

    ppt_io = io.BytesIO()
    output_ppt.save(ppt_io)
    ppt_io.seek(0)

    st.download_button(
        label="🎞️ Download PPT Design Mockup",
        data=ppt_io,
        file_name="Apollo_Design_Guide.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
