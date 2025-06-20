# Apollo Slide Visual Guide App
# Upload a PPT → Get back slide-by-slide layout, alignment, visual style, and designer instructions

import streamlit as st
import pandas as pd
from pptx import Presentation
import io
import re

st.set_page_config(page_title="PPT Visual Guide Generator", layout="wide")
st.title("🎯 Apollo PPT Visual Design Guide")
st.markdown("Upload your PowerPoint presentation below. We'll analyze and generate layout + visual suggestions for each slide.")

uploaded_file = st.file_uploader("Upload your .pptx file", type=["pptx"])

def clean_illegal_chars(text):
    # Remove illegal Unicode characters that cannot be used in Excel
    ILLEGAL_CHARACTERS_RE = re.compile(r"[\000-\010\013\014\016-\037]")
    return ILLEGAL_CHARACTERS_RE.sub("", text)

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
            content_align = (
                "Grid layout with even spacing" if "components" in full_text
                else "Center-aligned callout" if "definition" in full_text
                else "Left aligned with icons" if "objectives" in full_text
                else "Standard layout"
            )
            visual_style = (
                "4-quadrant icons" if "components" in full_text
                else "Quote bubble with WHO branding" if "definition" in full_text
                else "Checklist infographic" if "objective" in full_text
                else "Thematic healthcare image"
            )
            designer_note = (
                "Split: place this in the first half of the sequence." if part == "Part 1" else
                "Split: continue this from the previous slide." if part == "Part 2" else
                "All content can be on one slide."
            )

            data.append({
                "Slide Number": i,
                "Slide Part": part,
                "Block Title": block_texts[0][:60] + "..." if block_texts else "Untitled",
                "Content Alignment": content_align,
                "Suggested Visual Style": visual_style,
                "Designer Note": designer_note
            })

    return pd.DataFrame(data)

if uploaded_file:
    df = analyze_ppt(uploaded_file)
    st.success("✅ Analysis complete!")
    st.dataframe(df, use_container_width=True)

    towrite = io.BytesIO()
    df.to_excel(towrite, index=False, engine='openpyxl')
    towrite.seek(0)

    st.download_button(
        label="📥 Download Excel",
        data=towrite,
        file_name="Slide_Visual_Guidelines.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
