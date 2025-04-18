import streamlit as st
import json
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE  # Correct import
from pptx.dml.fill import FillFormat
from pptx.oxml.xmlchemy import OxmlElement
import io

# Streamlit app title
st.title("📊 PPT Generator from Outline")

# Helper function to add text to a shape
def add_text_to_shape(shape, text, font_name, font_size=18, bold=False, font_color=(0, 0, 0)):
    text_frame = shape.text_frame
    text_frame.clear()
    p = text_frame.paragraphs[0]
    run = p.add_run()
    run.text = text
    font = run.font
    font.name = font_name
    font.size = Pt(font_size)
    font.bold = bold
    font.color.rgb = RGBColor(*font_color)

# Helper function to add a bullet list
def add_bullet_list(slide, left, top, width, height, bullets, font_name, font_size=18, font_color=(0, 0, 0)):
    text_box = slide.shapes.add_textbox(left, top, width, height)
    text_frame = text_box.text_frame
    text_frame.word_wrap = True
    for bullet in bullets:
        p = text_frame.add_paragraph()
        p.text = bullet
        p.level = 0
        p.font.name = font_name
        p.font.size = Pt(font_size)
        p.font.color.rgb = RGBColor(*font_color)
        p.alignment = PP_ALIGN.LEFT

# Helper function to set slide background (solid or gradient)
def set_slide_background(slide, bg_type, color1, color2=None):
    background = slide.background
    fill = background.fill
    if bg_type == "Solid":
        fill.solid()
        fill.fore_color.rgb = RGBColor(*color1)
    elif bg_type == "Gradient":
        fill.gradient()
        fill.gradient_angle = 90  # Vertical gradient
        stop1 = fill.gradient_stops[0]
        stop1.color.rgb = RGBColor(*color1)
        stop2 = fill.gradient_stops[1]
        stop2.color.rgb = RGBColor(*color2 if color2 else color1)

# Sidebar for style options
st.sidebar.header("🎨 Presentation Style Options")
font_name = st.sidebar.selectbox(
    "Font Type",
    ["Arial", "Calibri", "Times New Roman", "Helvetica", "Verdana", "Georgia", "Roboto"],
    index=0
)
title_font_size = st.sidebar.selectbox(
    "Title Font Size",
    [20, 24, 28, 32],
    index=1
)
body_font_size = st.sidebar.selectbox(
    "Body Font Size",
    [14, 16, 18, 20],
    index=2
)
font_color_hex = st.sidebar.color_picker("Font Color", value="#000000")  # Black
font_color = tuple(int(font_color_hex[i:i+2], 16) for i in (1, 3, 5))

bg_type = st.sidebar.selectbox("Background Type", ["Solid", "Gradient"], index=0)
bg_color1_hex = st.sidebar.color_picker("Background Color 1", value="#FFFFFF")  # White
bg_color1 = tuple(int(bg_color1_hex[i:i+2], 16) for i in (1, 3, 5))
bg_color2_hex = st.sidebar.color_picker("Background Color 2 (Gradient)", value="#DDE4FF") if bg_type == "Gradient" else None
bg_color2 = tuple(int(bg_color2_hex[i:i+2], 16) for i in (1, 3, 5)) if bg_color2_hex else None

layout_name = st.sidebar.selectbox(
    "Slide Layout",
    ["Title Slide", "Title and Content", "Blank"],
    index=1
)
layout_indices = {
    "Title Slide": 0,
    "Title and Content": 1,
    "Blank": 6
}
layout_index = layout_indices[layout_name]

transition = st.sidebar.selectbox(
    "Transition Effect (Apply in PowerPoint)",
    ["None", "Fade", "Push", "Wipe", "Morph", "Zoom"],
    index=1
)

# Style preview
st.sidebar.header("Style Preview")
st.markdown(
    f"""
    <div style='font-family:{font_name}; font-size:{body_font_size}px; color:{font_color_hex}; background-color:{bg_color1_hex}; padding:10px;'>
        Sample Text (Font: {font_name}, Size: {body_font_size}pt, Color: {font_color_hex})
    </div>
    """,
    unsafe_allow_html=True
)

# Image uploader
uploaded_images = st.file_uploader("Upload Images (Optional)", type=["png", "jpg"], accept_multiple_files=True)

# Instructions for user
st.write("Enter a JSON outline for the PPT. Example:")
st.code(
    '''
[
  {
    "title": "Introduction",
    "content": ["Point 1", "Point 2"],
    "chart": "pie",
    "chart_data": {"categories": ["A", "B"], "values": [60, 40]},
    "image": "image1.png"
  },
  {
    "title": "Conclusion",
    "content": ["Summary"]
  }
]
    ''',
    language="json"
)

# Text area for JSON input
outline_json = st.text_area("Enter PPT Outline (JSON)", height=200, placeholder="Paste your JSON outline here")

# Button to generate PPT
if st.button("Generate PPT"):
    if not outline_json.strip():
        st.error("❌ Please provide a JSON outline.")
    else:
        try:
            # Parse JSON input
            outline = json.loads(outline_json)
            
            # Validate outline is a list
            if not isinstance(outline, list):
                raise ValueError("Outline must be a list of slides.")
            
            # Create a new presentation
            prs = Presentation()
            
            # Map uploaded images to filenames
            image_files = {f.name: f for f in (uploaded_images or [])}
            
            # Iterate through outline to create slides
            for slide_data in outline:
                if not isinstance(slide_data, dict):
                    st.warning(f"⚠️ Skipping invalid slide data: {slide_data}")
                    continue
                
                title = slide_data.get("title", "Untitled")
                content = slide_data.get("content", [])
                chart_type = slide_data.get("chart", "").lower()
                chart_data_input = slide_data.get("chart_data", None)
                image_path = slide_data.get("image", None)
                
                # Add slide with selected layout
                slide_layout = prs.slide_layouts[layout_index]
                slide = prs.slides.add_slide(slide_layout)
                
                # Set background
                set_slide_background(slide, bg_type, bg_color1, bg_color2)
                
                # Set title
                if slide.shapes.title:
                    title_shape = slide.shapes.title
                else:
                    title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
                add_text_to_shape(title_shape, title, font_name=font_name, font_size=title_font_size, bold=True, font_color=font_color)
                
                # Add bullet points
                if content:
                    add_bullet_list(slide, Inches(0.5), Inches(1.5), Inches(4), Inches(4), content, font_name=font_name, font_size=body_font_size, font_color=font_color)
                
                # Add chart if specified
                if chart_type in ["pie", "bar"] and chart_data_input:
                    try:
                        chart_data = CategoryChartData()
                        chart_data.categories = chart_data_input["categories"]
                        chart_data.add_series("Data", chart_data_input["values"])
                        chart_type_enum = XL_CHART_TYPE.PIE if chart_type == "pie" else XL_CHART_TYPE.COLUMN_CLUSTERED
                        chart = slide.shapes.add_chart(
                            chart_type_enum, Inches(5), Inches(1.5), Inches(4), Inches(3), chart_data
                        ).chart
                        chart.has_title = True
                        chart.chart_title.text_frame.text = f"{title} Chart"
                        p = chart.chart_title.text_frame.paragraphs[0]
                        p.font.name = font_name
                        p.font.size = Pt(14)
                        p.font.color.rgb = RGBColor(*font_color)
                    except (KeyError, TypeError, ValueError) as e:
                        st.warning(f"⚠️ Invalid chart data for slide '{title}': {str(e)}")
                
                # Add image if specified
                if image_path and image_path in image_files:
                    try:
                        img_stream = io.BytesIO(image_files[image_path].read())
                        slide.shapes.add_picture(img_stream, Inches(0.5), Inches(3.5), width=Inches(4))
                    except Exception as e:
                        st.warning(f"⚠️ Failed to add image '{image_path}': {str(e)}")
                
                # Add transition note
                slide.notes_slide.notes_text_frame.text = f"Recommended transition: {transition}"
            
            # Save PPT to a BytesIO buffer
            buffer = io.BytesIO()
            prs.save(buffer)
            buffer.seek(0)
            
            # Provide download button
            st.download_button(
                label="📥 Download PPT",
                data=buffer,
                file_name="Generated_Presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
            st.success("✅ PPT generated successfully!")
        
        except json.JSONDecodeError:
            st.error("❌ Invalid JSON format. Please check your input.")
        except ValueError as e:
            st.error(f"❌ Error in outline structure: {str(e)}")
        except Exception as e:
            st.error(f"❌ Error generating PPT: {str(e)}")
