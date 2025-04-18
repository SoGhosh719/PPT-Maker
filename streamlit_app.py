import streamlit as st
import json
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.charts import XL_CHART_TYPE
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

# Helper function to set slide background color
def set_slide_background(slide, rgb_color):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*rgb_color)

# Sidebar for style options
st.sidebar.header("🎨 Presentation Style Options")
font_name = st.sidebar.selectbox(
    "Font Type",
    ["Arial", "Calibri", "Times New Roman", "Helvetica"],
    index=0
)
font_color_name = st.sidebar.selectbox(
    "Font Color",
    ["Black", "Dark Blue", "Dark Gray"],
    index=0
)
font_colors = {
    "Black": (0, 0, 0),
    "Dark Blue": (0, 0, 128),
    "Dark Gray": (64, 64, 64)
}
font_color = font_colors[font_color_name]

background_color_name = st.sidebar.selectbox(
    "Slide Background Color",
    ["White", "Light Gray", "Light Blue"],
    index=0
)
background_colors = {
    "White": (255, 255, 255),
    "Light Gray": (240, 240, 240),
    "Light Blue": (200, 220, 255)
}
background_color = background_colors[background_color_name]

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

# Instructions for user
st.write("Enter a JSON outline for the PPT. Example:")
st.code(
    '''
[
  {
    "title": "Introduction",
    "content": ["Point 1", "Point 2"],
    "chart": "pie",
    "chart_data": {"categories": ["A", "B"], "values": [60, 40]}
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
            
            # Iterate through outline to create slides
            for slide_data in outline:
                if not isinstance(slide_data, dict):
                    st.warning(f"Skipping invalid slide data: {slide_data}")
                    continue
                
                title = slide_data.get("title", "Untitled")
                content = slide_data.get("content", [])
                chart_type = slide_data.get("chart", "").lower()
                chart_data_input = slide_data.get("chart_data", None)
                
                # Add slide with selected layout
                slide_layout = prs.slide_layouts[layout_index]
                slide = prs.slides.add_slide(slide_layout)
                
                # Set background color
                set_slide_background(slide, background_color)
                
                # Set title
                if slide.shapes.title:
                    title_shape = slide.shapes.title
                else:
                    title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
                add_text_to_shape(title_shape, title, font_name=font_name, font_size=24, bold=True, font_color=font_color)
                
                # Add bullet points
                if content:
                    add_bullet_list(slide, Inches(0.5), Inches(1.5), Inches(4), Inches(4), content, font_name=font_name, font_size=18, font_color=font_color)
                
                # Add chart if specified
                if chart_type == "pie" and chart_data_input:
                    try:
                        chart_data = CategoryChartData()
                        chart_data.categories = chart_data_input["categories"]
                        chart_data.add_series("Data", chart_data_input["values"])
                        chart = slide.shapes.add_chart(
                            XL_CHART_TYPE.PIE, Inches(5), Inches(1.5), Inches(4), Inches(3), chart_data
                        ).chart
                        chart.has_title = True
                        chart.chart_title.text_frame.text = f"{title} Chart"
                        
                        # Apply font style to chart title
                        p = chart.chart_title.text_frame.paragraphs[0]
                        p.font.name = font_name
                        p.font.size = Pt(14)
                        p.font.color.rgb = RGBColor(*font_color)
                    except (KeyError, TypeError, ValueError) as e:
                        st.warning(f"⚠️ Invalid chart data for slide '{title}': {str(e)}")
                
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
