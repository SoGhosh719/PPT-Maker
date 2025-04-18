import streamlit as st
import json
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
import io
import matplotlib.pyplot as plt
import base64
import pandas as pd
import plotly.express as px

# Custom CSS for visual polish
st.markdown(
    """
    <style>
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        border-radius: 8px;
        padding: 8px 16px;
        transition: background-color 0.3s;
    }
    .stButton>button:hover {
        background-color: #45a049;
    }
    .stTab {
        background-color: #f0f2f6;
        border-radius: 8px;
        padding: 10px;
    }
    .stMarkdown h1, .stMarkdown h2 {
        color: #1e3a8a;
    }
    .preview-container {
        border: 2px solid #e0e0e0;
        border-radius: 8px;
        padding: 10px;
        background-color: #ffffff;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Streamlit app title
st.title("📊 PPT Generator from Outline")

# Initialize session state
if "slides" not in st.session_state:
    st.session_state.slides = []
if "undo_stack" not in st.session_state:
    st.session_state.undo_stack = []
if "redo_stack" not in st.session_state:
    st.session_state.redo_stack = []
if "edit_index" not in st.session_state:
    st.session_state.edit_index = None
if "dataset" not in st.session_state:
    st.session_state.dataset = None

# Helper function to add text to a shape with shadow
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

# Helper function to set slide background
def set_slide_background(slide, bg_type, color1, color2=None):
    background = slide.background
    fill = background.fill
    if bg_type == "Solid":
        fill.solid()
        fill.fore_color.rgb = RGBColor(*color1)
    elif bg_type == "Gradient":
        fill.gradient()
        fill.gradient_angle = 90
        stop1 = fill.gradient_stops[0]
        stop1.color.rgb = RGBColor(*color1)
        stop2 = fill.gradient_stops[1]
        stop2.color.rgb = RGBColor(*color2 if color2 else color1)

# Helper function to generate chart preview (for manual chart input)
def generate_chart_preview(chart_type, categories, values, font_color_hex):
    if not categories or not values:
        return None
    plt.figure(figsize=(4, 3))
    if chart_type == "pie":
        plt.pie(values, labels=categories, autopct='%1.1f%%', colors=['#4CAF50', '#FF9800', '#2196F3'])
    elif chart_type == "bar":
        plt.bar(categories, values, color='#4CAF50')
    elif chart_type == "line":
        plt.plot(categories, values, marker='o', color='#4CAF50')
    plt.title("Chart Preview", color=font_color_hex)
    buf = io.BytesIO()
    plt.savefig(buf, format="png", bbox_inches="tight")
    plt.close()
    buf.seek(0)
    return buf

# Tabs for navigation
tab1, tab2, tab3, tab4 = st.tabs(["📝 Add Slides", "🎨 Style Options", "📋 JSON Input", "👀 Preview"])

with tab1:
    # Form for adding/editing slides
    st.header("Add or Edit Slide")
    edit_mode = st.session_state.edit_index is not None
    slide_data = st.session_state.slides[st.session_state.edit_index] if edit_mode else {}
    
    with st.form("slide_form", clear_on_submit=True):
        title = st.text_input(
            "Slide Title", 
            value=slide_data.get("title", ""), 
            placeholder="e.g., Introduction",
            help="Enter the title for your slide."
        )
        bullets = st.multiselect(
            "Bullet Points (add one at a time)",
            options=slide_data.get("content", []) + [""] if edit_mode else [],
            default=slide_data.get("content", []),
            placeholder="Type and select bullet points",
            help="Add bullet points one by one."
        )
        new_bullet = st.text_input("New Bullet Point", placeholder="Type new bullet and add to list")
        if new_bullet and new_bullet not in bullets:
            bullets.append(new_bullet)
        
        # Dataset-driven chart input
        st.subheader("Chart Options")
        chart_input_type = st.radio("Chart Input Method", ["Manual", "Dataset"], index=0)
        
        chart_type = None
        chart_data = None
        plotly_fig = None
        
        if chart_input_type == "Manual":
            chart_type = st.selectbox(
                "Chart Type", 
                ["None", "Pie", "Bar", "Line"],
                index=["None", "Pie", "Bar", "Line"].index(slide_data.get("chart", "None").capitalize()) if slide_data.get("chart") else 0,
                help="Select a chart type or None."
            )
            categories = []
            values = []
            if chart_type != "None":
                st.subheader("Chart Data")
                num_points = st.slider("Number of Data Points", 2, 10, 3)
                cols = st.columns(2)
                with cols[0]:
                    for i in range(num_points):
                        cat = st.text_input(
                            f"Category {i+1}", 
                            value=slide_data.get("chart_data", {}).get("categories", [])[i] if i < len(slide_data.get("chart_data", {}).get("categories", [])) else "",
                            key=f"cat_{i}"
                        )
                        categories.append(cat)
                with cols[1]:
                    for i in range(num_points):
                        val = st.number_input(
                            f"Value {i+1}", 
                            value=slide_data.get("chart_data", {}).get("values", [])[i] if i < len(slide_data.get("chart_data", {}).get("values", [])) else 0.0,
                            key=f"val_{i}"
                        )
                        values.append(val)
                chart_data = {"categories": [c for c in categories if c], "values": [v for v in values if v]}
        
        elif chart_input_type == "Dataset":
            uploaded_file = st.file_uploader("Upload Dataset (CSV)", type=["csv"], key="dataset_upload")
            if uploaded_file:
                try:
                    df = pd.read_csv(uploaded_file)
                    st.session_state.dataset = df
                    st.dataframe(df.head(), use_container_width=True)
                except Exception as e:
                    st.error(f"❌ Invalid CSV file: {str(e)}")
            
            if st.session_state.dataset is not None:
                df = st.session_state.dataset
                chart_type = st.selectbox(
                    "Chart Type",
                    ["Bar", "Line", "Pie", "Scatter"],
                    index=["Bar", "Line", "Pie", "Scatter"].index(slide_data.get("chart", "Bar")) if slide_data.get("chart") else 0,
                    help="Select a chart type for the dataset."
                )
                x_col = st.selectbox("X-Axis Column", df.columns, key="x_col")
                y_col = st.selectbox(
                    "Y-Axis Column",
                    df.select_dtypes(include=['float64', 'int64']).columns,
                    key="y_col"
                )
                category_col = st.selectbox("Category (Optional)", ["None"] + list(df.columns), key="category_col")
                
                # Validate selections
                if not pd.api.types.is_numeric_dtype(df[y_col]):
                    st.error("❌ Y-Axis column must be numeric.")
                elif df[x_col].isna().any() or df[y_col].isna().any():
                    st.error("❌ Selected columns contain missing values.")
                else:
                    # Generate Plotly chart
                    try:
                        if chart_type == "Bar":
                            fig = px.bar(df, x=x_col, y=y_col, color=category_col if category_col != "None" else None)
                        elif chart_type == "Line":
                            fig = px.line(df, x=x_col, y=y_col, color=category_col if category_col != "None" else None)
                        elif chart_type == "Pie":
                            fig = px.pie(df, names=x_col, values=y_col)
                        elif chart_type == "Scatter":
                            fig = px.scatter(df, x=x_col, y=y_col, color=category_col if category_col != "None" else None)
                        fig.update_layout(
                            title=title or "Chart",
                            font=dict(family=st.session_state.get("style", {}).get("font", "Arial"), color=st.session_state.get("style", {}).get("font_color", "#000000"))
                        )
                        st.plotly_chart(fig, use_container_width=True)
                        plotly_fig = fig
                        chart_data = {"x_col": x_col, "y_col": y_col, "category_col": category_col if category_col != "None" else None}
                    except Exception as e:
                        st.error(f"❌ Error generating chart: {str(e)}")
        
        image = st.text_input(
            "Image Filename (from uploaded images)",
            value=slide_data.get("image", ""),
            placeholder="e.g., image1.png",
            help="Enter the filename of an uploaded image."
        )
        slide_transition = st.selectbox(
            "Transition",
            ["None", "Fade", "Push", "Wipe", "Morph", "Zoom"],
            index=["None", "Fade", "Push", "Wipe", "Morph", "Zoom"].index(slide_data.get("transition", "Fade")) if slide_data.get("transition") else 1,
            help="Select a transition effect for the slide."
        )
        submit = st.form_submit_button("Save Slide" if edit_mode else "Add Slide")
        
        if submit:
            if not title.strip():
                st.error("❌ Please provide a slide title.")
            elif chart_input_type == "Dataset" and st.session_state.dataset is None:
                st.error("❌ Please upload a dataset for dataset-driven charts.")
            else:
                new_slide = {
                    "title": title,
                    "content": [b for b in bullets if b],
                    "chart": chart_type.lower() if chart_type and chart_type != "None" else "",
                    "chart_data": chart_data,
                    "chart_input_type": chart_input_type,
                    "plotly_fig": plotly_fig if chart_input_type == "Dataset" else None,
                    "image": image.strip() if image.strip() else None,
                    "transition": slide_transition
                }
                st.session_state.undo_stack.append(st.session_state.slides.copy())
                if edit_mode:
                    st.session_state.slides[st.session_state.edit_index] = new_slide
                    st.session_state.edit_index = None
                else:
                    st.session_state.slides.append(new_slide)
                st.session_state.redo_stack = []
                st.success("✅ Slide saved!" if edit_mode else "✅ Slide added!")

with tab2:
    # Style options with presets
    st.header("Style Options")
    style_presets = {
        "Professional": {"font": "Calibri", "font_color": "#000080", "bg_type": "Gradient", "bg_color1": "#DDE4FF", "bg_color2": "#FFFFFF"},
        "Creative": {"font": "Roboto", "font_color": "#D81B60", "bg_type": "Solid", "bg_color1": "#FFE082", "bg_color2": None},
        "Minimalist": {"font": "Arial", "font_color": "#000000", "bg_type": "Solid", "bg_color1": "#FFFFFF", "bg_color2": None}
    }
    preset = st.selectbox("Select Style Preset", ["Custom"] + list(style_presets.keys()))
    if preset != "Custom" and st.button("Apply Preset"):
        st.session_state.style = style_presets[preset]
    
    font_name = st.selectbox(
        "Font Type",
        [
            "Arial", "Calibri", "Times New Roman", "Helvetica", "Verdana", "Georgia",
            "Roboto", "Open Sans", "Trebuchet MS", "Tahoma", "Segoe UI", "Palatino Linotype",
            "Garamond", "Book Antiqua", "Courier New", "Consolas", "Impact", "Comic Sans MS"
        ],
        index=0 if "style" not in st.session_state else ["Arial", "Calibri", "Times New Roman", "Helvetica", "Verdana", "Georgia", "Roboto", "Open Sans", "Trebuchet MS", "Tahoma", "Segoe UI", "Palatino Linotype", "Garamond", "Book Antiqua", "Courier New", "Consolas", "Impact", "Comic Sans MS"].index(st.session_state.style.get("font", "Arial")),
        help0=0
    )
    st.info("ℹ️ Use widely available fonts (e.g., Arial, Calibri) for compatibility. Embed custom fonts in PowerPoint via File > Options > Save.")
    title_font_size = st.selectbox("Title Font Size", [20, 24, 28, 32], index=1)
    body_font_size = st.selectbox("Body Font Size", [14, 16, 18, 20], index=2)
    font_color_hex = st.color_picker(
        "Font Color", 
        value=st.session_state.style.get("font_color", "#000000") if "style" in st.session_state else "#000000"
    )
    font_color = tuple(int(font_color_hex[i:i+2], 16) for i in (1, 3, 5))
    
    bg_type = st.selectbox(
        "Background Type", 
        ["Solid", "Gradient"],
        index=["Solid", "Gradient"].index(st.session_state.style.get("bg_type", "Solid")) if "style" in st.session_state else 0
    )
    bg_color1_hex = st.color_picker(
        "Background Color 1", 
        value=st.session_state.style.get("bg_color1", "#FFFFFF") if "style" in st.session_state else "#FFFFFF"
    )
    bg_color1 = tuple(int(bg_color1_hex[i:i+2], 16) for i in (1, 3, 5))
    bg_color2_hex = st.color_picker(
        "Background Color 2 (Gradient)", 
        value=st.session_state.style.get("bg_color2", "#DDE4FF") if "style" in st.session_state and st.session_state.style.get("bg_color2") else "#DDE4FF"
    ) if bg_type == "Gradient" else None
    bg_color2 = tuple(int(bg_color2_hex[i:i+2], 16) for i in (1, 3, 5)) if bg_color2_hex else None
    
    layout_name = st.selectbox("Slide Layout", ["Title Slide", "Title and Content", "Blank"], index=1)
    layout_indices = {"Title Slide": 0, "Title and Content": 1, "Blank": 6}
    layout_index = layout_indices[layout_name]
    
    transition = st.selectbox("Default Transition", ["None", "Fade", "Push", "Wipe", "Morph", "Zoom"], index=1)
    
    st.header("Style Preview")
    st.markdown(
        f"""
        <div style='font-family:{font_name}; font-size:{body_font_size}px; color:{font_color_hex}; background-color:{bg_color1_hex}; padding:10px; border-radius:8px;'>
            Sample Text (Font: {font_name}, Size: {body_font_size}pt, Color: {font_color_hex})
        </div>
        """,
        unsafe_allow_html=True
    )

with tab3:
    # JSON input
    st.header("Paste JSON Outline (Advanced)")
    st.write("Paste a JSON outline to replace current slides. Example:")
    st.code(
        '''
[
  {
    "title": "Introduction",
    "content": ["Point 1", "Point 2"],
    "chart": "pie",
    "chart_data": {"categories": ["Category A", "Category B"], "values": [60, 40]},
    "chart_input_type": "Manual",
    "image": "image1.png",
    "transition": "Fade"
  },
  {
    "title": "Data Analysis",
    "content": ["Analysis Results"],
    "chart": "bar",
    "chart_data": {"x_col": "Region", "y_col": "Sales", "category_col": "Year"},
    "chart_input_type": "Dataset",
    "transition": "Wipe"
  }
]
        ''',
        language="json"
    )
    outline_json = st.text_area("Enter PPT Outline (JSON)", height=200, placeholder="Paste your JSON outline here")
    if st.button("Load JSON Outline"):
        if not outline_json.strip():
            st.error("❌ Please provide a JSON outline.")
        else:
            try:
                slides = json.loads(outline_json)
                if not isinstance(slides, list):
                    st.error("❌ Outline must be a list of slides.")
                else:
                    st.session_state.undo_stack.append(st.session_state.slides.copy())
                    st.session_state.slides = slides
                    st.session_state.redo_stack = []
                    st.success("✅ JSON outline loaded!")
            except json.JSONDecodeError:
                st.error("❌ Invalid JSON format.")

with tab4:
    # Slide preview and management
    st.header("Preview and Manage Slides")
    uploaded_images = st.file_uploader("Upload Images (Optional)", type=["png", "jpg"], accept_multiple_files=True, help="Upload images to use in slides.")
    image_files = {f.name: f for f in (uploaded_images or [])}
    
    if st.session_state.slides:
        st.subheader("Current Slides")
        for i, slide in enumerate(st.session_state.slides):
            with st.expander(f"Slide {i+1}: {slide.get('title', 'Untitled')}"):
                st.markdown(f"**Title**: {slide.get('title', 'Untitled')}")
                if slide.get("content"):
                    st.markdown("**Bullet Points**:")
                    for bullet in slide.get("content", []):
                        st.markdown(f"- {bullet}")
                if slide.get("chart") and slide.get("chart_data"):
                    if slide.get("chart_input_type") == "Manual":
                        chart_buf = generate_chart_preview(
                            slide["chart"], 
                            slide["chart_data"]["categories"], 
                            slide["chart_data"]["values"], 
                            font_color_hex
                        )
                        if chart_buf:
                            st.image(chart_buf, caption="Chart Preview", use_column_width=True)
                    elif slide.get("chart_input_type") == "Dataset" and slide.get("plotly_fig"):
                        st.plotly_chart(slide["plotly_fig"], use_container_width=True)
                if slide.get("image") and slide.get("image") in image_files:
                    st.image(image_files[slide["image"]], caption="Image Preview", width=200)
                st.markdown(f"**Transition**: {slide.get('transition', transition)}")
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("Edit", key=f"edit_{i}"):
                        st.session_state.edit_index = i
                        st.experimental_rerun()
                with col2:
                    if st.button("Delete", key=f"delete_{i}"):
                        st.session_state.undo_stack.append(st.session_state.slides.copy())
                        st.session_state.slides.pop(i)
                        st.session_state.redo_stack = []
                        st.success("✅ Slide deleted!")
        
        # Undo/Redo buttons
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Undo", disabled=not st.session_state.undo_stack):
                st.session_state.redo_stack.append(st.session_state.slides.copy())
                st.session_state.slides = st.session_state.undo_stack.pop()
                st.success("✅ Undone!")
        with col2:
            if st.button("Redo", disabled=not st.session_state.redo_stack):
                st.session_state.undo_stack.append(st.session_state.slides.copy())
                st.session_state.slides = st.session_state.redo_stack.pop()
                st.success("✅ Redone!")
    
    else:
        st.info("No slides added yet. Use the 'Add Slides' tab to create slides.")

# Generate PPT
if st.button("Generate PPT", key="generate_ppt"):
    if not st.session_state.slides:
        st.error("❌ Please add at least one slide.")
    else:
        try:
            prs = Presentation()
            for slide_data in st.session_state.slides:
                if not isinstance(slide_data, dict):
                    st.warning(f"⚠️ Skipping invalid slide data: {slide_data}")
                    continue
                
                title = slide_data.get("title", "Untitled")
                content = slide_data.get("content", [])
                chart_type = slide_data.get("chart", "").lower()
                chart_data_input = slide_data.get("chart_data", None)
                chart_input_type = slide_data.get("chart_input_type", "Manual")
                plotly_fig = slide_data.get("plotly_fig", None)
                image_path = slide_data.get("image", None)
                slide_transition = slide_data.get("transition", transition)
                
                slide_layout = prs.slide_layouts[layout_index]
                slide = prs.slides.add_slide(slide_layout)
                set_slide_background(slide, bg_type, bg_color1, bg_color2)
                
                if slide.shapes.title:
                    title_shape = slide.shapes.title
                else:
                    title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1))
                add_text_to_shape(title_shape, title, font_name=font_name, font_size=title_font_size, bold=True, font_color=font_color)
                
                if content:
                    add_bullet_list(slide, Inches(0.5), Inches(1.5), Inches(4), Inches(4), content, font_name=font_name, font_size=body_font_size, font_color=font_color)
                
                if chart_type in ["pie", "bar", "line"] and chart_data_input and chart_input_type == "Manual":
                    try:
                        chart_data = CategoryChartData()
                        chart_data.categories = chart_data_input["categories"]
                        chart_data.add_series("Data", chart_data_input["values"])
                        chart_type_enum = {
                            "pie": XL_CHART_TYPE.PIE,
                            "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
                            "line": XL_CHART_TYPE.LINE
                        }[chart_type]
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
                
                if chart_type in ["bar", "line", "pie", "scatter"] and chart_input_type == "Dataset" and plotly_fig:
                    try:
                        chart_buf = plotly_fig.write_image(format="png")
                        chart_stream = io.BytesIO(chart_buf)
                        slide.shapes.add_picture(chart_stream, Inches(5), Inches(1.5), width=Inches(4))
                    except Exception as e:
                        st.warning(f"⚠️ Failed to add dataset chart for slide '{title}': {str(e)}")
                
                if image_path and image_path in image_files:
                    try:
                        img_stream = io.BytesIO(image_files[image_path].read())
                        slide.shapes.add_picture(img_stream, Inches(0.5), Inches(3.5), width=Inches(4))
                    except Exception as e:
                        st.warning(f"⚠️ Failed to add image '{image_path}': {str(e)}")
                
                slide.notes_slide.notes_text_frame.text = f"Recommended transition: {slide_transition}"
            
            buffer | io.BytesIO()
            prs.save(buffer)
            buffer.seek(0)
            
            st.download_button(
                label="📥 Download PPT",
                data=buffer,
                file_name="Generated_Presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
            st.success("✅ PPT generated successfully!")
        
        except ValueError as e:
            st.error(f"❌ Error in slide structure: {str(e)}")
        except Exception as e:
            st.error(f"❌ Error generating PPT: {str(e)}")

# Tutorial button
if st.button("Show Tutorial"):
    st.markdown("""
    **Welcome to the PPT Generator!** Follow these steps:
    1. **Add Slides**: Use the 'Add Slides' tab to create slides with titles, bullet points, charts (manual or dataset-driven), and images.
    2. **Customize Styles**: In the 'Style Options' tab, choose fonts, colors, and backgrounds. Try a preset for quick styling!
    3. **Advanced Input**: Use the 'JSON Input' tab to paste a JSON outline (optional).
    4. **Preview and Edit**: Check your slides in the 'Preview' tab, edit or delete as needed.
    5. **Generate PPT**: Click 'Generate PPT' to download your presentation.
    """)
