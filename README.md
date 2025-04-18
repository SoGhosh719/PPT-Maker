# 📊 Automated PPT Generator  
*Create customizable PowerPoint presentations from outlines, JSON structures, or uploaded datasets with ease.*

---

## 🚀 Overview

**Streamlit PPT Generator** is a dynamic web app built with Python and Streamlit that allows users to build professional PowerPoint presentations through an intuitive interface.

Whether you're presenting research, pitching a business idea, or analyzing data — this app simplifies the process by offering:

- 📝 Slide building via forms or JSON
- 📊 Charts from manual input or CSV datasets
- 🎨 Theme styling with previews and presets
- 📥 PowerPoint file export
- 👁️ Real-time preview and editing
- 🔁 Undo/redo and drag-and-drop reordering
- 💾 Style import/export for team collaboration

---

## ✨ Key Features

### 🔧 Slide Creation Modes
- **Form-based input**: Add slides with titles, bullet points, images, and charts.
- **JSON import**: Paste your slide outline in JSON format.
- **Dataset charts**: Upload CSV files and generate charts with Plotly.

### 📊 Chart Types Supported
- Pie
- Bar
- Line
- Scatter (via Plotly dataset mode)

### 🎨 Theme Customization
- Choose from **style presets** (Professional, Minimalist, Creative) or create your own
- Customize:
  - Font family
  - Font size
  - Font color
  - Background (solid/gradient)
  - Slide layout and transition effects

### ✍️ Text Formatting (Per-Slide)
- Bold/italic toggles
- Font size overrides
- Alignment controls (left, center, right)

### ⚙️ Utility Tools
- **Undo/Redo**: Non-destructive editing with full history
- **Slide duplication**: Clone existing slides to speed up editing
- **Drag-and-drop reordering**: Rearrange slides using [`streamlit-sortables`](https://github.com/okld/streamlit-sortables)

### 🔄 Theme Export/Import
- Export your style config as JSON
- Reuse or share styles across sessions or teams

### 📤 PPTX Export
- Exports `.pptx` files with embedded images and charts using `python-pptx`
- Dataset charts are rendered using `plotly + kaleido`

---

## 📦 Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/ppt-generator.git
   cd ppt-generator
   ```

2. **Create a virtual environment (optional but recommended)**
   ```bash
   python -m venv venv
   source venv/bin/activate   # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the Streamlit app**
   ```bash
   streamlit run app.py
   ```

---

## ✅ Requirements

- Python 3.8+
- Streamlit
- plotly
- pandas
- matplotlib
- python-pptx
- kaleido (for chart image export)
- streamlit-sortables (for drag-and-drop)

Install them all using:
```bash
pip install -r requirements.txt
```

## 📤 Sample JSON Input

```json
[
  {
    "title": "Introduction",
    "content": ["What is the problem?", "Why does it matter?"],
    "chart": "pie",
    "chart_data": {"categories": ["A", "B"], "values": [60, 40]},
    "chart_input_type": "Manual",
    "transition": "Fade"
  },
  {
    "title": "Data-Driven Insights",
    "chart": "bar",
    "chart_input_type": "Dataset",
    "chart_data": {"x_col": "Region", "y_col": "Sales", "category_col": "Year"},
    "transition": "Wipe"
  }
]
```
