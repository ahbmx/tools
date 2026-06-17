Here's a comprehensive example that creates a professional PowerPoint and demonstrates many of the major features available in `python-pptx`:

* Title slide
* Agenda slide
* Bullet lists (multiple levels)
* Formatted text
* Tables
* Images
* Shapes
* Speaker notes
* Charts (using Excel-style embedded data)
* Section slides
* Custom colors
* Footer text
* Multiple slide layouts

Install dependencies:

```bash
pip install python-pptx
```

---

# Complete Example

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor

# --------------------------------------------------
# Create Presentation
# --------------------------------------------------

prs = Presentation()

# --------------------------------------------------
# Helper Functions
# --------------------------------------------------

def add_title_slide(title, subtitle):
    slide = prs.slides.add_slide(prs.slide_layouts[0])

    slide.shapes.title.text = title
    slide.placeholders[1].text = subtitle

    return slide


def add_section_slide(title):
    slide = prs.slides.add_slide(prs.slide_layouts[2])

    slide.shapes.title.text = title

    return slide


def add_bullet_slide(title, bullets):
    slide = prs.slides.add_slide(prs.slide_layouts[1])

    slide.shapes.title.text = title

    tf = slide.placeholders[1].text_frame
    tf.clear()

    for i, item in enumerate(bullets):

        if isinstance(item, tuple):
            text, level = item
        else:
            text = item
            level = 0

        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        p.text = text
        p.level = level

    return slide


# --------------------------------------------------
# Slide 1 - Title
# --------------------------------------------------

slide = add_title_slide(
    "Python PowerPoint Automation",
    "Comprehensive python-pptx Showcase"
)

# Speaker notes
slide.notes_slide.notes_text_frame.text = """
Introduce automation use cases.
Mention reporting pipelines.
"""

# --------------------------------------------------
# Slide 2 - Agenda
# --------------------------------------------------

add_bullet_slide(
    "Agenda",
    [
        "Text Formatting",
        "Tables",
        "Charts",
        "Images",
        "Shapes",
        "Speaker Notes",
        "Layouts"
    ]
)

# --------------------------------------------------
# Slide 3 - Formatted Text
# --------------------------------------------------

slide = prs.slides.add_slide(prs.slide_layouts[5])

slide.shapes.title.text = "Text Formatting"

textbox = slide.shapes.add_textbox(
    Inches(1),
    Inches(1.5),
    Inches(6),
    Inches(2)
)

tf = textbox.text_frame

p = tf.paragraphs[0]
run = p.add_run()

run.text = "Large Blue Heading"
run.font.size = Pt(28)
run.font.bold = True
run.font.color.rgb = RGBColor(0, 102, 204)

p2 = tf.add_paragraph()
p2.text = "Regular paragraph text"

p3 = tf.add_paragraph()
p3.text = "Indented bullet"
p3.level = 1

# --------------------------------------------------
# Slide 4 - Table
# --------------------------------------------------

slide = prs.slides.add_slide(prs.slide_layouts[5])

slide.shapes.title.text = "Sales Table"

rows = 5
cols = 3

table = slide.shapes.add_table(
    rows,
    cols,
    Inches(0.5),
    Inches(1.5),
    Inches(7),
    Inches(2)
).table

headers = ["Product", "Revenue", "Growth"]

for col, value in enumerate(headers):
    table.cell(0, col).text = value

data = [
    ("A", "$100k", "12%"),
    ("B", "$180k", "20%"),
    ("C", "$140k", "8%"),
    ("D", "$210k", "30%")
]

for row_idx, row_data in enumerate(data, start=1):
    for col_idx, value in enumerate(row_data):
        table.cell(row_idx, col_idx).text = value

# --------------------------------------------------
# Slide 5 - Chart
# --------------------------------------------------

slide = prs.slides.add_slide(prs.slide_layouts[5])

slide.shapes.title.text = "Quarterly Revenue"

chart_data = CategoryChartData()

chart_data.categories = [
    "Q1",
    "Q2",
    "Q3",
    "Q4"
]

chart_data.add_series(
    "Revenue",
    (120, 150, 180, 220)
)

chart = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(1),
    Inches(1.5),
    Inches(6),
    Inches(3),
    chart_data
).chart

chart.has_legend = False
chart.value_axis.has_major_gridlines = True

# --------------------------------------------------
# Slide 6 - Shapes
# --------------------------------------------------

slide = prs.slides.add_slide(prs.slide_layouts[5])

slide.shapes.title.text = "Shapes"

rect = slide.shapes.add_shape(
    MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
    Inches(1),
    Inches(1.5),
    Inches(2),
    Inches(1)
)

rect.text = "Process Start"

arrow = slide.shapes.add_shape(
    MSO_AUTO_SHAPE_TYPE.CHEVRON,
    Inches(3.2),
    Inches(1.5),
    Inches(1.5),
    Inches(1)
)

arrow.text = "Next"

circle = slide.shapes.add_shape(
    MSO_AUTO_SHAPE_TYPE.OVAL,
    Inches(5),
    Inches(1.5),
    Inches(1.5),
    Inches(1.5)
)

circle.text = "End"

# --------------------------------------------------
# Slide 7 - Process Diagram
# --------------------------------------------------

slide = prs.slides.add_slide(prs.slide_layouts[5])

slide.shapes.title.text = "Automation Workflow"

steps = [
    "Extract Data",
    "Transform",
    "Analyze",
    "Build PPT",
    "Distribute"
]

left = 0.5

for step in steps:

    box = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(left),
        Inches(2),
        Inches(1.3),
        Inches(0.7)
    )

    box.text = step

    left += 1.5

# --------------------------------------------------
# Slide 8 - Image Example
# --------------------------------------------------

slide = prs.slides.add_slide(prs.slide_layouts[5])

slide.shapes.title.text = "Image Example"

# Replace with your image
image_path = "example.png"

try:
    slide.shapes.add_picture(
        image_path,
        Inches(1),
        Inches(1.3),
        width=Inches(5)
    )
except:
    textbox = slide.shapes.add_textbox(
        Inches(1),
        Inches(1.5),
        Inches(4),
        Inches(1)
    )

    textbox.text = "Place an image at example.png"

# --------------------------------------------------
# Slide 9 - KPI Dashboard
# --------------------------------------------------

slide = prs.slides.add_slide(prs.slide_layouts[5])

slide.shapes.title.text = "Executive Dashboard"

metrics = [
    ("Revenue", "$1.2M"),
    ("Customers", "15,420"),
    ("Growth", "18%"),
    ("NPS", "72")
]

x = 0.5

for label, value in metrics:

    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(x),
        Inches(1.5),
        Inches(1.6),
        Inches(1.2)
    )

    shape.text = f"{label}\n{value}"

    x += 1.8

# --------------------------------------------------
# Slide 10 - Closing Slide
# --------------------------------------------------

slide = prs.slides.add_slide(prs.slide_layouts[2])

slide.shapes.title.text = "Questions?"

subtitle = slide.placeholders[1]
subtitle.text = "Thank you"

# --------------------------------------------------
# Save
# --------------------------------------------------

prs.save("comprehensive_demo.pptx")

print("Presentation created.")
```

---

## Features demonstrated

| Feature             | Included |
| ------------------- | -------- |
| Title slides        | ✓        |
| Section slides      | ✓        |
| Bullet lists        | ✓        |
| Multi-level bullets | ✓        |
| Text formatting     | ✓        |
| Colors              | ✓        |
| Font sizing         | ✓        |
| Tables              | ✓        |
| Charts              | ✓        |
| Shapes              | ✓        |
| Process diagrams    | ✓        |
| Images              | ✓        |
| KPI dashboard       | ✓        |
| Speaker notes       | ✓        |
| Multiple layouts    | ✓        |

For production reporting, you can extend this by pulling data from:

* `pandas` DataFrames
* SQL databases
* Excel files
* APIs
* AI-generated slide content

and automatically generate executive reports, sales decks, analytics presentations, or client-facing slide decks.
