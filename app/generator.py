"""
CIPETHUB Engine — Video Generation Pipeline V3.0
Generates professional educational lecture MP4 videos for CIPET students.

Upgrades in V3.0 (mobile-first redesign):
  1. Mobile-first LARGE fonts: 56px titles, 36px bullets, 28px minimum everywhere
  2. 12 unique per-slide layouts with different visual styles and accent colors
  3. Gemini Pro model with new V3 prompt focused purely on teaching concepts
  4. Reduced text (max 3 bullets per slide) — 70%+ visual content per slide
  5. Full-width diagrams: flowchart, bar chart, machine diagram, mind map, timeline etc.
  6. Per-slide accent color cycling through 12 distinct brand colors
  7. Zero repetition: CIPETHUB mentioned once in intro and once in outro only
  8. Ken Burns + fade transitions retained from V2.0
  9. Ambient background music retained from V2.0
 10. Progress bar retained from V2.0
"""

import array
import json
import math
import os
import re
import shutil
import subprocess
import wave as wave_module
from datetime import datetime
from pathlib import Path

from gtts import gTTS
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

DEPT_COLORS = {
    "Plastics":      RGBColor(0xFF, 0x6F, 0x00),
    "Mechanical":    RGBColor(0x21, 0x96, 0xF3),
    "Manufacturing": RGBColor(0x4C, 0xAF, 0x50),
    "All":           RGBColor(0xFF, 0xD5, 0x4F),
}

DEPT_COLORS_PIL = {
    "Plastics":      (0xFF, 0x6F, 0x00),
    "Mechanical":    (0x21, 0x96, 0xF3),
    "Manufacturing": (0x4C, 0xAF, 0x50),
    "All":           (0xFF, 0xD5, 0x4F),
}

BG_DARK      = RGBColor(0x0D, 0x1B, 0x2A)
CARD_BG      = RGBColor(0x15, 0x2A, 0x3D)
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY   = RGBColor(0xB0, 0xBE, 0xC5)
BOTTOM_BAR_BG = RGBColor(0x07, 0x11, 0x1A)

_BG_DARK    = (0x0D, 0x1B, 0x2A)
_CARD_BG    = (0x15, 0x2A, 0x3D)
_WHITE      = (0xFF, 0xFF, 0xFF)
_LIGHT_GRAY = (0xB0, 0xBE, 0xC5)
_BOTTOM_BG  = (0x07, 0x11, 0x1A)

_PX       = 120   # pixels per inch at 1920×1080 / 16×9 inches
_VIS_W    = 540   # visual canvas width  (fits inside right hint panel)
_VIS_H    = 648   # visual canvas height (fits inside right hint panel)

STDERR_TAIL = 500

# ---------------------------------------------------------------------------
# V3.0 Per-slide accent color palette (one distinct color per slide)
# ---------------------------------------------------------------------------

SLIDE_COLORS_PIL = [
    (0xFF, 0x6F, 0x00),  # 1:  Orange
    (0x21, 0x96, 0xF3),  # 2:  Blue
    (0x4C, 0xAF, 0x50),  # 3:  Green
    (0xE9, 0x1E, 0x63),  # 4:  Pink
    (0x9C, 0x27, 0xB0),  # 5:  Purple
    (0x00, 0xBC, 0xD4),  # 6:  Cyan
    (0xFF, 0x57, 0x22),  # 7:  Deep Orange
    (0x3F, 0x51, 0xB5),  # 8:  Indigo
    (0x00, 0x96, 0x88),  # 9:  Teal
    (0xFF, 0xC1, 0x07),  # 10: Amber
    (0x79, 0x55, 0x48),  # 11: Brown
    (0x60, 0x7D, 0x8B),  # 12: Blue Grey
]

# ---------------------------------------------------------------------------
# AI Script Generation
# ---------------------------------------------------------------------------


def ai_script(topic: str, department: str, video_type: str) -> dict:
    """Generate a structured lecture script using Google Gemini API."""
    api_key = os.environ.get("GEMINI_API_KEY", "")
    if not api_key or api_key == "your_key_here":
        print("[generator] No Gemini API key — using fallback script.")
        return _fallback(topic, department)

    try:
        import google.generativeai as genai

        genai.configure(api_key=api_key)
        print(f"[generator] Calling Gemini with key: {api_key[:10]}...")

        prompt = f"""You are creating a technical lecture script. Topic: {topic}, Department: {department}.

RULES:
- Generate exactly 12 slides
- Each slide: title (max 5 words), 3 bullet points (max 10 words each), narration (3 sentences, technical and specific), visual_description (detailed description of what diagram to draw)
- Include REAL numbers: temperatures, pressures, dimensions, speeds where applicable
- DO NOT repeat "CIPET", "CIPETHUB", "exam", "semester" — focus only on teaching the concept
- Slide 1: Brief intro to topic (1 sentence: "Today we learn about {topic}.", then start teaching immediately)
- Slides 2-11: Pure technical content with specific data
- Slide 12: Quick summary + "Subscribe for more" (1 sentence max)
- Narration style: Clear, direct, like an expert explaining to a colleague. NOT repetitive.
- Each bullet point: max 10 words, short and punchy, specific facts or values

Return ONLY valid JSON (no markdown, no code blocks):
{{
  "title": "video title here",
  "description": "video description here",
  "tags": ["tag1","tag2","tag3","tag4","tag5","tag6","tag7","tag8","tag9","tag10","tag11","tag12","tag13","tag14","tag15","tag16","tag17","tag18","tag19","tag20"],
  "slides": [
    {{
      "title": "slide title (max 5 words)",
      "bullets": ["bullet 1 (max 10 words)", "bullet 2 (max 10 words)", "bullet 3 (max 10 words)"],
      "narration": "3 sentences of technical narration.",
      "visual_description": "detailed description of diagram, chart, machine, or flowchart to render for this slide"
    }}
  ]
}}

Generate exactly 12 slides. The tags array must have exactly 20 items."""

        # Try gemini-pro first (user has Gemini Pro active), fall back to gemini-1.5-flash
        try:
            model = genai.GenerativeModel("gemini-pro")
            response = model.generate_content(prompt)
        except Exception as pro_exc:
            print(f"[generator] gemini-pro failed ({pro_exc}), trying gemini-1.5-flash ...")
            model = genai.GenerativeModel("gemini-1.5-flash")
            response = model.generate_content(prompt)

        raw = response.text.strip()
        print(f"[generator] Gemini response length: {len(raw)} chars")

        # Strip markdown code fences if present
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        data = json.loads(raw)

        # Validate and normalise
        slides = data.get("slides", [])
        if len(slides) < 12:
            slides += _fallback(topic, department)["slides"][len(slides):]
        data["slides"] = slides[:12]

        tags = data.get("tags", [])
        if len(tags) < 20:
            tags += [topic, department, "engineering", "lecture"] * 5
        data["tags"] = tags[:20]

        return data

    except Exception as exc:
        import traceback
        print(f"[generator] Gemini error ({type(exc).__name__}): {exc}")
        traceback.print_exc()
        return _fallback(topic, department)


def _fallback(topic: str, department: str) -> dict:
    """Return a 12-slide template script when Gemini is unavailable."""

    slides = [
        {
            "title": f"Intro: {topic[:25]}",
            "bullets": [
                f"{topic} — core concept overview",
                f"Key applications in {department} engineering",
                "Why this topic matters today",
            ],
            "narration": (
                f"Today we learn about {topic}. "
                f"This is a fundamental concept in {department} engineering with wide industrial applications. "
                "We will cover principles, processes, and real-world examples in this lecture."
            ),
            "visual_description": f"Labeled diagram introducing {topic} with key components highlighted",
        },
        {
            "title": "Fundamental Concepts",
            "bullets": [
                "Core definitions and terminology",
                "Underlying scientific principles",
                "Standard units and measurements",
            ],
            "narration": (
                f"The fundamental concepts of {topic} start with precise definitions and terminology. "
                "Understanding the scientific basis helps us predict material and process behaviour accurately. "
                "Standard units ensure consistent measurement across all industrial applications."
            ),
            "visual_description": "Concept map showing fundamental definitions, relationships and key terms with connecting arrows",
        },
        {
            "title": "Classification & Types",
            "bullets": [
                "Primary classification criteria",
                "Type A vs Type B vs Type C",
                "Selection guide by application",
            ],
            "narration": (
                f"Classification of {topic} is based on structural differences and performance properties. "
                "Each type has distinct advantages that determine its suitability for specific applications. "
                "Knowing these differences is essential for correct material or process selection."
            ),
            "visual_description": "Three side-by-side comparison cards showing Type A, Type B, Type C with key properties listed in each card",
        },
        {
            "title": "Properties & Characteristics",
            "bullets": [
                "Mechanical: strength, stiffness, toughness",
                "Thermal: Tg, HDT, conductivity values",
                "Chemical resistance and durability",
            ],
            "narration": (
                f"The properties of {topic} span mechanical, thermal, and chemical domains. "
                "Tensile strength, heat deflection temperature, and chemical resistance are the most important parameters. "
                "These values directly determine service life and product quality."
            ),
            "visual_description": "Bar chart comparing key property values (strength, temperature resistance, chemical resistance) across variants",
        },
        {
            "title": "Manufacturing Process",
            "bullets": [
                "Raw material → processing → product",
                "Critical temperature and pressure ranges",
                "Quality control checkpoints",
            ],
            "narration": (
                f"The manufacturing process for {topic} begins with raw material preparation and conditioning. "
                "Precise control of temperature, pressure, and time at each step is critical for quality. "
                "Quality checkpoints at each stage prevent defects and ensure consistency."
            ),
            "visual_description": "Step-by-step flowchart: Raw Material → Preparation → Processing → Cooling → Inspection → Product, with temperature/pressure values at each step",
        },
        {
            "title": "Equipment & Machinery",
            "bullets": [
                "Main machine: components and function",
                "Drive system and control unit",
                "Safety interlocks and sensors",
            ],
            "narration": (
                f"The equipment used in {topic} consists of a main processing unit, drive system, and control instrumentation. "
                "Each component performs a specific function that contributes to the overall process efficiency. "
                "Safety interlocks and sensors protect operators and ensure consistent product quality."
            ),
            "visual_description": "Labeled machine cross-section diagram showing Hopper, Barrel, Screw, Heater Bands, Nozzle, Mould/Die, Drive Motor with arrows indicating material flow",
        },
        {
            "title": "Design Calculations",
            "bullets": [
                "Key formulas with variable definitions",
                "Sample calculation: given → find → solve",
                "Safety factor: typically 1.5–3.0×",
            ],
            "narration": (
                f"Design calculations for {topic} require applying fundamental formulas with correct parameter values. "
                "A worked example clarifies how to substitute known values and solve for the required quantity. "
                "Always apply a safety factor of 1.5 to 3.0 as per applicable engineering standards."
            ),
            "visual_description": "Left: Key formula in large text with labelled variables. Right: Worked example table with given values, formula substitution, and numerical result",
        },
        {
            "title": "Industrial Applications",
            "bullets": [
                "Packaging and consumer goods",
                "Automotive and aerospace components",
                "Construction and infrastructure",
            ],
            "narration": (
                f"{topic} is applied across packaging, automotive, and construction sectors worldwide. "
                "In packaging, it provides lightweight, barrier-rich solutions at low cost. "
                "Automotive and aerospace applications demand high strength-to-weight ratio and thermal stability."
            ),
            "visual_description": "2x2 grid of application cards: Packaging (with icon), Automotive (with icon), Construction (with icon), Aerospace (with icon) — each card with 1 key fact",
        },
        {
            "title": "Quality Standards",
            "bullets": [
                "IS / ISO / ASTM standard references",
                "Test methods: tensile, impact, thermal",
                "Acceptance criteria and tolerances",
            ],
            "narration": (
                f"Quality standards for {topic} are defined by IS, ISO, and ASTM specifications. "
                "Testing methods cover tensile strength, impact resistance, and thermal performance. "
                "Acceptance criteria set the minimum performance thresholds for commercial products."
            ),
            "visual_description": "Comparison table with alternating row colors: Standard | Test Method | Property Tested | Acceptance Limit — 5 rows of data",
        },
        {
            "title": "Environmental Aspects",
            "bullets": [
                "Recyclability and end-of-life options",
                "Energy consumption per kg produced",
                "Regulatory compliance requirements",
            ],
            "narration": (
                f"Environmental considerations for {topic} focus on recyclability, energy use, and regulatory compliance. "
                "Life-cycle analysis shows that energy-efficient processing reduces carbon footprint significantly. "
                "Compliance with environmental regulations is mandatory for all industrial operations."
            ),
            "visual_description": "Circular cycle diagram: Production → Use → Collection → Recycling → Back to Production, with labels for energy savings and CO2 reduction at each stage",
        },
        {
            "title": "Recent Trends",
            "bullets": [
                "Nanocomposites and smart materials",
                "Industry 4.0 and IoT integration",
                "Biobased and sustainable alternatives",
            ],
            "narration": (
                f"Recent advances in {topic} include nanocomposite reinforcement and smart material integration. "
                "Industry 4.0 enables real-time process monitoring using IoT sensors and AI-based control. "
                "Bio-based alternatives are gaining traction as sustainable replacements for conventional materials."
            ),
            "visual_description": "Horizontal timeline from 2000 to 2030: key milestones marked — nanotechnology (2005), bioplastics (2010), Industry 4.0 (2018), smart materials (2025+)",
        },
        {
            "title": "Summary",
            "bullets": [
                f"{topic} — key principles recap",
                "Critical values and formulas",
                "Subscribe for more lectures",
            ],
            "narration": (
                f"We have covered the complete fundamentals of {topic} from principles to applications. "
                "Remember the key property values, process parameters, and design formulas from this lecture. "
                "Subscribe for more technical lectures like this one."
            ),
            "visual_description": f"Mind map with '{topic}' at centre and branches: Principles, Types, Properties, Process, Equipment, Applications, Standards, Environment, Trends",
        },
    ]

    tags = [
        topic, department, "engineering", "lecture",
        "education", "tutorial", "technical", "manufacturing",
        f"{department} engineering", "study material",
        "polymer", "plastics", "materials", "process engineering",
        "industrial", "technology", "science", "India",
        "CIPETHUB", "YouTube lecture",
    ]

    return {
        "title": f"{topic} | {department} Engineering Lecture",
        "description": (
            f"Complete technical lecture on {topic} for {department} engineering students. "
            f"Covers fundamentals, processes, properties, applications, and design calculations. "
            f"#{topic.replace(' ', '')} #{department.replace(' ', '')} #engineering #lecture"
        ),
        "tags": tags[:20],
        "slides": slides,
    }


# ---------------------------------------------------------------------------
# PowerPoint Creation
# ---------------------------------------------------------------------------


def make_ppt(script: dict, department: str, work_dir: str) -> str:
    """Create a professional 16:9 widescreen PPT from the script."""
    dept_color = DEPT_COLORS.get(department, DEPT_COLORS["All"])
    prs = Presentation()
    prs.slide_width = Inches(16)
    prs.slide_height = Inches(9)

    blank_layout = prs.slide_layouts[6]  # completely blank

    for idx, slide_data in enumerate(script["slides"], start=1):
        slide = prs.slides.add_slide(blank_layout)

        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = BG_DARK

        top_bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(16), Inches(0.12))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = dept_color
        top_bar.line.fill.background()

        badge = slide.shapes.add_shape(1, Inches(0.2), Inches(0.18), Inches(2.0), Inches(0.38))
        badge.fill.solid()
        badge.fill.fore_color.rgb = dept_color
        badge.line.fill.background()
        badge_tf = badge.text_frame
        badge_tf.word_wrap = False
        badge_para = badge_tf.paragraphs[0]
        badge_para.alignment = 1
        run = badge_para.add_run()
        run.text = department.upper()
        run.font.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = WHITE

        brand_box = slide.shapes.add_textbox(Inches(12.5), Inches(0.15), Inches(3.3), Inches(0.55))
        brand_tf = brand_box.text_frame
        brand_para = brand_tf.paragraphs[0]
        brand_para.alignment = 2
        run = brand_para.add_run()
        run.text = "CIPETHUB"
        run.font.bold = True
        run.font.size = Pt(18)
        run.font.color.rgb = dept_color

        sub_para = brand_tf.add_paragraph()
        sub_para.alignment = 2
        sub_run = sub_para.add_run()
        sub_run.text = "CIPET Study Material"
        sub_run.font.size = Pt(9)
        sub_run.font.color.rgb = LIGHT_GRAY

        num_box = slide.shapes.add_textbox(Inches(14.8), Inches(8.4), Inches(1.0), Inches(0.4))
        num_tf = num_box.text_frame
        num_para = num_tf.paragraphs[0]
        num_para.alignment = 2
        run = num_para.add_run()
        run.text = str(idx)
        run.font.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = dept_color

        title_box = slide.shapes.add_textbox(Inches(0.3), Inches(0.65), Inches(10.5), Inches(0.75))
        title_tf = title_box.text_frame
        title_tf.word_wrap = False
        title_para = title_tf.paragraphs[0]
        run = title_para.add_run()
        run.text = slide_data.get("title", f"Slide {idx}")
        run.font.bold = True
        run.font.size = Pt(36)
        run.font.color.rgb = dept_color

        underline_bar = slide.shapes.add_shape(1, Inches(0.3), Inches(1.45), Inches(10.5), Inches(0.05))
        underline_bar.fill.solid()
        underline_bar.fill.fore_color.rgb = dept_color
        underline_bar.line.fill.background()

        card = slide.shapes.add_shape(1, Inches(0.3), Inches(1.55), Inches(10.5), Inches(6.1))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.fill.background()

        content_box = slide.shapes.add_textbox(Inches(0.45), Inches(1.65), Inches(10.2), Inches(5.9))
        content_tf = content_box.text_frame
        content_tf.word_wrap = True
        first = True
        for bullet in slide_data.get("bullets", []):
            if first:
                para = content_tf.paragraphs[0]
                first = False
            else:
                para = content_tf.add_paragraph()
            para.space_before = Pt(6)
            run = para.add_run()
            run.text = f"▸  {bullet}"
            run.font.size = Pt(22)
            run.font.color.rgb = WHITE

        hint_box_bg = slide.shapes.add_shape(1, Inches(11.0), Inches(1.55), Inches(4.7), Inches(6.1))
        hint_box_bg.fill.solid()
        hint_box_bg.fill.fore_color.rgb = CARD_BG
        hint_box_bg.line.color.rgb = dept_color
        hint_box_bg.line.width = Pt(1)

        hint_label = slide.shapes.add_textbox(Inches(11.05), Inches(1.65), Inches(4.6), Inches(0.4))
        hl_tf = hint_label.text_frame
        hl_para = hl_tf.paragraphs[0]
        hl_para.alignment = 1
        run = hl_para.add_run()
        run.text = "VISUAL DIAGRAM"
        run.font.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = dept_color

        hint_text_box = slide.shapes.add_textbox(Inches(11.05), Inches(2.1), Inches(4.6), Inches(5.4))
        ht_tf = hint_text_box.text_frame
        ht_tf.word_wrap = True
        ht_para = ht_tf.paragraphs[0]
        run = ht_para.add_run()
        run.text = slide_data.get("visual_hint", "Diagram here")
        run.font.size = Pt(14)
        run.font.color.rgb = LIGHT_GRAY

        bottom_bar = slide.shapes.add_shape(1, Inches(0), Inches(8.7), Inches(16), Inches(0.3))
        bottom_bar.fill.solid()
        bottom_bar.fill.fore_color.rgb = BOTTOM_BAR_BG
        bottom_bar.line.fill.background()

        bottom_text = slide.shapes.add_textbox(Inches(0), Inches(8.7), Inches(16), Inches(0.3))
        bt_tf = bottom_text.text_frame
        bt_para = bt_tf.paragraphs[0]
        bt_para.alignment = 1
        run = bt_para.add_run()
        run.text = (
            f"CIPETHUB  •  {department} Engineering  •  CIPET Study Material  •  Subscribe & Like!"
        )
        run.font.size = Pt(10)
        run.font.color.rgb = LIGHT_GRAY

    ppt_path = os.path.join(work_dir, "lecture.pptx")
    prs.save(ppt_path)
    print(f"[generator] PPT saved: {ppt_path}")
    return ppt_path


# ---------------------------------------------------------------------------
# Font / Text helpers
# ---------------------------------------------------------------------------


def _load_font(bold: bool = False, size: int = 20) -> ImageFont.ImageFont:
    """Load DejaVu Sans from disk, falling back to PIL's built-in font."""
    candidates = (
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold
        else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf" if bold
        else "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    )
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
    return ImageFont.load_default()


def _draw_wrapped_text(
    draw: ImageDraw.ImageDraw,
    text: str,
    font: ImageFont.ImageFont,
    x: int,
    y: int,
    max_width: int,
    fill: tuple,
    line_spacing: int = 6,
) -> int:
    """Draw word-wrapped text and return the y position after the last line."""
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        test = f"{current} {word}".strip()
        bbox = draw.textbbox((0, 0), test, font=font)
        if bbox[2] - bbox[0] <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)

    for line in lines:
        draw.text((x, y), line, font=font, fill=fill)
        bbox = draw.textbbox((0, 0), line, font=font)
        y += (bbox[3] - bbox[1]) + line_spacing
    return y


# ---------------------------------------------------------------------------
# Visual diagram renderers (Upgrade 1 + 8)
# ---------------------------------------------------------------------------

# Palette of contrasting colours used in charts
_CHART_COLORS = [
    (0xFF, 0x6F, 0x00),  # orange
    (0x21, 0x96, 0xF3),  # blue
    (0x4C, 0xAF, 0x50),  # green
    (0xFF, 0xD5, 0x4F),  # yellow
    (0xE9, 0x1E, 0x63),  # pink
    (0x9C, 0x27, 0xB0),  # purple
    (0x00, 0xBC, 0xD4),  # cyan
]


def _render_flowchart(items: list, dept_color: tuple) -> Image.Image:
    """Render a vertical flowchart: boxes connected by arrows."""
    img = Image.new("RGB", (_VIS_W, _VIS_H), _CARD_BG)
    draw = ImageDraw.Draw(img)
    steps = [str(s)[:40] for s in items[:7]] if items else ["Step 1", "Step 2", "Step 3"]
    font = _load_font(bold=False, size=15)

    box_h = 54
    arrow_h = 22
    pad_x = 24
    box_w = _VIS_W - 2 * pad_x
    total_h = len(steps) * box_h + (len(steps) - 1) * arrow_h
    start_y = max(14, (_VIS_H - total_h) // 2)

    shades = [
        dept_color,
        tuple(max(0, c - 25) for c in dept_color),
        tuple(min(255, c + 20) for c in dept_color),
    ]

    for idx, step in enumerate(steps):
        y = start_y + idx * (box_h + arrow_h)
        color = shades[idx % len(shades)]
        draw.rectangle([pad_x, y, pad_x + box_w, y + box_h],
                       fill=color, outline=_WHITE, width=1)
        bb = draw.textbbox((0, 0), step, font=font)
        tx = pad_x + (box_w - (bb[2] - bb[0])) // 2
        ty = y + (box_h - (bb[3] - bb[1])) // 2
        draw.text((tx, ty), step, font=font, fill=_WHITE)

        if idx < len(steps) - 1:
            ax = _VIS_W // 2
            ay1 = y + box_h
            ay2 = y + box_h + arrow_h
            draw.line([(ax, ay1), (ax, ay2 - 8)], fill=dept_color, width=3)
            draw.polygon([(ax, ay2), (ax - 8, ay2 - 12), (ax + 8, ay2 - 12)], fill=dept_color)

    return img


def _render_comparison_table(items: list, dept_color: tuple) -> Image.Image:
    """Render a two-column comparison table."""
    img = Image.new("RGB", (_VIS_W, _VIS_H), _CARD_BG)
    draw = ImageDraw.Draw(img)
    header_font = _load_font(bold=True, size=16)
    cell_font   = _load_font(bold=False, size=14)
    col_w = (_VIS_W - 20) // 2
    row_h = 56
    header_y = 14

    # Header row
    draw.rectangle([10, header_y, 10 + col_w, header_y + row_h], fill=dept_color)
    draw.rectangle([10 + col_w, header_y, _VIS_W - 10, header_y + row_h],
                   fill=tuple(max(0, c - 30) for c in dept_color))
    draw.text((18, header_y + 16), "Property", font=header_font, fill=_WHITE)
    draw.text((18 + col_w, header_y + 16), "Details", font=header_font, fill=_WHITE)

    alt_bg = (0x1E, 0x38, 0x50)
    for i, item in enumerate(items[:8]):
        y = header_y + (i + 1) * row_h
        if y + row_h > _VIS_H - 6:
            break
        bg = alt_bg if i % 2 == 0 else _CARD_BG
        draw.rectangle([10, y, _VIS_W - 10, y + row_h],
                       fill=bg, outline=(0x28, 0x48, 0x68), width=1)
        if ":" in item:
            parts = item.split(":", 1)
            label, value = parts[0].strip()[:22], parts[1].strip()[:30]
        else:
            label = item[:22]
            value = "—"
        _draw_wrapped_text(draw, label, cell_font, 16, y + 10, col_w - 12, dept_color, 3)
        _draw_wrapped_text(draw, value, cell_font, 16 + col_w, y + 10, col_w - 12, _WHITE, 3)

    return img


def _render_bar_chart(items: list, dept_color: tuple) -> Image.Image:
    """Render a vertical bar chart."""
    img = Image.new("RGB", (_VIS_W, _VIS_H), _CARD_BG)
    draw = ImageDraw.Draw(img)
    labels = [str(s)[:14] for s in items[:6]] if items else ["A", "B", "C", "D"]
    n = len(labels)
    font  = _load_font(bold=False, size=13)
    afont = _load_font(bold=False, size=12)

    ml, mr, mt, mb = 52, 16, 20, 80
    chart_w = _VIS_W - ml - mr
    chart_h = _VIS_H - mt - mb

    ax = ml; ay_bot = _VIS_H - mb; ax_r = _VIS_W - mr; ay_top = mt
    draw.line([(ax, ay_top), (ax, ay_bot)], fill=_LIGHT_GRAY, width=2)
    draw.line([(ax, ay_bot), (ax_r, ay_bot)], fill=_LIGHT_GRAY, width=2)

    # Grid lines
    for pct in [25, 50, 75, 100]:
        gy = ay_bot - int(chart_h * pct / 100)
        draw.line([(ax, gy), (ax_r, gy)], fill=(0x28, 0x42, 0x58), width=1)
        draw.text((4, gy - 8), f"{pct}%", font=afont, fill=_LIGHT_GRAY)

    bar_w = max(28, (chart_w - 10 * n) // n)
    heights = [0.45 + 0.55 * ((n - i) / n) for i in range(n)]

    for i, (lbl, h) in enumerate(zip(labels, heights)):
        bx = ax + 8 + i * (bar_w + 10)
        bh = int(chart_h * h)
        by = ay_bot - bh
        color = _CHART_COLORS[i % len(_CHART_COLORS)]
        draw.rectangle([bx, by, bx + bar_w, ay_bot], fill=color)
        draw.text((bx + bar_w // 2 - 12, by - 18), f"{int(h*100)}%", font=afont, fill=_WHITE)
        lb = draw.textbbox((0, 0), lbl, font=font)
        lx = bx + (bar_w - (lb[2] - lb[0])) // 2
        draw.text((lx, ay_bot + 6), lbl, font=font, fill=_LIGHT_GRAY)

    return img


def _render_pie_chart(items: list, dept_color: tuple) -> Image.Image:
    """Render a pie chart with coloured segments and labels."""
    img = Image.new("RGB", (_VIS_W, _VIS_H), _CARD_BG)
    draw = ImageDraw.Draw(img)
    labels = [str(s)[:28] for s in items[:7]] if items else ["Part A", "Part B", "Part C"]
    n = len(labels)
    font = _load_font(bold=False, size=13)

    cx, cy, r = _VIS_W // 2, (_VIS_H - 80) // 2 + 10, min(_VIS_W, _VIS_H - 80) // 2 - 20
    slices = [360 / n] * n
    start = -90
    for i, (lbl, deg) in enumerate(zip(labels, slices)):
        color = _CHART_COLORS[i % len(_CHART_COLORS)]
        end = start + deg
        draw.pieslice([cx - r, cy - r, cx + r, cy + r], start, end, fill=color, outline=_CARD_BG)

        # Label at midpoint angle
        mid_rad = math.radians((start + end) / 2)
        lx = int(cx + (r * 0.68) * math.cos(mid_rad))
        ly = int(cy + (r * 0.68) * math.sin(mid_rad))
        bb = draw.textbbox((0, 0), f"{int(deg)}%", font=font)
        draw.text((lx - (bb[2] - bb[0]) // 2, ly - (bb[3] - bb[1]) // 2),
                  f"{int(deg)}%", font=font, fill=_WHITE)
        start = end

    # Legend
    leg_y = cy + r + 12
    lfont = _load_font(bold=False, size=12)
    for i, lbl in enumerate(labels):
        color = _CHART_COLORS[i % len(_CHART_COLORS)]
        lx = 8 + (i % 3) * (_VIS_W // 3)
        draw.rectangle([lx, leg_y + (i // 3) * 22, lx + 14, leg_y + (i // 3) * 22 + 14], fill=color)
        draw.text((lx + 18, leg_y + (i // 3) * 22), lbl[:18], font=lfont, fill=_LIGHT_GRAY)

    return img


def _render_mind_map(items: list, dept_color: tuple, title: str) -> Image.Image:
    """Render a mind-map: central node with radiating branches."""
    img = Image.new("RGB", (_VIS_W, _VIS_H), _CARD_BG)
    draw = ImageDraw.Draw(img)
    font_c  = _load_font(bold=True, size=15)
    font_b  = _load_font(bold=False, size=13)
    branches = [str(s)[:30] for s in items[:8]] if items else ["Topic A", "Topic B", "Topic C"]
    n = len(branches)

    cx, cy = _VIS_W // 2, _VIS_H // 2
    cr = 52  # central circle radius
    br = 48  # branch node radius
    arm_len = min(150, int(min(_VIS_W, _VIS_H) * 0.34))

    # Central circle
    draw.ellipse([cx - cr, cy - cr, cx + cr, cy + cr], fill=dept_color)
    short_title = (title[:12] + "..") if len(title) > 14 else title
    tb = draw.textbbox((0, 0), short_title, font=font_c)
    draw.text((cx - (tb[2] - tb[0]) // 2, cy - (tb[3] - tb[1]) // 2),
              short_title, font=font_c, fill=_WHITE)

    for i, branch in enumerate(branches):
        angle = math.radians(-90 + i * 360 / n)
        bx = int(cx + arm_len * math.cos(angle))
        by = int(cy + arm_len * math.sin(angle))
        color = _CHART_COLORS[i % len(_CHART_COLORS)]

        # Line from center to branch
        draw.line([(cx, cy), (bx, by)], fill=color, width=2)

        # Branch node
        draw.ellipse([bx - br, by - br // 2, bx + br, by + br // 2],
                     fill=color, outline=_WHITE, width=1)
        bb = draw.textbbox((0, 0), branch[:16], font=font_b)
        draw.text((bx - (bb[2] - bb[0]) // 2, by - (bb[3] - bb[1]) // 2),
                  branch[:16], font=font_b, fill=_WHITE)

    return img


def _render_timeline(items: list, dept_color: tuple) -> Image.Image:
    """Render a horizontal timeline with events marked."""
    img = Image.new("RGB", (_VIS_W, _VIS_H), _CARD_BG)
    draw = ImageDraw.Draw(img)
    events = [str(s)[:30] for s in items[:7]] if items else ["Event 1", "Event 2", "Event 3"]
    n = len(events)
    font = _load_font(bold=False, size=13)
    tfont = _load_font(bold=True, size=14)

    line_y = _VIS_H // 2
    pad_x  = 40
    avail  = _VIS_W - 2 * pad_x

    # Main axis
    draw.line([(pad_x, line_y), (_VIS_W - pad_x, line_y)], fill=dept_color, width=3)
    # Arrow at end
    draw.polygon([(_VIS_W - pad_x + 10, line_y),
                  (_VIS_W - pad_x, line_y - 6),
                  (_VIS_W - pad_x, line_y + 6)], fill=dept_color)

    for i, event in enumerate(events):
        ex = pad_x + int(avail * i / max(n - 1, 1))
        color = _CHART_COLORS[i % len(_CHART_COLORS)]

        # Dot
        draw.ellipse([ex - 8, line_y - 8, ex + 8, line_y + 8], fill=color, outline=_WHITE, width=1)

        # Label alternating above/below
        label = event[:22]
        bb = draw.textbbox((0, 0), label, font=font)
        lw = bb[2] - bb[0]
        lx = max(2, min(ex - lw // 2, _VIS_W - lw - 4))
        if i % 2 == 0:
            draw.text((lx, line_y - 36 - (bb[3] - bb[1])), label, font=font, fill=_LIGHT_GRAY)
            draw.line([(ex, line_y - 10), (ex, line_y - 26)], fill=color, width=1)
        else:
            draw.text((lx, line_y + 26), label, font=font, fill=_LIGHT_GRAY)
            draw.line([(ex, line_y + 10), (ex, line_y + 26)], fill=color, width=1)

        # Step number
        num = str(i + 1)
        nb = draw.textbbox((0, 0), num, font=tfont)
        draw.text((ex - (nb[2] - nb[0]) // 2, line_y - (nb[3] - nb[1]) // 2),
                  num, font=tfont, fill=_WHITE)

    return img


def _render_tree_diagram(items: list, dept_color: tuple, title: str) -> Image.Image:
    """Render a simple tree diagram: one root, N children."""
    img = Image.new("RGB", (_VIS_W, _VIS_H), _CARD_BG)
    draw = ImageDraw.Draw(img)
    children = [str(s)[:28] for s in items[:6]] if items else ["Branch A", "Branch B", "Branch C"]
    n = len(children)
    font = _load_font(bold=False, size=14)
    rfont = _load_font(bold=True, size=15)

    root_text = (title[:16] + "..") if len(title) > 18 else title
    rw, rh = 180, 50
    rx = (_VIS_W - rw) // 2
    ry = 20

    draw.rectangle([rx, ry, rx + rw, ry + rh], fill=dept_color, outline=_WHITE, width=1)
    rb = draw.textbbox((0, 0), root_text, font=rfont)
    draw.text((rx + (rw - (rb[2] - rb[0])) // 2, ry + (rh - (rb[3] - rb[1])) // 2),
              root_text, font=rfont, fill=_WHITE)

    cw, ch = 140, 48
    spacing = max(cw + 10, (_VIS_W - 20) // max(n, 1))
    child_y = ry + rh + 90

    for i, child in enumerate(children):
        cx = 10 + i * spacing + (spacing - cw) // 2
        color = _CHART_COLORS[i % len(_CHART_COLORS)]
        if child_y + ch > _VIS_H - 10:
            break
        # Connection line
        draw.line([(rx + rw // 2, ry + rh), (cx + cw // 2, child_y)],
                  fill=color, width=2)
        draw.rectangle([cx, child_y, cx + cw, child_y + ch],
                       fill=color, outline=_WHITE, width=1)
        cb = draw.textbbox((0, 0), child[:18], font=font)
        draw.text((cx + (cw - (cb[2] - cb[0])) // 2, child_y + (ch - (cb[3] - cb[1])) // 2),
                  child[:18], font=font, fill=_WHITE)

    return img


def _render_formula_sheet(items: list, dept_color: tuple) -> Image.Image:
    """Render a formula sheet: each item shown as a labelled formula row."""
    img = Image.new("RGB", (_VIS_W, _VIS_H), _CARD_BG)
    draw = ImageDraw.Draw(img)
    font  = _load_font(bold=False, size=15)
    hfont = _load_font(bold=True, size=16)
    formulae = [str(s)[:50] for s in items[:8]] if items else ["F = ma", "E = mc^2"]

    draw.rectangle([10, 10, _VIS_W - 10, 44], fill=dept_color)
    draw.text((16, 16), "Key Formulas / Parameters", font=hfont, fill=_WHITE)

    row_h = 62
    for i, item in enumerate(formulae):
        y = 54 + i * row_h
        if y + row_h > _VIS_H - 6:
            break
        bg = (0x1E, 0x38, 0x50) if i % 2 == 0 else _CARD_BG
        draw.rectangle([10, y, _VIS_W - 10, y + row_h],
                       fill=bg, outline=(0x28, 0x45, 0x62), width=1)
        draw.rectangle([10, y, 28, y + row_h], fill=dept_color)
        num_b = draw.textbbox((0, 0), str(i + 1), font=font)
        draw.text((19 - (num_b[2] - num_b[0]) // 2, y + (row_h - (num_b[3] - num_b[1])) // 2),
                  str(i + 1), font=font, fill=_WHITE)
        _draw_wrapped_text(draw, item, font, 36, y + 10, _VIS_W - 48, _WHITE, 3)

    return img


def _render_labeled_diagram(items: list, dept_color: tuple, title: str) -> Image.Image:
    """Render a labeled diagram: central box with labeled callouts around it."""
    img = Image.new("RGB", (_VIS_W, _VIS_H), _CARD_BG)
    draw = ImageDraw.Draw(img)
    parts = [str(s)[:24] for s in items[:8]] if items else ["Part A", "Part B", "Part C"]
    n = len(parts)
    font  = _load_font(bold=False, size=13)
    cfont = _load_font(bold=True, size=15)

    cx, cy, cr = _VIS_W // 2, _VIS_H // 2, 66
    draw.ellipse([cx - cr, cy - cr, cx + cr, cy + cr], fill=dept_color, outline=_WHITE, width=2)
    short = (title[:10] + "..") if len(title) > 12 else title
    tb = draw.textbbox((0, 0), short, font=cfont)
    draw.text((cx - (tb[2] - tb[0]) // 2, cy - (tb[3] - tb[1]) // 2),
              short, font=cfont, fill=_WHITE)

    arm = int(min(_VIS_W, _VIS_H) * 0.36)
    bw, bh = 120, 36
    for i, part in enumerate(parts):
        angle = math.radians(-90 + i * 360 / n)
        bx = int(cx + arm * math.cos(angle))
        by = int(cy + arm * math.sin(angle))
        color = _CHART_COLORS[i % len(_CHART_COLORS)]
        draw.line([(cx, cy), (bx, by)], fill=color, width=2)

        rx = max(4, min(bx - bw // 2, _VIS_W - bw - 4))
        ry = max(4, min(by - bh // 2, _VIS_H - bh - 4))
        draw.rectangle([rx, ry, rx + bw, ry + bh], fill=color, outline=_WHITE, width=1)
        pb = draw.textbbox((0, 0), part[:16], font=font)
        draw.text((rx + (bw - (pb[2] - pb[0])) // 2, ry + (bh - (pb[3] - pb[1])) // 2),
                  part[:16], font=font, fill=_WHITE)

    return img


def render_visual(
    visual_hint: str,
    slide_data: dict,
    department: str,
    slide_index: int,
) -> Image.Image:
    """Choose and render the appropriate diagram based on *visual_hint* keywords.

    Returns a Pillow Image sized (_VIS_W × _VIS_H).
    """
    hint  = visual_hint.lower()
    items = slide_data.get("bullets", [])
    title = slide_data.get("title", "")
    dept_color = DEPT_COLORS_PIL.get(department, DEPT_COLORS_PIL["All"])

    # NOTE: More-specific multi-word keywords are checked before broader single words
    # to avoid false matches (e.g. "bar chart" before "comparison", "process flow"
    # before "process", "labelled diagram" before "flow").
    if any(kw in hint for kw in ["bar chart", "bar graph", "property comparison",
                                  "strength", "infographic", "showing industries"]):
        return _render_bar_chart(items, dept_color)
    if any(kw in hint for kw in ["pie", "composition", "percentage", "breakdown"]):
        return _render_pie_chart(items, dept_color)
    if any(kw in hint for kw in ["mind map", "summarising", "concept map", "summary"]):
        return _render_mind_map(items, dept_color, title)
    if any(kw in hint for kw in ["timeline", "history", "evolution", "development", "milestones"]):
        return _render_timeline(items, dept_color)
    if any(kw in hint for kw in ["tree", "classification", "hierarchy", "types", "showing all"]):
        return _render_tree_diagram(items, dept_color, title)
    if any(kw in hint for kw in ["formula", "calculation", "equation", "sheet", "variables",
                                  "key parameters"]):
        return _render_formula_sheet(items, dept_color)
    if any(kw in hint for kw in ["labelled diagram", "labeled diagram", "parts", "components"]):
        return _render_labeled_diagram(items, dept_color, title)
    if any(kw in hint for kw in ["process flow", "flowchart", "flow diagram",
                                  "step-by-step", "sequence", "stages"]):
        return _render_flowchart(items, dept_color)
    if any(kw in hint for kw in ["comparison", "table", "compare", "standard", "test method",
                                  "listing", "is/iso"]):
        return _render_comparison_table(items, dept_color)
    # default: labelled diagram
    return _render_labeled_diagram(items, dept_color, title)


# ---------------------------------------------------------------------------
# V3.0 Full-slide visual draw functions
# Each draws INTO a rectangular region (rx, ry, rw, rh) of the main canvas.
# All text: minimum 28px. These are used by render_slide_image_v3().
# ---------------------------------------------------------------------------


def _v3_draw_flowchart(
    draw: ImageDraw.ImageDraw,
    items: list,
    color: tuple,
    rx: int, ry: int, rw: int, rh: int,
) -> None:
    """Draw a vertical flowchart with colored boxes and arrows."""
    steps = [str(s)[:50] for s in items[:7]] if items else ["Step 1", "Step 2", "Step 3"]
    n = len(steps)
    box_h = min(100, (rh - 20) // (n + (n - 1) // 2))
    arrow_h = max(20, box_h // 2)
    total_h = n * box_h + (n - 1) * arrow_h
    start_y = ry + (rh - total_h) // 2
    pad_x = rx + 40
    box_w = rw - 80
    font = _load_font(bold=True, size=28)
    shades = [
        color,
        tuple(max(0, c - 30) for c in color),
        tuple(min(255, c + 25) for c in color),
    ]
    for i, step in enumerate(steps):
        y = start_y + i * (box_h + arrow_h)
        c = shades[i % len(shades)]
        draw.rectangle([pad_x, y, pad_x + box_w, y + box_h],
                       fill=c, outline=_WHITE, width=2)
        bb = draw.textbbox((0, 0), step, font=font)
        tx = pad_x + (box_w - (bb[2] - bb[0])) // 2
        ty = y + (box_h - (bb[3] - bb[1])) // 2
        draw.text((tx, ty), step, font=font, fill=_WHITE)
        if i < n - 1:
            ax = rx + rw // 2
            ay1 = y + box_h
            ay2 = y + box_h + arrow_h
            draw.line([(ax, ay1), (ax, ay2 - 10)], fill=color, width=4)
            draw.polygon([(ax, ay2), (ax - 12, ay2 - 16), (ax + 12, ay2 - 16)], fill=color)


def _v3_draw_bar_chart(
    draw: ImageDraw.ImageDraw,
    items: list,
    color: tuple,
    rx: int, ry: int, rw: int, rh: int,
) -> None:
    """Draw a large vertical bar chart with labeled axes."""
    labels = [str(s)[:18] for s in items[:6]] if items else ["A", "B", "C", "D"]
    n = len(labels)
    lbl_font = _load_font(bold=False, size=28)
    num_font = _load_font(bold=True, size=28)
    axis_font = _load_font(bold=False, size=24)

    ml, mr, mt, mb = rx + 60, rx + rw - 20, ry + 20, ry + rh - 80
    chart_w = mr - ml
    chart_h = mb - mt

    # Axes
    draw.line([(ml, mt), (ml, mb)], fill=_LIGHT_GRAY, width=3)
    draw.line([(ml, mb), (mr, mb)], fill=_LIGHT_GRAY, width=3)

    # Grid lines
    for pct in [25, 50, 75, 100]:
        gy = mb - int(chart_h * pct / 100)
        draw.line([(ml, gy), (mr, gy)], fill=(0x28, 0x42, 0x58), width=1)
        draw.text((ml - 56, gy - 14), f"{pct}%", font=axis_font, fill=_LIGHT_GRAY)

    bar_w = max(40, (chart_w - 20 * n) // n)
    heights = [0.45 + 0.55 * ((n - i) / n) for i in range(n)]
    for i, (lbl, h) in enumerate(zip(labels, heights)):
        bx = ml + 12 + i * (bar_w + 20)
        bh = int(chart_h * h)
        by = mb - bh
        c = _CHART_COLORS[i % len(_CHART_COLORS)]
        draw.rectangle([bx, by, bx + bar_w, mb], fill=c)
        val_text = f"{int(h * 100)}%"
        vb = draw.textbbox((0, 0), val_text, font=num_font)
        draw.text((bx + (bar_w - (vb[2] - vb[0])) // 2, by - 36),
                  val_text, font=num_font, fill=_WHITE)
        lb = draw.textbbox((0, 0), lbl, font=lbl_font)
        lx = bx + (bar_w - (lb[2] - lb[0])) // 2
        draw.text((lx, mb + 8), lbl, font=lbl_font, fill=_LIGHT_GRAY)


def _v3_draw_labeled_diagram(
    draw: ImageDraw.ImageDraw,
    items: list,
    color: tuple,
    title: str,
    rx: int, ry: int, rw: int, rh: int,
) -> None:
    """Draw a central labeled diagram with radiating callout boxes."""
    parts = [str(s)[:30] for s in items[:8]] if items else ["Part A", "Part B", "Part C"]
    n = len(parts)
    cfont = _load_font(bold=True, size=32)
    bfont = _load_font(bold=False, size=28)

    cx = rx + rw // 2
    cy = ry + rh // 2
    cr = min(rw, rh) // 6
    arm = min(rw, rh) // 3

    draw.ellipse([cx - cr, cy - cr, cx + cr, cy + cr], fill=color, outline=_WHITE, width=3)
    short = (title[:14] + "..") if len(title) > 16 else title
    tb = draw.textbbox((0, 0), short, font=cfont)
    draw.text((cx - (tb[2] - tb[0]) // 2, cy - (tb[3] - tb[1]) // 2),
              short, font=cfont, fill=_WHITE)

    bw = max(160, rw // 5)
    bh = 54
    for i, part in enumerate(parts):
        angle = math.radians(-90 + i * 360 / n)
        bx_c = int(cx + arm * math.cos(angle))
        by_c = int(cy + arm * math.sin(angle))
        c = _CHART_COLORS[i % len(_CHART_COLORS)]
        draw.line([(cx, cy), (bx_c, by_c)], fill=c, width=2)
        rx2 = max(rx + 4, min(bx_c - bw // 2, rx + rw - bw - 4))
        ry2 = max(ry + 4, min(by_c - bh // 2, ry + rh - bh - 4))
        draw.rectangle([rx2, ry2, rx2 + bw, ry2 + bh], fill=c, outline=_WHITE, width=1)
        pb = draw.textbbox((0, 0), part[:22], font=bfont)
        draw.text((rx2 + (bw - (pb[2] - pb[0])) // 2,
                   ry2 + (bh - (pb[3] - pb[1])) // 2),
                  part[:22], font=bfont, fill=_WHITE)


def _v3_draw_comparison_cards(
    draw: ImageDraw.ImageDraw,
    items: list,
    color: tuple,
    rx: int, ry: int, rw: int, rh: int,
) -> None:
    """Draw 3 side-by-side colored comparison cards."""
    n_cards = min(3, len(items)) if items else 3
    texts = [str(s) for s in items[:n_cards]] if items else ["Type A", "Type B", "Type C"]
    gap = 24
    card_w = (rw - gap * (n_cards + 1)) // n_cards
    card_h = rh - 40
    title_font = _load_font(bold=True, size=36)
    body_font = _load_font(bold=False, size=28)
    card_colors = [
        color,
        tuple(max(0, c - 40) for c in color),
        tuple(min(255, c + 35) for c in color),
    ]
    for i, text in enumerate(texts):
        cx = rx + gap + i * (card_w + gap)
        cy = ry + 20
        c = card_colors[i % len(card_colors)]
        draw.rectangle([cx, cy, cx + card_w, cy + card_h],
                       fill=_CARD_BG, outline=c, width=3)
        # Color header band
        draw.rectangle([cx, cy, cx + card_w, cy + 70], fill=c)
        header = f"Type {chr(65 + i)}"
        words = text.split()
        if len(words) >= 2 and words[0].lower() == "type":
            header = f"{words[0]} {words[1]}"
        hb = draw.textbbox((0, 0), header[:20], font=title_font)
        draw.text((cx + (card_w - (hb[2] - hb[0])) // 2, cy + (70 - (hb[3] - hb[1])) // 2),
                  header[:20], font=title_font, fill=_WHITE)
        # Body text wrapped
        _draw_wrapped_text(draw, text, body_font, cx + 16, cy + 84, card_w - 32, _LIGHT_GRAY, 8)


def _v3_draw_machine_diagram(
    draw: ImageDraw.ImageDraw,
    items: list,
    color: tuple,
    topic: str,
    rx: int, ry: int, rw: int, rh: int,
) -> None:
    """Draw a simplified machine cross-section diagram with labeled parts."""
    parts_font = _load_font(bold=False, size=28)
    label_font = _load_font(bold=True, size=30)
    title_font = _load_font(bold=True, size=34)

    # Draw process flow: Hopper → Barrel → Screw → Nozzle → Mould → Product
    steps = ["Hopper", "Barrel+Screw", "Heater Bands", "Nozzle", "Mould", "Product"]
    if items:
        steps = [str(s)[:20] for s in items[:6]]

    n = len(steps)
    box_h = min(120, rh // 3)
    box_w = max(140, (rw - 60 * n) // n)
    start_x = rx + 30
    mid_y = ry + rh // 2 - box_h // 2

    # Draw main body
    body_y1 = mid_y - 10
    body_y2 = mid_y + box_h + 10
    body_x1 = start_x
    body_x2 = rx + rw - 30
    draw.rectangle([body_x1, body_y1, body_x2, body_y2],
                   fill=_CARD_BG, outline=color, width=3)

    # Draw each component box
    for i, step in enumerate(steps):
        bx = start_x + i * (box_w + 30)
        by = mid_y
        c = _CHART_COLORS[i % len(_CHART_COLORS)]
        draw.rectangle([bx, by, bx + box_w, by + box_h],
                       fill=c, outline=_WHITE, width=2)
        sb = draw.textbbox((0, 0), step, font=parts_font)
        draw.text((bx + (box_w - (sb[2] - sb[0])) // 2,
                   by + (box_h - (sb[3] - sb[1])) // 2),
                  step, font=parts_font, fill=_WHITE)
        # Arrow between boxes
        if i < n - 1:
            ax = bx + box_w + 2
            ay = mid_y + box_h // 2
            draw.line([(ax, ay), (ax + 26, ay)], fill=color, width=4)
            draw.polygon([(ax + 30, ay), (ax + 18, ay - 8), (ax + 18, ay + 8)], fill=color)

    # Material flow label
    flow_text = "Material Flow →"
    fb = draw.textbbox((0, 0), flow_text, font=label_font)
    draw.text((rx + (rw - (fb[2] - fb[0])) // 2, body_y2 + 20),
              flow_text, font=label_font, fill=color)

    # Topic title at top of diagram area — append "Machine" if not already in topic
    if "machine" not in topic.lower():
        tt = topic[:30] + " Machine"
    else:
        tt = topic[:38]
    tb = draw.textbbox((0, 0), tt, font=title_font)
    draw.text((rx + (rw - (tb[2] - tb[0])) // 2, ry + 10),
              tt, font=title_font, fill=color)


def _v3_draw_formula_sheet(
    draw: ImageDraw.ImageDraw,
    items: list,
    color: tuple,
    rx: int, ry: int, rw: int, rh: int,
) -> None:
    """Draw a formula sheet with numbered rows and large text."""
    formulae = [str(s)[:60] for s in items[:6]] if items else ["F = m × a", "E = mc²"]
    hfont = _load_font(bold=True, size=34)
    ffont = _load_font(bold=True, size=32)
    row_h = min(120, (rh - 60) // max(len(formulae), 1))

    draw.rectangle([rx, ry, rx + rw, ry + 52], fill=color)
    header = "Key Formulas & Parameters"
    hb = draw.textbbox((0, 0), header, font=hfont)
    draw.text((rx + (rw - (hb[2] - hb[0])) // 2, ry + 8), header, font=hfont, fill=_WHITE)

    for i, formula in enumerate(formulae):
        y = ry + 60 + i * row_h
        bg = (0x1E, 0x38, 0x50) if i % 2 == 0 else _CARD_BG
        draw.rectangle([rx, y, rx + rw, y + row_h], fill=bg, outline=(0x28, 0x45, 0x62), width=1)
        # Numbered badge
        draw.rectangle([rx, y, rx + 46, y + row_h], fill=color)
        nb = draw.textbbox((0, 0), str(i + 1), font=ffont)
        draw.text((rx + (46 - (nb[2] - nb[0])) // 2,
                   y + (row_h - (nb[3] - nb[1])) // 2),
                  str(i + 1), font=ffont, fill=_WHITE)
        _draw_wrapped_text(draw, formula, ffont, rx + 56, y + 12, rw - 70, _WHITE, 4)


def _v3_draw_grid_cards(
    draw: ImageDraw.ImageDraw,
    items: list,
    color: tuple,
    rx: int, ry: int, rw: int, rh: int,
) -> None:
    """Draw a 2×2 grid of colored application cards."""
    texts = [str(s) for s in items[:4]] if items else ["App A", "App B", "App C", "App D"]
    while len(texts) < 4:
        texts.append(f"Item {len(texts) + 1}")
    gap = 20
    card_w = (rw - gap * 3) // 2
    card_h = (rh - gap * 3) // 2
    title_font = _load_font(bold=True, size=32)
    body_font = _load_font(bold=False, size=28)
    positions = [
        (rx + gap, ry + gap),
        (rx + gap * 2 + card_w, ry + gap),
        (rx + gap, ry + gap * 2 + card_h),
        (rx + gap * 2 + card_w, ry + gap * 2 + card_h),
    ]
    for i, (text, (cx, cy)) in enumerate(zip(texts, positions)):
        c = _CHART_COLORS[i % len(_CHART_COLORS)]
        draw.rectangle([cx, cy, cx + card_w, cy + card_h],
                       fill=_CARD_BG, outline=c, width=3)
        draw.rectangle([cx, cy, cx + card_w, cy + 56], fill=c)
        # Card number badge
        badge = str(i + 1)
        bfont = _load_font(bold=True, size=32)
        bb = draw.textbbox((0, 0), badge, font=bfont)
        draw.text((cx + 16, cy + (56 - (bb[3] - bb[1])) // 2), badge, font=bfont, fill=_WHITE)
        # Card title (first ~20 chars of text)
        tb = draw.textbbox((0, 0), text[:24], font=title_font)
        draw.text((cx + 52, cy + (56 - (tb[3] - tb[1])) // 2),
                  text[:24], font=title_font, fill=_WHITE)
        # Card body (remaining text)
        if len(text) > 24:
            _draw_wrapped_text(draw, text[24:], body_font, cx + 16, cy + 68, card_w - 32, _LIGHT_GRAY, 6)


def _v3_draw_comparison_table(
    draw: ImageDraw.ImageDraw,
    items: list,
    color: tuple,
    rx: int, ry: int, rw: int, rh: int,
) -> None:
    """Draw a full-width comparison table with alternating row colors."""
    rows = [str(s) for s in items[:8]] if items else ["Property: Value"]
    hfont = _load_font(bold=True, size=32)
    rfont = _load_font(bold=False, size=28)
    col_w = rw // 2
    row_h = min(80, (rh - 60) // max(len(rows) + 1, 1))

    # Header row
    draw.rectangle([rx, ry, rx + col_w, ry + 56], fill=color)
    draw.rectangle([rx + col_w, ry, rx + rw, ry + 56],
                   fill=tuple(max(0, c - 40) for c in color))
    draw.text((rx + 16, ry + 8), "Property / Standard", font=hfont, fill=_WHITE)
    draw.text((rx + col_w + 16, ry + 8), "Value / Details", font=hfont, fill=_WHITE)

    alt_bg = (0x1E, 0x38, 0x50)
    for i, row in enumerate(rows):
        y = ry + 60 + i * row_h
        if y + row_h > ry + rh - 4:
            break
        bg = alt_bg if i % 2 == 0 else _CARD_BG
        draw.rectangle([rx, y, rx + rw, y + row_h], fill=bg, outline=(0x28, 0x48, 0x68), width=1)
        if ":" in row:
            parts = row.split(":", 1)
            label, val = parts[0].strip(), parts[1].strip()
        else:
            label, val = row, "—"
        _draw_wrapped_text(draw, label[:35], rfont, rx + 16, y + 8, col_w - 24, color, 3)
        _draw_wrapped_text(draw, val[:40], rfont, rx + col_w + 16, y + 8, col_w - 24, _WHITE, 3)


def _v3_draw_cycle_diagram(
    draw: ImageDraw.ImageDraw,
    items: list,
    color: tuple,
    rx: int, ry: int, rw: int, rh: int,
) -> None:
    """Draw a circular cycle diagram with arrows."""
    steps = [str(s)[:22] for s in items[:6]] if items else ["Step 1", "Step 2", "Step 3", "Step 4"]
    n = len(steps)
    font = _load_font(bold=True, size=28)

    cx = rx + rw // 2
    cy = ry + rh // 2
    orbit_r = min(rw, rh) // 3
    node_r = min(rw, rh) // 10

    # Draw orbit circle (dashed appearance via many short lines)
    for deg in range(0, 360, 6):
        a1 = math.radians(deg)
        a2 = math.radians(deg + 3)
        px1 = int(cx + orbit_r * math.cos(a1))
        py1 = int(cy + orbit_r * math.sin(a1))
        px2 = int(cx + orbit_r * math.cos(a2))
        py2 = int(cy + orbit_r * math.sin(a2))
        draw.line([(px1, py1), (px2, py2)], fill=color, width=2)

    for i, step in enumerate(steps):
        angle = math.radians(-90 + i * 360 / n)
        nx = int(cx + orbit_r * math.cos(angle))
        ny = int(cy + orbit_r * math.sin(angle))
        c = _CHART_COLORS[i % len(_CHART_COLORS)]

        # Draw arrow along orbit to next node
        next_angle = math.radians(-90 + (i + 1) * 360 / n)
        mid_angle = (angle + next_angle) / 2
        ax = int(cx + orbit_r * math.cos(mid_angle))
        ay = int(cy + orbit_r * math.sin(mid_angle))
        d_angle = next_angle - angle
        perp = mid_angle + math.pi / 2
        arrow_tip_x = int(ax + 14 * math.cos(perp) * (1 if d_angle > 0 else -1))
        arrow_tip_y = int(ay + 14 * math.sin(perp) * (1 if d_angle > 0 else -1))
        draw.polygon([
            (arrow_tip_x, arrow_tip_y),
            (int(ax - 8 * math.cos(perp)), int(ay - 8 * math.sin(perp))),
            (int(ax + 8 * math.cos(perp)), int(ay + 8 * math.sin(perp))),
        ], fill=color)

        draw.ellipse([nx - node_r, ny - node_r, nx + node_r, ny + node_r],
                     fill=c, outline=_WHITE, width=2)
        sb = draw.textbbox((0, 0), step[:16], font=font)
        # Place text outside the node
        text_r = orbit_r + node_r + 20
        tx = int(cx + text_r * math.cos(angle)) - (sb[2] - sb[0]) // 2
        ty = int(cy + text_r * math.sin(angle)) - (sb[3] - sb[1]) // 2
        tx = max(rx + 4, min(tx, rx + rw - (sb[2] - sb[0]) - 4))
        ty = max(ry + 4, min(ty, ry + rh - (sb[3] - sb[1]) - 4))
        draw.text((tx, ty), step[:16], font=font, fill=_LIGHT_GRAY)


def _v3_draw_timeline(
    draw: ImageDraw.ImageDraw,
    items: list,
    color: tuple,
    rx: int, ry: int, rw: int, rh: int,
) -> None:
    """Draw a full-width horizontal timeline."""
    events = [str(s)[:28] for s in items[:7]] if items else ["Event 1", "Event 2", "Event 3"]
    n = len(events)
    font = _load_font(bold=False, size=28)
    tfont = _load_font(bold=True, size=32)

    line_y = ry + rh // 2
    pad_x = rx + 50
    avail = rw - 100

    # Main axis
    draw.line([(pad_x, line_y), (rx + rw - pad_x, line_y)], fill=color, width=5)
    # Arrow at end
    draw.polygon([
        (rx + rw - pad_x + 16, line_y),
        (rx + rw - pad_x, line_y - 10),
        (rx + rw - pad_x, line_y + 10),
    ], fill=color)

    for i, event in enumerate(events):
        ex = pad_x + int(avail * i / max(n - 1, 1))
        c = _CHART_COLORS[i % len(_CHART_COLORS)]

        # Dot marker
        draw.ellipse([ex - 14, line_y - 14, ex + 14, line_y + 14],
                     fill=c, outline=_WHITE, width=2)
        num = str(i + 1)
        nb = draw.textbbox((0, 0), num, font=tfont)
        draw.text((ex - (nb[2] - nb[0]) // 2, line_y - (nb[3] - nb[1]) // 2),
                  num, font=tfont, fill=_WHITE)

        # Label alternating above/below
        label = event[:24]
        bb = draw.textbbox((0, 0), label, font=font)
        lw = bb[2] - bb[0]
        lx = max(rx + 2, min(ex - lw // 2, rx + rw - lw - 4))
        if i % 2 == 0:
            draw.line([(ex, line_y - 16), (ex, line_y - 50)], fill=c, width=2)
            draw.text((lx, line_y - 56 - (bb[3] - bb[1])),
                      label, font=font, fill=_LIGHT_GRAY)
        else:
            draw.line([(ex, line_y + 16), (ex, line_y + 50)], fill=c, width=2)
            draw.text((lx, line_y + 56), label, font=font, fill=_LIGHT_GRAY)


def _v3_draw_mind_map(
    draw: ImageDraw.ImageDraw,
    items: list,
    color: tuple,
    title: str,
    rx: int, ry: int, rw: int, rh: int,
) -> None:
    """Draw a mind map with central topic node and radiating branches."""
    branches = [str(s)[:28] for s in items[:8]] if items else ["Topic A", "Topic B", "Topic C"]
    n = len(branches)
    cfont = _load_font(bold=True, size=36)
    bfont = _load_font(bold=False, size=28)

    cx = rx + rw // 2
    cy = ry + rh // 2
    cr = min(rw, rh) // 7
    arm_len = min(rw, rh) // 3
    bw = max(160, rw // 6)
    bh = 52

    draw.ellipse([cx - cr, cy - cr, cx + cr, cy + cr],
                 fill=color, outline=_WHITE, width=3)
    short = (title[:14] + "..") if len(title) > 16 else title
    tb = draw.textbbox((0, 0), short, font=cfont)
    draw.text((cx - (tb[2] - tb[0]) // 2, cy - (tb[3] - tb[1]) // 2),
              short, font=cfont, fill=_WHITE)

    for i, branch in enumerate(branches):
        angle = math.radians(-90 + i * 360 / n)
        bx_c = int(cx + arm_len * math.cos(angle))
        by_c = int(cy + arm_len * math.sin(angle))
        c = _CHART_COLORS[i % len(_CHART_COLORS)]
        draw.line([(cx, cy), (bx_c, by_c)], fill=c, width=3)

        lx = max(rx + 4, min(bx_c - bw // 2, rx + rw - bw - 4))
        ly = max(ry + 4, min(by_c - bh // 2, ry + rh - bh - 4))
        draw.rectangle([lx, ly, lx + bw, ly + bh], fill=c, outline=_WHITE, width=1)
        pb = draw.textbbox((0, 0), branch[:22], font=bfont)
        draw.text((lx + (bw - (pb[2] - pb[0])) // 2,
                   ly + (bh - (pb[3] - pb[1])) // 2),
                  branch[:22], font=bfont, fill=_WHITE)


# ---------------------------------------------------------------------------
# V3.0 Per-slide layout functions
# Each layout fills the content_rect = (rx, ry, rw, rh) below the title bar.
# ---------------------------------------------------------------------------


def _v3_render_common_header(
    draw: ImageDraw.ImageDraw,
    slide_data: dict,
    slide_index: int,
    total_slides: int,
    department: str,
    accent: tuple,
) -> int:
    """Draw common V3 header (top bar, title, underline) and return content start y."""
    # Top accent bar
    draw.rectangle([0, 0, 1920, 14], fill=accent)

    # Slide title
    title_font = _load_font(bold=True, size=56)
    title_text = slide_data.get("title", f"Slide {slide_index}")
    title_x, title_y = 32, 24
    draw.text((title_x, title_y), title_text, font=title_font, fill=accent)

    # Title underline
    tb = draw.textbbox((title_x, title_y), title_text, font=title_font)
    uline_y = tb[3] + 6
    draw.rectangle([title_x, uline_y, 1888, uline_y + 4], fill=accent)

    # Slide number badge (top right)
    num_font = _load_font(bold=True, size=28)
    num_text = f"{slide_index}/{total_slides}"
    nb = draw.textbbox((0, 0), num_text, font=num_font)
    draw.text((1888 - (nb[2] - nb[0]), 30), num_text, font=num_font, fill=accent)

    return uline_y + 14  # content area starts here


def _v3_render_common_footer(draw: ImageDraw.ImageDraw, department: str, accent: tuple) -> int:
    """Draw common V3 bottom bar and return its y position."""
    bot_y = 1000
    draw.rectangle([0, bot_y, 1920, 1080], fill=_BOTTOM_BG)
    bot_font = _load_font(bold=False, size=24)
    bot_text = f"CIPETHUB  \u2022  {department} Engineering  \u2022  Subscribe for more"
    bb = draw.textbbox((0, 0), bot_text, font=bot_font)
    bx = (1920 - (bb[2] - bb[0])) // 2
    by = bot_y + (80 - (bb[3] - bb[1])) // 2
    draw.text((bx, by), bot_text, font=bot_font, fill=_LIGHT_GRAY)
    return bot_y


def _v3_draw_bullets(
    draw: ImageDraw.ImageDraw,
    bullets: list,
    accent: tuple,
    rx: int, ry: int, rw: int, rh: int,
) -> None:
    """Draw large bullet points (36px) in the specified rectangle."""
    font = _load_font(bold=False, size=36)
    y = ry + 16
    for i, bullet in enumerate(bullets[:4]):
        if i > 0:
            sep_y = y - 8
            draw.line([(rx + 12, sep_y), (rx + rw - 12, sep_y)],
                      fill=(0x25, 0x40, 0x58), width=1)
        prefix = "\u25cf  " if i == 0 else "\u25b8  "
        y = _draw_wrapped_text(
            draw, f"{prefix}{bullet}", font,
            rx + 16, y, rw - 32, _WHITE, line_spacing=10,
        )
        y += 20


# ---------------------------------------------------------------------------
# V3.0 Slide Renderer — 12 distinct layouts with per-slide accent colors
# ---------------------------------------------------------------------------


def render_slide_image_v3(
    slide_data: dict,
    slide_index: int,
    total_slides: int,
    department: str,
) -> Image.Image:
    """Render one 1920×1080 PNG slide image V3 with per-slide layout and color.

    Uses SLIDE_COLORS_PIL to assign a unique accent color to each slide.
    12 different layouts are cycled through based on slide_index.
    Minimum font size: 28px. Title: 56px. Bullets: 36px.
    70%+ of content area is visual.
    """
    idx0 = (slide_index - 1) % 12  # 0-based layout index
    accent = SLIDE_COLORS_PIL[idx0]

    img = Image.new("RGB", (1920, 1080), _BG_DARK)
    draw = ImageDraw.Draw(img)

    bullets = slide_data.get("bullets", [])
    title = slide_data.get("title", f"Slide {slide_index}")
    visual_desc = slide_data.get("visual_description",
                                 slide_data.get("visual_hint", ""))

    content_y = _v3_render_common_header(draw, slide_data, slide_index,
                                         total_slides, department, accent)
    bot_y = _v3_render_common_footer(draw, department, accent)

    # Content rectangle
    rx, ry = 32, content_y
    rw = 1920 - 64
    rh = bot_y - ry - 8

    # Reuse bullets drawn to left panel for split layouts
    left_w = rw * 35 // 100   # ~35% text
    right_w = rw - left_w - 20  # ~65% visual

    # --- Layout 1: INTRO — full-screen title card ---
    if idx0 == 0:
        # Draw large topic name
        big_font = _load_font(bold=True, size=72)
        sub_font = _load_font(bold=False, size=36)

        # Gradient-like background band
        for dy in range(rh):
            frac = dy / rh
            r = int(_BG_DARK[0] + (accent[0] - _BG_DARK[0]) * frac * 0.3)
            g = int(_BG_DARK[1] + (accent[1] - _BG_DARK[1]) * frac * 0.3)
            b = int(_BG_DARK[2] + (accent[2] - _BG_DARK[2]) * frac * 0.3)
            draw.line([(0, ry + dy), (1920, ry + dy)], fill=(r, g, b))

        tb = draw.textbbox((0, 0), title, font=big_font)
        tx = (1920 - (tb[2] - tb[0])) // 2
        ty = ry + (rh - (tb[3] - tb[1])) // 2 - 40
        draw.text((tx, ty), title, font=big_font, fill=accent)

        sub = "Let's Learn!"
        sb = draw.textbbox((0, 0), sub, font=sub_font)
        draw.text(((1920 - (sb[2] - sb[0])) // 2, ty + (tb[3] - tb[1]) + 24),
                  sub, font=sub_font, fill=_LIGHT_GRAY)

    # --- Layout 2: CONCEPTS — left bullets | right labeled diagram ---
    elif idx0 == 1:
        draw.rectangle([rx, ry, rx + left_w, ry + rh], fill=_CARD_BG)
        _v3_draw_bullets(draw, bullets, accent, rx, ry, left_w, rh)
        vis_x = rx + left_w + 20
        draw.rectangle([vis_x, ry, vis_x + right_w, ry + rh], fill=_CARD_BG)
        _v3_draw_labeled_diagram(draw, bullets, accent, title,
                                 vis_x, ry, right_w, rh)

    # --- Layout 3: TYPES — full-width 3 comparison cards ---
    elif idx0 == 2:
        _v3_draw_comparison_cards(draw, bullets, accent, rx, ry, rw, rh)

    # --- Layout 4: PROPERTIES — left mini text | right large bar chart ---
    elif idx0 == 3:
        draw.rectangle([rx, ry, rx + left_w, ry + rh], fill=_CARD_BG)
        _v3_draw_bullets(draw, bullets, accent, rx, ry, left_w, rh)
        vis_x = rx + left_w + 20
        draw.rectangle([vis_x, ry, vis_x + right_w, ry + rh], fill=_CARD_BG)
        _v3_draw_bar_chart(draw, bullets, accent, vis_x, ry, right_w, rh)

    # --- Layout 5: PROCESS — full-width flowchart ---
    elif idx0 == 4:
        draw.rectangle([rx, ry, rx + rw, ry + rh], fill=_CARD_BG)
        _v3_draw_flowchart(draw, bullets, accent, rx, ry, rw, rh)

    # --- Layout 6: EQUIPMENT — full-width machine diagram ---
    elif idx0 == 5:
        draw.rectangle([rx, ry, rx + rw, ry + rh], fill=_CARD_BG)
        _v3_draw_machine_diagram(draw, bullets, accent, title, rx, ry, rw, rh)

    # --- Layout 7: CALCULATIONS — left formula | right worked example ---
    elif idx0 == 6:
        half_w = rw // 2 - 10
        draw.rectangle([rx, ry, rx + half_w, ry + rh], fill=_CARD_BG)
        _v3_draw_formula_sheet(draw, bullets[:3], accent, rx, ry, half_w, rh)
        vis_x = rx + half_w + 20
        draw.rectangle([vis_x, ry, vis_x + half_w, ry + rh], fill=_CARD_BG)
        _v3_draw_bullets(draw, bullets[3:] or bullets, accent, vis_x, ry, half_w, rh)

    # --- Layout 8: APPLICATIONS — 2×2 grid cards ---
    elif idx0 == 7:
        _v3_draw_grid_cards(draw, bullets, accent, rx, ry, rw, rh)

    # --- Layout 9: STANDARDS — full-width comparison table ---
    elif idx0 == 8:
        draw.rectangle([rx, ry, rx + rw, ry + rh], fill=_CARD_BG)
        _v3_draw_comparison_table(draw, bullets, accent, rx, ry, rw, rh)

    # --- Layout 10: ENVIRONMENT — left key points | right cycle diagram ---
    elif idx0 == 9:
        draw.rectangle([rx, ry, rx + left_w, ry + rh], fill=_CARD_BG)
        _v3_draw_bullets(draw, bullets, accent, rx, ry, left_w, rh)
        vis_x = rx + left_w + 20
        draw.rectangle([vis_x, ry, vis_x + right_w, ry + rh], fill=_CARD_BG)
        _v3_draw_cycle_diagram(draw, bullets, accent, vis_x, ry, right_w, rh)

    # --- Layout 11: TRENDS — full-width timeline ---
    elif idx0 == 10:
        draw.rectangle([rx, ry, rx + rw, ry + rh], fill=_CARD_BG)
        _v3_draw_timeline(draw, bullets, accent, rx, ry, rw, rh)

    # --- Layout 12: SUMMARY — full-width mind map ---
    elif idx0 == 11:
        draw.rectangle([rx, ry, rx + rw, ry + rh], fill=_CARD_BG)
        _v3_draw_mind_map(draw, bullets, accent, title, rx, ry, rw, rh)

    return img


# ---------------------------------------------------------------------------
# Slide Image Rendering (Pillow) — V2 (kept for reference)
# ---------------------------------------------------------------------------

# Scale factor: 1920 px / 16 inches = 120 px per inch


def render_slide_image(
    slide_data: dict,
    department: str,
    slide_index: int,
    total_slides: int,
) -> Image.Image:
    """Render one 1920×1080 PNG slide image using Pillow.

    Left panel  (≈60 %): title + bullet points.
    Right panel (≈40 %): actual generated diagram/chart from render_visual().

    Args:
        slide_data:   Slide dict (title, bullets, visual_hint).
        department:   Used to look up accent colour.
        slide_index:  1-based index shown in slide-number label.
        total_slides: Total count shown in slide-number label.

    Returns:
        RGB Pillow Image, 1920 × 1080.
    """
    dept_color = DEPT_COLORS_PIL.get(department, DEPT_COLORS_PIL["All"])
    img  = Image.new("RGB", (1920, 1080), _BG_DARK)
    draw = ImageDraw.Draw(img)

    # --- Top accent bar ---
    draw.rectangle([0, 0, 1920, int(0.12 * _PX)], fill=dept_color)

    # --- Subtle gradient strip below accent bar ---
    for dy in range(6):
        alpha = int(40 * (1 - dy / 6))
        r = min(255, dept_color[0] + alpha)
        g = min(255, dept_color[1] + alpha)
        b = min(255, dept_color[2] + alpha)
        draw.line([(0, int(0.12 * _PX) + dy), (1920, int(0.12 * _PX) + dy)],
                  fill=(r, g, b))

    # --- Department badge (top left) ---
    bx, by, bw, bh = int(0.2 * _PX), int(0.18 * _PX), int(2.0 * _PX), int(0.38 * _PX)
    draw.rectangle([bx, by, bx + bw, by + bh], fill=dept_color)
    badge_font = _load_font(bold=True, size=14)
    badge_text = department.upper()
    bb = draw.textbbox((0, 0), badge_text, font=badge_font)
    draw.text((bx + (bw - (bb[2] - bb[0])) // 2,
               by + (bh - (bb[3] - bb[1])) // 2),
              badge_text, font=badge_font, fill=_WHITE)

    # --- CIPETHUB branding (top right) ---
    brand_font = _load_font(bold=True, size=26)
    sub_font   = _load_font(bold=False, size=13)
    brand_text = "CIPETHUB"
    sub_text   = "CIPET Study Material"
    brand_bb = draw.textbbox((0, 0), brand_text, font=brand_font)
    sub_bb   = draw.textbbox((0, 0), sub_text,   font=sub_font)
    brand_x  = 1920 - int(0.2 * _PX) - (brand_bb[2] - brand_bb[0])
    sub_x    = 1920 - int(0.2 * _PX) - (sub_bb[2] - sub_bb[0])
    draw.text((brand_x, int(0.18 * _PX)), brand_text, font=brand_font, fill=dept_color)
    draw.text((sub_x, int(0.18 * _PX) + (brand_bb[3] - brand_bb[1]) + 4),
              sub_text, font=sub_font, fill=_LIGHT_GRAY)

    # --- Slide title ---
    title_font = _load_font(bold=True, size=46)
    title_text = slide_data.get("title", f"Slide {slide_index}")
    draw.text((int(0.3 * _PX), int(0.65 * _PX)), title_text,
              font=title_font, fill=dept_color)

    # --- Title underline bar ---
    ubar_y = int(1.45 * _PX)
    draw.rectangle([int(0.3 * _PX), ubar_y, int(10.8 * _PX), ubar_y + int(0.05 * _PX)],
                   fill=dept_color)

    # --- Content card ---
    card_x1, card_y1 = int(0.3 * _PX), int(1.55 * _PX)
    card_x2, card_y2 = int(10.8 * _PX), int(7.65 * _PX)
    draw.rectangle([card_x1, card_y1, card_x2, card_y2], fill=_CARD_BG)

    # --- Bullet points with thin separator lines ---
    bullet_font = _load_font(bold=False, size=28)
    bul_x  = int(0.55 * _PX)
    bul_y  = int(1.75 * _PX)
    max_bw = card_x2 - bul_x - int(0.2 * _PX)
    for bi, bullet in enumerate(slide_data.get("bullets", [])):
        if bi > 0:
            # thin separator
            sep_y = bul_y - 6
            draw.line([(bul_x, sep_y), (card_x2 - int(0.1 * _PX), sep_y)],
                      fill=(0x25, 0x40, 0x58), width=1)
        prefix = "📌  " if bi == 0 else "▸   "
        bul_y = _draw_wrapped_text(
            draw, f"{prefix}{bullet}", bullet_font,
            bul_x, bul_y, max_bw, _WHITE, line_spacing=8,
        )
        bul_y += 12

    # --- Right panel — actual rendered visual ---
    hint_x1 = int(11.0 * _PX)
    hint_y1 = int(1.55 * _PX)
    hint_x2 = int(15.7 * _PX)
    hint_y2 = int(7.65 * _PX)
    draw.rectangle([hint_x1, hint_y1, hint_x2, hint_y2], fill=_CARD_BG)
    draw.rectangle([hint_x1, hint_y1, hint_x2, hint_y2], outline=dept_color, width=2)

    # Visual panel header
    hint_label_font = _load_font(bold=True, size=15)
    label_text = "VISUAL DIAGRAM"
    lb = draw.textbbox((0, 0), label_text, font=hint_label_font)
    lx = hint_x1 + ((hint_x2 - hint_x1) - (lb[2] - lb[0])) // 2
    draw.text((lx, hint_y1 + int(0.1 * _PX)), label_text,
              font=hint_label_font, fill=dept_color)

    # Thin line under label
    sep_y2 = hint_y1 + int(0.44 * _PX)
    draw.line([(hint_x1 + 8, sep_y2), (hint_x2 - 8, sep_y2)], fill=dept_color, width=1)

    # Render and paste the actual visual
    try:
        visual_img = render_visual(
            slide_data.get("visual_hint", ""),
            slide_data,
            department,
            slide_index,
        )
        # Resize to fit panel (keeping aspect ratio)
        panel_w = hint_x2 - hint_x1 - 16
        panel_h = hint_y2 - hint_y1 - int(0.55 * _PX) - 8
        visual_img = visual_img.resize(
            (panel_w, panel_h),
            Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS,
        )
        img.paste(visual_img, (hint_x1 + 8, hint_y1 + int(0.55 * _PX)))
    except Exception as vis_err:
        print(f"[generator] Visual render error (slide {slide_index}): {vis_err}")
        # Fallback: draw hint text
        hint_text_font = _load_font(bold=False, size=17)
        _draw_wrapped_text(
            draw,
            slide_data.get("visual_hint", "Diagram here"),
            hint_text_font,
            hint_x1 + int(0.1 * _PX),
            hint_y1 + int(0.55 * _PX),
            (hint_x2 - hint_x1) - int(0.2 * _PX),
            _LIGHT_GRAY,
            line_spacing=6,
        )

    # --- Bottom bar ---
    bot_y = int(8.7 * _PX)
    draw.rectangle([0, bot_y, 1920, 1080], fill=_BOTTOM_BG)
    bot_font = _load_font(bold=False, size=14)
    bot_text = f"CIPETHUB  \u2022  {department} Engineering  \u2022  Subscribe & Like!"
    bot_bb   = draw.textbbox((0, 0), bot_text, font=bot_font)
    bot_tx   = (1920 - (bot_bb[2] - bot_bb[0])) // 2
    bot_ty   = bot_y + (1080 - bot_y - (bot_bb[3] - bot_bb[1])) // 2
    draw.text((bot_tx, bot_ty), bot_text, font=bot_font, fill=_LIGHT_GRAY)

    # --- Slide number (bottom right) ---
    num_font = _load_font(bold=True, size=20)
    num_text = f"{slide_index}/{total_slides}"
    nb = draw.textbbox((0, 0), num_text, font=num_font)
    draw.text((1920 - int(0.3 * _PX) - (nb[2] - nb[0]), bot_ty),
              num_text, font=num_font, fill=dept_color)

    return img


# ---------------------------------------------------------------------------
# Intro / Outro slides (Upgrade 6)
# ---------------------------------------------------------------------------


def render_intro_image(topic: str, department: str) -> Image.Image:
    """Render a 1920×1080 intro slide — large topic name, minimal branding."""
    accent = SLIDE_COLORS_PIL[0]  # Orange for intro
    img  = Image.new("RGB", (1920, 1080), _BG_DARK)
    draw = ImageDraw.Draw(img)

    # Full-width accent bars
    draw.rectangle([0, 0, 1920, 14], fill=accent)
    draw.rectangle([0, 1066, 1920, 1080], fill=accent)
    draw.rectangle([0, 0, 10, 1080], fill=accent)

    # Gradient background
    for dy in range(1080):
        frac = dy / 1080
        r = int(_BG_DARK[0] + (accent[0] - _BG_DARK[0]) * frac * 0.25)
        g = int(_BG_DARK[1] + (accent[1] - _BG_DARK[1]) * frac * 0.25)
        b = int(_BG_DARK[2] + (accent[2] - _BG_DARK[2]) * frac * 0.25)
        draw.line([(10, dy), (1920, dy)], fill=(r, g, b))

    big_font   = _load_font(bold=True, size=96)
    dept_font  = _load_font(bold=True, size=40)
    sub_font   = _load_font(bold=False, size=32)
    brand_font = _load_font(bold=False, size=26)

    # Topic name — HUGE
    topic_text = topic[:60]
    tb = draw.textbbox((0, 0), topic_text, font=big_font)
    tx = (1920 - (tb[2] - tb[0])) // 2
    draw.text((tx, 280), topic_text, font=big_font, fill=accent)

    # Horizontal divider
    draw.rectangle([200, 420, 1720, 426], fill=accent)

    # Department badge
    dept_text = f"{department} Engineering"
    db = draw.textbbox((0, 0), dept_text, font=dept_font)
    dx = (1920 - (db[2] - db[0])) // 2
    draw.rectangle([dx - 24, 446, dx + (db[2] - db[0]) + 24, 446 + (db[3] - db[1]) + 20],
                   fill=accent)
    draw.text((dx, 456), dept_text, font=dept_font, fill=_WHITE)

    # "Let's learn!" tagline
    sub_text = "Let's Learn!"
    sb = draw.textbbox((0, 0), sub_text, font=sub_font)
    draw.text(((1920 - (sb[2] - sb[0])) // 2, 560), sub_text, font=sub_font, fill=_LIGHT_GRAY)

    # CIPETHUB branding (small, once)
    brand_text = "CIPETHUB"
    bb = draw.textbbox((0, 0), brand_text, font=brand_font)
    draw.text((1920 - (bb[2] - bb[0]) - 20, 1032), brand_text, font=brand_font, fill=accent)

    return img


def render_outro_image(department: str) -> Image.Image:
    """Render a 1920×1080 outro slide with subscribe CTA."""
    accent = SLIDE_COLORS_PIL[11]  # Blue Grey for outro
    img  = Image.new("RGB", (1920, 1080), _BG_DARK)
    draw = ImageDraw.Draw(img)

    draw.rectangle([0, 0, 1920, 14], fill=accent)
    draw.rectangle([0, 1066, 1920, 1080], fill=accent)
    draw.rectangle([0, 0, 10, 1080], fill=accent)
    draw.rectangle([1910, 0, 1920, 1080], fill=accent)

    ty_font  = _load_font(bold=True, size=96)
    cta_font = _load_font(bold=True, size=44)
    sub_font = _load_font(bold=False, size=32)

    ty_text = "Thank You!"
    tb = draw.textbbox((0, 0), ty_text, font=ty_font)
    draw.text(((1920 - (tb[2] - tb[0])) // 2, 160),
              ty_text, font=ty_font, fill=accent)

    draw.rectangle([300, 320, 1620, 326], fill=accent)

    # Subscribe / Like / Share buttons
    actions = [("👍", "Like"), ("🔔", "Subscribe"), ("📤", "Share")]
    for i, (icon, act) in enumerate(actions):
        bx = 280 + i * 450
        bw, bh = 380, 100
        draw.rectangle([bx, 360, bx + bw, 360 + bh], fill=accent, outline=_WHITE, width=2)
        label = f"{icon} {act}"
        lb = draw.textbbox((0, 0), label, font=cta_font)
        draw.text((bx + (bw - (lb[2] - lb[0])) // 2, 360 + (bh - (lb[3] - lb[1])) // 2),
                  label, font=cta_font, fill=_WHITE)

    brand_text = "CIPETHUB — Subscribe for more lectures"
    blb = draw.textbbox((0, 0), brand_text, font=sub_font)
    draw.text(((1920 - (blb[2] - blb[0])) // 2, 520),
              brand_text, font=sub_font, fill=accent)

    next_text = "New lectures every week!"
    nb = draw.textbbox((0, 0), next_text, font=sub_font)
    draw.text(((1920 - (nb[2] - nb[0])) // 2, 590),
              next_text, font=sub_font, fill=_LIGHT_GRAY)

    return img


def render_slide_images(script_slides: list, department: str, img_dir: str) -> list:
    """Render each slide to a 1920×1080 PNG via Pillow V3 and save to *img_dir*."""
    os.makedirs(img_dir, exist_ok=True)
    total = len(script_slides)
    paths = []
    for i, slide_data in enumerate(script_slides, start=1):
        out_path = os.path.join(img_dir, f"slide_{i:02d}.png")
        img = render_slide_image_v3(slide_data, i, total, department)
        img.save(out_path, "PNG")
        paths.append(out_path)
        print(f"[generator] Rendered slide image (V3): slide_{i:02d}.png")
    print(f"[generator] Rendered {len(paths)} slide images.")
    return paths


# ---------------------------------------------------------------------------
# Text-to-Speech
# ---------------------------------------------------------------------------


def generate_voices(slides: list, audio_dir: str) -> list:
    """Generate MP3 narration files for each slide using gTTS."""
    os.makedirs(audio_dir, exist_ok=True)
    paths = []
    for i, slide_data in enumerate(slides, start=1):
        narration = slide_data.get("narration", f"Slide {i}.")
        out_path = os.path.join(audio_dir, f"audio_{i:02d}.mp3")
        tts = gTTS(text=narration, lang="en", tld="co.in")
        tts.save(out_path)
        paths.append(out_path)
        print(f"[generator] Audio generated: audio_{i:02d}.mp3")
    return paths


# ---------------------------------------------------------------------------
# Ambient Background Music Generator (Upgrade 5)
# ---------------------------------------------------------------------------


def generate_ambient_audio(duration_sec: float, output_path: str) -> None:
    """Generate a soft ambient background tone and save as a WAV file.

    Uses a low-register minor chord (A2-E3-A3) synthesised entirely in Python
    (no external assets or network access required).  The result is then
    extended to *duration_sec* seconds via FFmpeg's stream_loop.
    """
    loop_dur   = 8.0          # seconds for the base loop
    sample_rate = 22050
    num_samples = int(sample_rate * loop_dur)
    amplitude   = 450         # very quiet (max 32 767)
    frequencies = [110.0, 146.8, 164.8, 220.0]  # A2, D3, E3, A3

    samples = array.array("h")
    for n in range(num_samples):
        t = n / sample_rate
        # Smooth fade-in / fade-out at loop boundaries for seamless looping
        env = 1.0
        if t < 1.5:
            env = t / 1.5
        elif t > loop_dur - 1.5:
            env = (loop_dur - t) / 1.5
        value = sum(math.sin(2 * math.pi * f * t) for f in frequencies)
        sample = int(amplitude * env * value / len(frequencies))
        samples.append(max(-32768, min(32767, sample)))

    loop_wav = output_path + ".loop.wav"
    with wave_module.open(loop_wav, "w") as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)
        wf.setframerate(sample_rate)
        wf.writeframes(samples.tobytes())

    # Extend the loop to the required duration using FFmpeg
    cmd = [
        "ffmpeg", "-y",
        "-stream_loop", "-1",
        "-i", loop_wav,
        "-t", str(duration_sec + 2),  # +2 s buffer ensures audio covers video incl. fade-out
        "-c:a", "pcm_s16le",
        output_path,
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if result.returncode != 0:
            print(f"[generator] Ambient audio extend warning: {result.stderr[-200:]}")
            # Fallback: just copy the short loop as the output
            shutil.copy2(loop_wav, output_path)
    except Exception as ffmpeg_err:
        print(f"[generator] FFmpeg not available for ambient loop ({ffmpeg_err}); "
              "using short loop file.")
        shutil.copy2(loop_wav, output_path)

    try:
        os.remove(loop_wav)
    except OSError:
        pass


# ---------------------------------------------------------------------------
# Duration Helper
# ---------------------------------------------------------------------------


def get_dur(path: str) -> float:
    """Return media file duration in seconds via ffprobe."""
    try:
        result = subprocess.run(
            [
                "ffprobe", "-v", "error",
                "-show_entries", "format=duration",
                "-of", "default=noprint_wrappers=1:nokey=1",
                path,
            ],
            capture_output=True, text=True, timeout=30,
        )
        return float(result.stdout.strip())
    except Exception:
        return 10.0


# ---------------------------------------------------------------------------
# Video Assembly (Upgrades 2, 3, 4, 7)
# ---------------------------------------------------------------------------


def _ffmpeg_escape(text: str) -> str:
    """Escape text so it is safe to embed in an FFmpeg drawtext expression."""
    return (
        text
        .replace("\\", "\\\\")
        .replace("'",  "\\'")
        .replace(":",  "\\:")
        .replace("%",  "\\%")
    )


def build_video(
    image_paths: list,
    audio_paths: list,
    output: str,
    dept_color: tuple,
) -> None:
    """Combine slide images and audio into a final MP4 video.

    Each segment gets:
      • Ken Burns pan effect (scale-to-108 % + animated crop, fast)
        — odd  segments pan left→right
        — even segments pan right→left
      • 12-frame fade-in and 12-frame fade-out (≈ 0.5 s each)
      • CIPET exam text banner during seconds 1–3
    """
    out_dir = os.path.dirname(output)
    segment_paths = []

    for i, (img_path, audio_path) in enumerate(zip(image_paths, audio_paths)):
        if not os.path.exists(img_path) or not os.path.exists(audio_path):
            print(f"[generator] Missing file for segment {i + 1}, skipping.")
            continue

        duration   = get_dur(audio_path) + 2.0
        total_fr   = max(1, int(duration * 25))
        fade_out_f = max(0, total_fr - 12)
        seg_path   = os.path.join(out_dir, f"seg_{i + 1:02d}.mp4")

        # Ken Burns: scale image to 108% then animate crop position
        scale_w = 2074  # 1920 * 1.083
        scale_h = 1166  # 1080 * 1.080
        pan_range_x = scale_w - 1920   # = 154 px
        pan_range_y = (scale_h - 1080) // 2  # = 43 px

        if i % 2 == 0:
            # Pan left → right
            crop_x = f"({pan_range_x})*t/{duration:.3f}"
            crop_y = str(pan_range_y)
        else:
            # Pan right → left
            crop_x = f"{pan_range_x}-({pan_range_x})*t/{duration:.3f}"
            crop_y = str(pan_range_y)

        vf = (
            f"scale={scale_w}:{scale_h}:force_original_aspect_ratio=increase,"
            f"crop=1920:1080:{crop_x}:{crop_y},"
            f"fade=in:0:12,"
            f"fade=out:{fade_out_f}:12,"
            f"format=yuv420p"
        )

        cmd = [
            "ffmpeg", "-y",
            "-loop", "1", "-i", img_path,
            "-i", audio_path,
            "-c:v", "libx264",
            "-c:a", "aac",
            "-b:a", "192k",
            "-vf", vf,
            "-t", str(duration),
            seg_path,
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        if result.returncode != 0:
            print(f"[generator] ffmpeg segment error (seg {i + 1}): "
                  f"{result.stderr[-STDERR_TAIL:]}")
            continue
        segment_paths.append(seg_path)
        print(f"[generator] Segment {i + 1}/{len(image_paths)} encoded.")

    if not segment_paths:
        raise RuntimeError("No video segments were created.")

    # Concatenate segments
    concat_file = os.path.join(out_dir, "concat.txt")
    with open(concat_file, "w") as fh:
        for seg in segment_paths:
            fh.write(f"file '{seg}'\n")

    cmd = [
        "ffmpeg", "-y",
        "-f", "concat", "-safe", "0",
        "-i", concat_file,
        "-c", "copy",
        output,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    if result.returncode != 0:
        raise RuntimeError(f"ffmpeg concat failed: {result.stderr[-STDERR_TAIL:]}")

    # Cleanup segment files
    for seg in segment_paths:
        try:
            os.remove(seg)
        except OSError:
            pass
    try:
        os.remove(concat_file)
    except OSError:
        pass

    print(f"[generator] Raw video assembled: {output}")


def _postprocess_video(
    input_path: str,
    output_path: str,
    dept_color: tuple,
    work_dir: str,
) -> None:
    """Add progress bar (Upgrade 7) and ambient background music (Upgrade 5).

    Generates a soft ambient sine-wave tone, mixes it at 12 % volume under the
    voice narration, and draws a thin progress bar that grows from left to right
    across the bottom of the video.
    """
    total_dur = get_dur(input_path)
    if total_dur <= 0:
        shutil.copy2(input_path, output_path)
        return

    # Generate ambient background audio
    ambient_wav = os.path.join(work_dir, "ambient.wav")
    try:
        generate_ambient_audio(total_dur, ambient_wav)
    except Exception as ae:
        print(f"[generator] Ambient audio generation failed: {ae}")
        ambient_wav = None

    hex_color = "#{:02X}{:02X}{:02X}".format(*dept_color)

    if ambient_wav and os.path.exists(ambient_wav):
        filter_complex = (
            f"[0:v]drawbox=x=0:y=ih-6:w='min((t/{total_dur:.3f})*iw,iw)':h=6"
            f":color={hex_color}@1.0:t=fill[vout];"
            "[0:a]volume=1.0[voice];"
            "[1:a]volume=0.12[music];"
            "[voice][music]amix=inputs=2:duration=first:dropout_transition=3[aout]"
        )
        cmd = [
            "ffmpeg", "-y",
            "-i", input_path,
            "-i", ambient_wav,
            "-filter_complex", filter_complex,
            "-map", "[vout]",
            "-map", "[aout]",
            "-c:v", "libx264",
            "-c:a", "aac",
            "-b:a", "192k",
            output_path,
        ]
    else:
        # No ambient audio — still add progress bar
        vf = (
            f"drawbox=x=0:y=ih-6:w='min((t/{total_dur:.3f})*iw,iw)':h=6"
            f":color={hex_color}@1.0:t=fill"
        )
        cmd = [
            "ffmpeg", "-y",
            "-i", input_path,
            "-vf", vf,
            "-c:v", "libx264",
            "-c:a", "copy",
            output_path,
        ]

    result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
    if result.returncode != 0:
        print(f"[generator] Post-process warning: {result.stderr[-STDERR_TAIL:]}")
        shutil.copy2(input_path, output_path)
    else:
        print(f"[generator] Post-processed video: {output_path}")

    if ambient_wav:
        try:
            os.remove(ambient_wav)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# Full Pipeline
# ---------------------------------------------------------------------------


def generate_full_video(topic: str, department: str, video_type: str) -> dict:
    """Run the complete video generation pipeline and return metadata."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    work_dir  = f"/videos/{timestamp}_{department.lower()}"
    os.makedirs(work_dir, exist_ok=True)

    img_dir    = os.path.join(work_dir, "images")
    audio_dir  = os.path.join(work_dir, "audio")
    raw_output = os.path.join(work_dir, "lecture_raw.mp4")
    output_path = os.path.join(work_dir, "lecture.mp4")

    os.makedirs(img_dir,   exist_ok=True)
    os.makedirs(audio_dir, exist_ok=True)

    dept_color = DEPT_COLORS_PIL.get(department, DEPT_COLORS_PIL["All"])

    # 1. Generate AI script
    print(f"[generator] Generating AI script for topic='{topic}' dept='{department}'")
    script = ai_script(topic, department, video_type)

    # 2. Create PPT (for download)
    print("[generator] Creating PowerPoint presentation …")
    ppt_path = make_ppt(script, department, work_dir)

    # 3. Intro image + audio
    print("[generator] Rendering intro slide …")
    intro_img = render_intro_image(topic, department)
    intro_img_path = os.path.join(img_dir, "slide_00.png")
    intro_img.save(intro_img_path, "PNG")

    intro_narration = (
        f"Welcome! Today we learn about {topic}. "
        f"This lecture covers the complete technical content for {department} engineering students. "
        "Watch till the end, then subscribe for more lectures like this."
    )
    intro_audio_path = os.path.join(audio_dir, "audio_00.mp3")
    gTTS(text=intro_narration, lang="en", tld="co.in").save(intro_audio_path)
    print("[generator] Intro audio generated.")

    # 4. Render main slide images
    print("[generator] Rendering slide images …")
    slide_img_paths = render_slide_images(script["slides"], department, img_dir)

    # 5. Generate main slide voice narrations
    print("[generator] Generating TTS audio …")
    slide_audio_paths = generate_voices(script["slides"], audio_dir)

    # 6. Outro image + audio
    print("[generator] Rendering outro slide …")
    outro_img = render_outro_image(department)
    outro_img_path = os.path.join(img_dir, "slide_13.png")
    outro_img.save(outro_img_path, "PNG")

    outro_narration = (
        "That wraps up today's lecture. "
        "If this was helpful, please like, subscribe, and share with your classmates. "
        "Subscribe to CIPETHUB for more technical lectures."
    )
    outro_audio_path = os.path.join(audio_dir, "audio_13.mp3")
    gTTS(text=outro_narration, lang="en", tld="co.in").save(outro_audio_path)
    print("[generator] Outro audio generated.")

    # 7. Assemble ordered lists: intro + slides + outro
    all_img_paths   = [intro_img_path] + slide_img_paths + [outro_img_path]
    all_audio_paths = [intro_audio_path] + slide_audio_paths + [outro_audio_path]

    count = min(len(all_img_paths), len(all_audio_paths))
    print(f"[generator] Using {count} segments for video assembly.")

    # 8. Build raw video (Ken Burns + fades, no music yet)
    print("[generator] Assembling raw video …")
    build_video(all_img_paths[:count], all_audio_paths[:count],
                raw_output, dept_color)

    # 9. Post-process: add progress bar + ambient background music
    print("[generator] Adding progress bar and background music …")
    _postprocess_video(raw_output, output_path, dept_color, work_dir)

    # Clean up raw file
    try:
        os.remove(raw_output)
    except OSError:
        pass

    # 10. Collect stats
    duration_sec = get_dur(output_path)
    size_bytes   = os.path.getsize(output_path) if os.path.exists(output_path) else 0

    return {
        "status": "success",
        "video_path": output_path,
        "ppt_path": ppt_path,
        "title": script.get("title", topic),
        "description": script.get("description", ""),
        "tags": script.get("tags", []),
        "topic": topic,
        "department": department,
        "video_type": video_type,
        "slides_count": count,
        "duration_seconds": round(duration_sec, 1),
        "duration_minutes": round(duration_sec / 60, 2),
        "size_mb": round(size_bytes / (1024 * 1024), 2),
        "channel": "CIPETHUB",
        "engine_version": "3.0",
    }
