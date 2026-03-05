"""
CIPETHUB Engine — Video Generation Pipeline
Generates educational lecture MP4 videos for CIPET students.
"""

import json
import os
import re
import subprocess
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
    "Plastics": RGBColor(0xFF, 0x6F, 0x00),
    "Mechanical": RGBColor(0x21, 0x96, 0xF3),
    "Manufacturing": RGBColor(0x4C, 0xAF, 0x50),
    "All": RGBColor(0xFF, 0xD5, 0x4F),
}

# Pillow-compatible (R, G, B) tuples for the same department palette
DEPT_COLORS_PIL = {
    "Plastics":      (0xFF, 0x6F, 0x00),
    "Mechanical":    (0x21, 0x96, 0xF3),
    "Manufacturing": (0x4C, 0xAF, 0x50),
    "All":           (0xFF, 0xD5, 0x4F),
}

BG_DARK = RGBColor(0x0D, 0x1B, 0x2A)
CARD_BG = RGBColor(0x15, 0x2A, 0x3D)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xB0, 0xBE, 0xC5)
BOTTOM_BAR_BG = RGBColor(0x07, 0x11, 0x1A)

STDERR_TAIL = 500  # characters to include from end of stderr in error messages

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
        model = genai.GenerativeModel("gemini-1.5-flash")

        prompt = f"""You are a professional Indian engineering professor creating a YouTube lecture video script for CIPET (Central Institute of Petrochemicals Engineering & Technology) students.

Topic: {topic}
Department: {department}  (can be Plastics, Mechanical, Manufacturing, or All)
Video Type: {video_type}  (can be Simulation or PPT Lecture)

Create a complete lecture script for a YouTube video with exactly 12 slides.

IMPORTANT INSTRUCTIONS:
- Address students as "dear students" or "friends"
- Slide 1 narration MUST start with "Namaskar and welcome to CIPETHUB!"
- Slide 12 narration MUST end with "Thank you dear students. Please like and subscribe to CIPETHUB channel. Share with your CIPET classmates. Jai Hind!"
- Include exam tips such as "This is important for your CIPET semester exam"
- Mention relevant Indian companies like Reliance, Supreme, Astral, Tata where applicable
- Write narration in Indian English professor style (4-5 sentences per slide)
- Keep slide titles to a maximum of 6 words
- Each slide has exactly 5 bullet points

Return ONLY valid JSON in this exact format (no markdown, no code blocks):
{{
  "title": "video title here",
  "description": "video description here #CIPETHUB #CIPET #engineering #lecture",
  "tags": ["tag1", "tag2", "tag3", "tag4", "tag5", "tag6", "tag7", "tag8", "tag9", "tag10", "tag11", "tag12", "tag13", "tag14", "tag15", "tag16", "tag17", "tag18", "tag19", "tag20"],
  "slides": [
    {{
      "title": "slide title (max 6 words)",
      "bullets": ["bullet 1", "bullet 2", "bullet 3", "bullet 4", "bullet 5"],
      "narration": "4-5 sentence Indian English narration for this slide.",
      "visual_hint": "description of diagram or visual for this slide"
    }}
  ]
}}

Generate exactly 12 slides. The tags array must have exactly 20 items."""

        response = model.generate_content(prompt)
        raw = response.text.strip()
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
            tags += [topic, department, "CIPET", "engineering"] * 5
        data["tags"] = tags[:20]

        return data

    except Exception as exc:
        import traceback
        print(f"[generator] Gemini error ({type(exc).__name__}): {exc}")
        traceback.print_exc()
        return _fallback(topic, department)


def _fallback(topic: str, department: str) -> dict:
    """Return a 12-slide template script when Gemini is unavailable."""

    def slide(number: int, title_suffix: str, content_key: str) -> dict:
        return {
            "title": f"{title_suffix}",
            "bullets": [
                f"Key concept {number}.1 of {topic}",
                f"Key concept {number}.2 of {topic}",
                f"Key concept {number}.3 of {topic}",
                f"Key concept {number}.4 of {topic}",
                f"Key concept {number}.5 of {topic}",
            ],
            "narration": (
                f"Namaskar and welcome to CIPETHUB! "
                if number == 1
                else ""
            )
            + (
                f"Dear students, today we are going to study {topic} in detail. "
                f"This is an important topic for {department} engineering students at CIPET. "
                f"Please pay close attention as this is important for your CIPET semester exam. "
                f"Let us begin with {title_suffix}."
            ),
            "visual_hint": f"Diagram showing {title_suffix} concept for {topic}",
        }

    slides = [
        {
            "title": f"Introduction to {topic[:30]}",
            "bullets": [
                f"Overview of {topic}",
                f"Importance in {department} engineering",
                "Historical background",
                "Applications in Indian industry",
                "Scope of this lecture",
            ],
            "narration": (
                f"Namaskar and welcome to CIPETHUB! Dear students, today we will study {topic}. "
                f"This topic is extremely important for {department} engineering students. "
                f"Companies like Reliance, Tata, Supreme and Astral use these concepts extensively. "
                f"This is important for your CIPET semester exam. Let us start our journey!"
            ),
            "visual_hint": f"Introduction diagram for {topic} showing key applications",
        },
        {
            "title": "Fundamental Concepts",
            "bullets": [
                "Basic definitions and terminology",
                "Core principles involved",
                "Scientific basis",
                "Standard notations used",
                "Units and measurements",
            ],
            "narration": (
                f"Dear students, let us understand the fundamental concepts of {topic}. "
                "First, we must be clear about the basic definitions and terminology. "
                "These form the foundation of all advanced topics. "
                "Please note these definitions carefully as they are important for your CIPET semester exam."
            ),
            "visual_hint": "Concept map showing fundamental definitions and relationships",
        },
        {
            "title": "Classification & Types",
            "bullets": [
                "Primary classification criteria",
                "Type A — properties and uses",
                "Type B — properties and uses",
                "Type C — properties and uses",
                "Comparison of different types",
            ],
            "narration": (
                f"Friends, now we will look at the classification and types in {topic}. "
                "Understanding the classification helps us choose the right approach for each application. "
                f"In the Indian industry, especially in companies like Reliance Industries, these classifications are used daily. "
                "Make sure you remember the comparison table for your examinations."
            ),
            "visual_hint": "Classification tree diagram showing all major types",
        },
        {
            "title": "Properties & Characteristics",
            "bullets": [
                "Mechanical properties overview",
                "Thermal properties overview",
                "Chemical resistance",
                "Electrical properties",
                "Standard testing methods",
            ],
            "narration": (
                f"Dear students, the properties and characteristics are the heart of {topic}. "
                "We study both mechanical and thermal properties in this section. "
                "Chemical resistance is particularly important in the petrochemicals industry. "
                "These properties determine which material or process we select for a given application."
            ),
            "visual_hint": "Property comparison bar chart showing values for different variants",
        },
        {
            "title": "Manufacturing Process",
            "bullets": [
                "Raw material preparation",
                "Processing steps in sequence",
                "Equipment and machinery used",
                "Quality control checkpoints",
                "Industry safety norms",
            ],
            "narration": (
                f"Friends, the manufacturing process for {topic} involves several well-defined steps. "
                "Starting from raw material preparation to the final product, each step must be carefully controlled. "
                "Companies such as Supreme Industries and Astral follow strict quality control checkpoints. "
                "This is important for your CIPET semester exam — please draw and label the process flow diagram."
            ),
            "visual_hint": "Step-by-step process flow diagram with equipment labels",
        },
        {
            "title": "Equipment & Machinery",
            "bullets": [
                "Main processing equipment",
                "Auxiliary systems",
                "Control and instrumentation",
                "Maintenance requirements",
                "Safety interlocks",
            ],
            "narration": (
                f"Dear students, let us now look at the equipment and machinery used in {topic}. "
                "Every piece of equipment has a specific function in the overall process. "
                "Understanding the control and instrumentation is essential for a {department} engineer. "
                "Tata Engineering and other major corporations use automated systems for precision control."
            ),
            "visual_hint": "Labelled diagram of main processing equipment with part names",
        },
        {
            "title": "Design Calculations",
            "bullets": [
                "Key design parameters",
                "Important formulas to remember",
                "Sample calculation walkthrough",
                "Safety factors and standards",
                "Software tools used in industry",
            ],
            "narration": (
                f"Friends, design calculations are a very important part of {topic}. "
                "You must memorise the key formulas as they are frequently asked in CIPET examinations. "
                "Let us walk through a sample calculation step by step. "
                "Always apply the appropriate safety factor as per Indian Standards."
            ),
            "visual_hint": "Formula sheet with labelled variables and sample calculation table",
        },
        {
            "title": "Industrial Applications",
            "bullets": [
                "Application in packaging industry",
                "Application in automotive sector",
                "Application in construction industry",
                "Application in agriculture",
                "Emerging applications",
            ],
            "narration": (
                f"Dear students, {topic} has a wide range of industrial applications in India and worldwide. "
                "From packaging to the automotive sector, this knowledge is directly applicable. "
                "Companies like Reliance Polymers and Supreme Industries produce millions of products using these principles. "
                "Understanding real applications will help you in placement interviews after CIPET."
            ),
            "visual_hint": "Infographic showing industries using this technology with logos",
        },
        {
            "title": "Quality Standards & Testing",
            "bullets": [
                "Relevant IS and ISO standards",
                "Testing methods and procedures",
                "Acceptance criteria",
                "Documentation requirements",
                "Certification process",
            ],
            "narration": (
                f"Friends, quality standards and testing are essential aspects of {topic}. "
                "In India, the Bureau of Indian Standards (BIS) publishes IS standards that all manufacturers must follow. "
                "Knowing the testing methods is important for your CIPET semester exam and future career. "
                "Always refer to the latest version of the applicable standard in practice."
            ),
            "visual_hint": "Table listing IS/ISO standards with corresponding test methods",
        },
        {
            "title": "Environmental & Safety Aspects",
            "bullets": [
                "Environmental regulations in India",
                "Waste management practices",
                "Worker safety protocols",
                "Green manufacturing initiatives",
                "Sustainability considerations",
            ],
            "narration": (
                f"Dear students, environmental and safety aspects of {topic} are increasingly important. "
                "India has stringent environmental regulations that all industries must comply with. "
                "Sustainable manufacturing is now a key focus for companies like Tata and Reliance. "
                "As future engineers, you must champion green practices in your workplace."
            ),
            "visual_hint": "Green manufacturing cycle diagram showing waste reduction strategies",
        },
        {
            "title": "Recent Advances & Trends",
            "bullets": [
                "Latest research developments",
                "Industry 4.0 integration",
                "Smart manufacturing trends",
                "Future material innovations",
                "Career opportunities in this field",
            ],
            "narration": (
                f"Friends, the field of {topic} is advancing rapidly. "
                "Industry 4.0 and smart manufacturing are transforming how engineers work. "
                "There are excellent career opportunities in this area across India and globally. "
                "CIPET graduates are highly sought after by top companies in this sector."
            ),
            "visual_hint": "Timeline showing evolution of technology with future milestones",
        },
        {
            "title": "Summary & Revision",
            "bullets": [
                "Key topics covered today",
                "Important formulas recap",
                "Exam tips and focus areas",
                "Recommended reference books",
                "Subscribe to CIPETHUB for more!",
            ],
            "narration": (
                f"Dear students, we have now completed our lecture on {topic}. "
                "Let us quickly revise the key concepts we covered today. "
                "Remember to focus on the formulas and process diagrams for your CIPET semester exam. "
                "Thank you dear students. Please like and subscribe to CIPETHUB channel. "
                "Share with your CIPET classmates. Jai Hind!"
            ),
            "visual_hint": "Mind map summarising all key concepts from the lecture",
        },
    ]

    tags = [
        topic, department, "CIPET", "CIPETHUB", "engineering", "lecture",
        "India", "students", "education", "tutorial",
        f"{department} engineering", "CIPET exam", "semester", "study material",
        "petrochemicals", "polymer", "manufacturing", "Plastics", "technical",
        "YouTube lecture",
    ]

    return {
        "title": f"{topic} | {department} Engineering | CIPETHUB",
        "description": (
            f"Complete lecture on {topic} for CIPET {department} Engineering students. "
            f"Covers all important concepts for CIPET semester examination. "
            f"#CIPETHUB #CIPET #{department.replace(' ', '')} #engineering #lecture #India"
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

        # --- Background ---
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = BG_DARK

        # --- Top accent bar (department color, thin) ---
        top_bar = slide.shapes.add_shape(
            1, Inches(0), Inches(0), Inches(16), Inches(0.12)
        )
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = dept_color
        top_bar.line.fill.background()

        # --- Department badge (top left) ---
        badge = slide.shapes.add_shape(
            1, Inches(0.2), Inches(0.18), Inches(2.0), Inches(0.38)
        )
        badge.fill.solid()
        badge.fill.fore_color.rgb = dept_color
        badge.line.fill.background()
        badge_tf = badge.text_frame
        badge_tf.word_wrap = False
        badge_para = badge_tf.paragraphs[0]
        badge_para.alignment = 1  # PP_ALIGN.CENTER
        run = badge_para.add_run()
        run.text = department.upper()
        run.font.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = WHITE

        # --- CIPETHUB branding (top right) ---
        brand_box = slide.shapes.add_textbox(Inches(12.5), Inches(0.15), Inches(3.3), Inches(0.55))
        brand_tf = brand_box.text_frame
        brand_para = brand_tf.paragraphs[0]
        brand_para.alignment = 2  # PP_ALIGN.RIGHT
        run = brand_para.add_run()
        run.text = "CIPETHUB"
        run.font.bold = True
        run.font.size = Pt(18)
        run.font.color.rgb = dept_color

        sub_para = brand_tf.add_paragraph()
        sub_para.alignment = 2  # PP_ALIGN.RIGHT
        sub_run = sub_para.add_run()
        sub_run.text = "CIPET Study Material"
        sub_run.font.size = Pt(9)
        sub_run.font.color.rgb = LIGHT_GRAY

        # --- Slide number (bottom right) ---
        num_box = slide.shapes.add_textbox(Inches(14.8), Inches(8.4), Inches(1.0), Inches(0.4))
        num_tf = num_box.text_frame
        num_para = num_tf.paragraphs[0]
        num_para.alignment = 2  # PP_ALIGN.RIGHT
        run = num_para.add_run()
        run.text = str(idx)
        run.font.bold = True
        run.font.size = Pt(14)
        run.font.color.rgb = dept_color

        # --- Title ---
        title_box = slide.shapes.add_textbox(Inches(0.3), Inches(0.65), Inches(10.5), Inches(0.75))
        title_tf = title_box.text_frame
        title_tf.word_wrap = False
        title_para = title_tf.paragraphs[0]
        run = title_para.add_run()
        run.text = slide_data.get("title", f"Slide {idx}")
        run.font.bold = True
        run.font.size = Pt(36)
        run.font.color.rgb = dept_color

        # Underline bar below title
        underline_bar = slide.shapes.add_shape(
            1, Inches(0.3), Inches(1.45), Inches(10.5), Inches(0.05)
        )
        underline_bar.fill.solid()
        underline_bar.fill.fore_color.rgb = dept_color
        underline_bar.line.fill.background()

        # --- Content card (dark rectangle) ---
        card = slide.shapes.add_shape(
            1, Inches(0.3), Inches(1.55), Inches(10.5), Inches(6.1)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.fill.background()

        # Bullets on top of card
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

        # --- Visual hint box (right side) ---
        hint_box_bg = slide.shapes.add_shape(
            1, Inches(11.0), Inches(1.55), Inches(4.7), Inches(6.1)
        )
        hint_box_bg.fill.solid()
        hint_box_bg.fill.fore_color.rgb = CARD_BG
        hint_box_bg.line.color.rgb = dept_color
        hint_box_bg.line.width = Pt(1)

        hint_label = slide.shapes.add_textbox(Inches(11.05), Inches(1.65), Inches(4.6), Inches(0.4))
        hl_tf = hint_label.text_frame
        hl_para = hl_tf.paragraphs[0]
        hl_para.alignment = 1  # CENTER
        run = hl_para.add_run()
        run.text = "📊 VISUAL"
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

        # --- Bottom bar ---
        bottom_bar = slide.shapes.add_shape(
            1, Inches(0), Inches(8.7), Inches(16), Inches(0.3)
        )
        bottom_bar.fill.solid()
        bottom_bar.fill.fore_color.rgb = BOTTOM_BAR_BG
        bottom_bar.line.fill.background()

        bottom_text = slide.shapes.add_textbox(Inches(0), Inches(8.7), Inches(16), Inches(0.3))
        bt_tf = bottom_text.text_frame
        bt_para = bt_tf.paragraphs[0]
        bt_para.alignment = 1  # CENTER
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
# Slide Image Rendering (Pillow)
# ---------------------------------------------------------------------------

# Scale factor: 1920 px / 16 inches = 120 px per inch
_PX = 120

# Pillow color constants
_BG_DARK    = (0x0D, 0x1B, 0x2A)
_CARD_BG    = (0x15, 0x2A, 0x3D)
_WHITE      = (0xFF, 0xFF, 0xFF)
_LIGHT_GRAY = (0xB0, 0xBE, 0xC5)
_BOTTOM_BG  = (0x07, 0x11, 0x1A)


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


def render_slide_image(
    slide_data: dict,
    department: str,
    slide_index: int,
    total_slides: int,
) -> Image.Image:
    """Render one 1920×1080 PNG slide image using Pillow.

    Args:
        slide_data: Slide dict with keys ``title`` (str), ``bullets`` (list of str),
                    and ``visual_hint`` (str).
        department: Department name (e.g. "Plastics", "Mechanical").  Used to look up
                    the accent color from DEPT_COLORS_PIL.
        slide_index: 1-based index of this slide (shown in the slide-number label).
        total_slides: Total number of slides in the presentation (shown in the
                      slide-number label as ``slide_index/total_slides``).

    Returns:
        A Pillow :class:`~PIL.Image.Image` object (RGB, 1920×1080).
    """
    dept_color = DEPT_COLORS_PIL.get(department, DEPT_COLORS_PIL["All"])
    img = Image.new("RGB", (1920, 1080), _BG_DARK)
    draw = ImageDraw.Draw(img)

    # --- Top accent bar ---
    draw.rectangle([0, 0, 1920, int(0.12 * _PX)], fill=dept_color)

    # --- Department badge (top left) ---
    bx, by, bw, bh = int(0.2 * _PX), int(0.18 * _PX), int(2.0 * _PX), int(0.38 * _PX)
    draw.rectangle([bx, by, bx + bw, by + bh], fill=dept_color)
    badge_font = _load_font(bold=True, size=14)
    badge_text = department.upper()
    bb = draw.textbbox((0, 0), badge_text, font=badge_font)
    tx = bx + (bw - (bb[2] - bb[0])) // 2
    ty = by + (bh - (bb[3] - bb[1])) // 2
    draw.text((tx, ty), badge_text, font=badge_font, fill=_WHITE)

    # --- CIPETHUB branding (top right) ---
    brand_font = _load_font(bold=True, size=26)
    sub_font   = _load_font(bold=False, size=13)
    brand_text = "CIPETHUB"
    sub_text   = "CIPET Study Material"
    brand_bb = draw.textbbox((0, 0), brand_text, font=brand_font)
    sub_bb   = draw.textbbox((0, 0), sub_text,   font=sub_font)
    brand_x = 1920 - int(0.2 * _PX) - (brand_bb[2] - brand_bb[0])
    draw.text((brand_x, int(0.18 * _PX)), brand_text, font=brand_font, fill=dept_color)
    sub_x = 1920 - int(0.2 * _PX) - (sub_bb[2] - sub_bb[0])
    draw.text((sub_x, int(0.18 * _PX) + (brand_bb[3] - brand_bb[1]) + 4),
              sub_text, font=sub_font, fill=_LIGHT_GRAY)

    # --- Slide title ---
    title_font = _load_font(bold=True, size=46)
    title_text = slide_data.get("title", f"Slide {slide_index}")
    tx_start = int(0.3 * _PX)
    ty_start = int(0.65 * _PX)
    draw.text((tx_start, ty_start), title_text, font=title_font, fill=dept_color)

    # --- Title underline bar ---
    ubar_y = int(1.45 * _PX)
    draw.rectangle([int(0.3 * _PX), ubar_y, int(10.8 * _PX), ubar_y + int(0.05 * _PX)],
                   fill=dept_color)

    # --- Content card ---
    card_x1 = int(0.3  * _PX)
    card_y1 = int(1.55 * _PX)
    card_x2 = int(10.8 * _PX)
    card_y2 = int(7.65 * _PX)
    draw.rectangle([card_x1, card_y1, card_x2, card_y2], fill=_CARD_BG)

    # --- Bullet points ---
    bullet_font = _load_font(bold=False, size=28)
    bul_x = int(0.55 * _PX)
    bul_y = int(1.75 * _PX)
    max_bul_w = card_x2 - bul_x - int(0.2 * _PX)
    for bullet in slide_data.get("bullets", []):
        bul_y = _draw_wrapped_text(
            draw, f"\u25b8  {bullet}", bullet_font,
            bul_x, bul_y, max_bul_w, _WHITE, line_spacing=8,
        )
        bul_y += 10  # extra gap between bullets

    # --- Visual hint box (right side) ---
    hint_x1 = int(11.0 * _PX)
    hint_y1 = int(1.55 * _PX)
    hint_x2 = int(15.7 * _PX)
    hint_y2 = int(7.65 * _PX)
    draw.rectangle([hint_x1, hint_y1, hint_x2, hint_y2], fill=_CARD_BG)
    draw.rectangle([hint_x1, hint_y1, hint_x2, hint_y2], outline=dept_color, width=2)

    # Hint label
    hint_label_font = _load_font(bold=True, size=16)
    label_text = "\U0001f4ca VISUAL"
    lb = draw.textbbox((0, 0), label_text, font=hint_label_font)
    lx = hint_x1 + ((hint_x2 - hint_x1) - (lb[2] - lb[0])) // 2
    draw.text((lx, hint_y1 + int(0.1 * _PX)), label_text,
              font=hint_label_font, fill=dept_color)

    # Hint text (word-wrapped)
    hint_text_font = _load_font(bold=False, size=18)
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
    bot_font  = _load_font(bold=False, size=14)
    bot_text  = f"CIPETHUB  \u2022  {department} Engineering  \u2022  Subscribe & Like!"
    bot_bb    = draw.textbbox((0, 0), bot_text, font=bot_font)
    bot_tx    = (1920 - (bot_bb[2] - bot_bb[0])) // 2
    bot_ty    = bot_y + (1080 - bot_y - (bot_bb[3] - bot_bb[1])) // 2
    draw.text((bot_tx, bot_ty), bot_text, font=bot_font, fill=_LIGHT_GRAY)

    # --- Slide number (bottom right) ---
    num_font = _load_font(bold=True, size=20)
    num_text = f"{slide_index}/{total_slides}"
    nb = draw.textbbox((0, 0), num_text, font=num_font)
    draw.text((1920 - int(0.3 * _PX) - (nb[2] - nb[0]), bot_ty),
              num_text, font=num_font, fill=dept_color)

    return img


def render_slide_images(script_slides: list, department: str, img_dir: str) -> list:
    """Render each slide to a 1920×1080 PNG via Pillow and save to *img_dir*."""
    os.makedirs(img_dir, exist_ok=True)
    total = len(script_slides)
    paths = []
    for i, slide_data in enumerate(script_slides, start=1):
        out_path = os.path.join(img_dir, f"slide_{i:02d}.png")
        img = render_slide_image(slide_data, department, i, total)
        img.save(out_path, "PNG")
        paths.append(out_path)
        print(f"[generator] Rendered slide image: slide_{i:02d}.png")
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
# Video Assembly
# ---------------------------------------------------------------------------


def build_video(slides: list, img_dir: str, audio_dir: str, output: str) -> None:
    """Combine slide images and audio into a final MP4 video."""
    segment_paths = []

    for i in range(len(slides)):
        img_path = os.path.join(img_dir, f"slide_{i + 1:02d}.png")
        audio_path = os.path.join(audio_dir, f"audio_{i + 1:02d}.mp3")

        if not os.path.exists(img_path) or not os.path.exists(audio_path):
            print(f"[generator] Missing file for slide {i + 1}, skipping.")
            continue

        duration = get_dur(audio_path) + 2.0
        seg_path = os.path.join(os.path.dirname(output), f"seg_{i + 1:02d}.mp4")

        cmd = [
            "ffmpeg", "-y",
            "-loop", "1", "-i", img_path,
            "-i", audio_path,
            "-c:v", "libx264",
            "-tune", "stillimage",
            "-c:a", "aac",
            "-b:a", "192k",
            "-vf", (
                "scale=1920:1080:force_original_aspect_ratio=decrease,"
                "pad=1920:1080:(ow-iw)/2:(oh-ih)/2:color=#0D1B2A"
            ),
            "-pix_fmt", "yuv420p",
            "-t", str(duration),
            seg_path,
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if result.returncode != 0:
            print(f"[generator] ffmpeg segment error for slide {i + 1}: {result.stderr[-STDERR_TAIL:]}")
            continue
        segment_paths.append(seg_path)

    if not segment_paths:
        raise RuntimeError("No video segments were created.")

    # Create concat list
    concat_file = os.path.join(os.path.dirname(output), "concat.txt")
    with open(concat_file, "w") as fh:
        for seg in segment_paths:
            fh.write(f"file '{seg}'\n")

    # Join segments
    cmd = [
        "ffmpeg", "-y",
        "-f", "concat",
        "-safe", "0",
        "-i", concat_file,
        "-c", "copy",
        output,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    if result.returncode != 0:
        raise RuntimeError(f"ffmpeg concat failed: {result.stderr[-STDERR_TAIL:]}")

    # Cleanup temporary files
    for seg in segment_paths:
        try:
            os.remove(seg)
        except OSError:
            pass
    try:
        os.remove(concat_file)
    except OSError:
        pass

    print(f"[generator] Video built: {output}")


# ---------------------------------------------------------------------------
# Full Pipeline
# ---------------------------------------------------------------------------


def generate_full_video(topic: str, department: str, video_type: str) -> dict:
    """Run the complete video generation pipeline and return metadata."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    work_dir = f"/videos/{timestamp}_{department.lower()}"
    os.makedirs(work_dir, exist_ok=True)

    img_dir = os.path.join(work_dir, "images")
    audio_dir = os.path.join(work_dir, "audio")
    output_path = os.path.join(work_dir, "lecture.mp4")

    # 1. Generate script
    print(f"[generator] Generating AI script for topic='{topic}' dept='{department}'")
    script = ai_script(topic, department, video_type)

    # 2. Create PPT
    print("[generator] Creating PowerPoint presentation …")
    ppt_path = make_ppt(script, department, work_dir)

    # 3. Render slide images with Pillow
    print("[generator] Rendering slide images …")
    images = render_slide_images(script["slides"], department, img_dir)

    # 4. Generate voice narrations
    print("[generator] Generating TTS audio …")
    audios = generate_voices(script["slides"], audio_dir)

    # 5. Handle slide count mismatch
    count = min(len(images), len(audios))
    slides_used = script["slides"][:count]
    print(f"[generator] Using {count} slides for video assembly.")

    # 6. Build video
    print("[generator] Assembling final video …")
    build_video(slides_used, img_dir, audio_dir, output_path)

    # 7. Collect stats
    duration_sec = get_dur(output_path)
    size_bytes = os.path.getsize(output_path) if os.path.exists(output_path) else 0

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
    }
