"""
CIPETHUB Engine — Video Generation Pipeline
Generates educational lecture MP4 videos for CIPET students.
"""

import asyncio
import json
import os
import re
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path

import edge_tts
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

BG_DARK = RGBColor(0x0D, 0x1B, 0x2A)
CARD_BG = RGBColor(0x15, 0x2A, 0x3D)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xB0, 0xBE, 0xC5)
BOTTOM_BAR_BG = RGBColor(0x07, 0x11, 0x1A)

TTS_VOICE_DEFAULT = "en-IN-PrabhatNeural"
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
        print(f"[generator] Gemini error: {exc} — using fallback.")
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
# PPT → Images
# ---------------------------------------------------------------------------


def ppt_to_images(ppt_path: str, img_dir: str) -> list:
    """Convert PPT slides to PNG images using LibreOffice headless."""
    os.makedirs(img_dir, exist_ok=True)
    cmd = [
        "libreoffice", "--headless", "--convert-to", "png",
        "--outdir", img_dir, ppt_path,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    if result.returncode != 0:
        print(f"[generator] LibreOffice stderr: {result.stderr}")
        raise RuntimeError(
            f"LibreOffice PPT-to-PNG conversion failed (cmd={cmd!r}): {result.stderr[-STDERR_TAIL:]}"
        )

    # LibreOffice names files as <basename>_<n>.png or <basename>.png for single slide
    base = Path(ppt_path).stem
    raw_images = sorted(
        Path(img_dir).glob(f"{base}*.png"),
        key=lambda p: p.name,
    )

    # Rename to slide_01.png, slide_02.png, …
    renamed = []
    for i, img in enumerate(raw_images, start=1):
        new_name = Path(img_dir) / f"slide_{i:02d}.png"
        img.rename(new_name)
        renamed.append(str(new_name))

    print(f"[generator] Converted {len(renamed)} slides to images.")
    return renamed


# ---------------------------------------------------------------------------
# Text-to-Speech
# ---------------------------------------------------------------------------


def generate_voices(slides: list, audio_dir: str) -> list:
    """Generate MP3 narration files for each slide using edge-tts."""
    os.makedirs(audio_dir, exist_ok=True)
    voice = os.environ.get("TTS_VOICE_EN", TTS_VOICE_DEFAULT)
    rate = "-5%"

    async def _synthesize_all():
        paths = []
        for i, slide_data in enumerate(slides, start=1):
            narration = slide_data.get("narration", f"Slide {i}.")
            out_path = os.path.join(audio_dir, f"audio_{i:02d}.mp3")
            communicate = edge_tts.Communicate(narration, voice, rate=rate)
            await communicate.save(out_path)
            paths.append(out_path)
            print(f"[generator] Audio generated: audio_{i:02d}.mp3")
        return paths

    return asyncio.run(_synthesize_all())


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

    # 3. Convert PPT to images
    print("[generator] Converting PPT to images …")
    images = ppt_to_images(ppt_path, img_dir)

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
