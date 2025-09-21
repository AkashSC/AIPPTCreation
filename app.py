import io, os, re
import streamlit as st
import pdfplumber
from docx import Document
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from groq import Groq

# ------------------------------
# Config
# ------------------------------
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
DEFAULT_MODEL = "llama-3.1-8b-instant"
client = Groq(api_key=GROQ_API_KEY)

# ------------------------------
# File extractors
# ------------------------------
def extract_text_from_pdf(file_bytes: bytes) -> str:
    text_parts = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            if (txt := page.extract_text()):
                text_parts.append(txt)
    return "\n\n".join(text_parts)

def extract_text_from_docx(file_bytes: bytes) -> str:
    doc = Document(io.BytesIO(file_bytes))
    return "\n\n".join([p.text for p in doc.paragraphs if p.text.strip()])

def extract_text_from_txt(file_bytes: bytes) -> str:
    return file_bytes.decode("utf-8", errors="replace")

def extract_text(uploaded_file):
    data = uploaded_file.read()
    name = uploaded_file.name.lower()
    if name.endswith(".pdf"):
        return extract_text_from_pdf(data)
    elif name.endswith(".docx") or name.endswith(".doc"):
        return extract_text_from_docx(data)
    else:
        return extract_text_from_txt(data)

# ------------------------------
# Cleanup helper
# ------------------------------
def clean_text(text: str) -> str:
    return re.sub(r"<[^>]+>", "", text).strip()

# ------------------------------
# Parse style from user prompt
# ------------------------------
def parse_style_from_prompt(prompt: str):
    style = {
        "background_color": "#FFFFFF",
        "font": "Arial",
        "font_size": 18,
        "font_color": "#000000",
        "emoji_in_bullets": False,
        "footer_text": ""
    }

    prompt_l = prompt.lower()

    colors = {"dark blue":"#003366","blue":"#3366CC","dark yellow":"#FFCC00","yellow":"#FFFF00",
              "black":"#000000","white":"#FFFFFF","green":"#008000","red":"#FF0000"}
    for k,v in colors.items():
        if k in prompt_l:
            style["background_color"] = v
            break

    fonts = ["arial","calibri","times new roman","helvetica","comic sans ms","verdana"]
    for f in fonts:
        if f.lower() in prompt_l:
            style["font"] = f
            break

    m = re.search(r'size[:= ]?(\d+)', prompt_l)
    if m:
        style["font_size"] = int(m.group(1))
    elif "large" in prompt_l or "big" in prompt_l:
        style["font_size"] = 20
    elif "small" in prompt_l:
        style["font_size"] = 12

    for k,v in colors.items():
        if f"color {k}" in prompt_l:
            style["font_color"] = v
            break

    if "emoji" in prompt_l or "emojis" in prompt_l:
        style["emoji_in_bullets"] = True

    m = re.search(r'footer[:= ]?(.+)', prompt_l)
    if m:
        style["footer_text"] = m.group(1).strip()

    return style

# ------------------------------
# Robust hex/named color to RGB
# ------------------------------
def hex_to_rgb_safe(color_str):
    named_colors = {"dark blue":"#003366","blue":"#3366CC","dark yellow":"#FFCC00",
                    "yellow":"#FFFF00","black":"#000000","white":"#FFFFFF",
                    "green":"#008000","red":"#FF0000"}
    color_str = color_str.strip().lower()
    if color_str in named_colors:
        color_str = named_colors[color_str]
    if color_str.startswith("#"):
        color_str = color_str[1:]
    if len(color_str) != 6:
        return RGBColor(255,255,255)
    try:
        r = int(color_str[0:2], 16)
        g = int(color_str[2:4], 16)
        b = int(color_str[4:6], 16)
        return RGBColor(r,g,b)
    except:
        return RGBColor(255,255,255)

# ------------------------------
# Robust bullet parsing
# ------------------------------
def parse_bullets(lines):
    bullets = []
    for line in lines:
        line_clean = clean_text(line.strip())
        if not line_clean:
            continue
        # Ignore common headers
        if re.match(r'^(bullet|points|summary|slide)', line_clean.lower()):
            continue
        # Remove bullets, numbers, or emojis at start
        line_clean = re.sub(r'^([\-•*\d\.\s]+|[\U0001F300-\U0001FAFF]+)\s*', '', line_clean)
        if line_clean:
            bullets.append(line_clean)
    return bullets

# ------------------------------
# Generate slide text using Groq
# ------------------------------
def generate_slide_text(text: str, model: str = DEFAULT_MODEL, max_chunk_chars=3000):
    slides = []
    if not text.strip():
        return [{"title": "Empty Document", "bullets": ["No extractable text"]}]
    chunks = [text[i:i+max_chunk_chars] for i in range(0, len(text), max_chunk_chars)] if len(text) > max_chunk_chars else [text]
    for idx, chunk in enumerate(chunks, start=1):
        prompt = f"""
        Summarize this text into a PowerPoint slide:
        - 1 short title
        - 4-5 concise bullet points

        Text:
        {chunk}
        """
        out = None
        try:
            chat = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3,
                max_tokens=400
            )
            out = chat.choices[0].message.content
        except Exception:
            out = chunk[:200]

        lines = [clean_text(l.strip()) for l in out.splitlines() if l.strip()]
        title = lines[0] if lines else f"Part {idx}"
        bullets = parse_bullets(lines[1:]) or [clean_text(out)]
        slides.append({"title": title, "bullets": bullets[:6]})
    return slides

# ------------------------------
# PPT generator
# ------------------------------
def make_ppt(slides, style=None, logo_file=None):
    prs = Presentation()
    bg_color = style.get("background_color", "#FFFFFF")
    font_name = style.get("font", "Arial")
    font_size = style.get("font_size", 18)
    font_color = style.get("font_color", "#000000")
    emoji = style.get("emoji_in_bullets", False)
    footer_text = style.get("footer_text", "")

    # Title slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "Auto-generated PPT"
    title_slide.placeholders[1].text = "via Groq + Agentic AI"
    title_slide.background.fill.solid()
    title_slide.background.fill.fore_color.rgb = hex_to_rgb_safe(bg_color)

    # Content slides
    for s in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = hex_to_rgb_safe(bg_color)

        slide.shapes.title.text = clean_text(s["title"])
        tf = slide.placeholders[1].text_frame
        tf.clear()
        for b in s["bullets"]:
            if emoji:
                b = "👉 " + b
            p = tf.add_paragraph()
            p.text = clean_text(b)
            p.level = 0
            p.font.size = Pt(font_size)
            p.font.name = font_name
            p.font.color.rgb = hex_to_rgb_safe(font_color)

        if footer_text:
            p = tf.add_paragraph()
            p.text = clean_text(footer_text)
            p.font.size = Pt(12)
            p.font.color.rgb = RGBColor(150,150,150)

        if logo_file:
            slide.shapes.add_picture(logo_file, Inches(7), Inches(5), Inches(1.2), Inches(1))

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()

# ------------------------------
# Streamlit UI
# ------------------------------
st.title("📄 Files to PPT Convertor")

files = st.file_uploader("Upload PDF / DOCX / TXT", type=["pdf","docx","txt"], accept_multiple_files=True)
design_prompt = st.text_area(
    "Design & Styling Instructions",
    "Example:\n- Background: dark blue (#003366)\n- Font: Calibri, size 20, color white\n- Footer: Company Confidential\n- Add emojis to bullets"
)
logo = st.file_uploader("Upload Logo/Image (optional)", type=["png","jpg","jpeg"])

model_choice = DEFAULT_MODEL  # fixed model, dropdown removed

if files and st.button("Generate PPT"):
    all_slides = []
    for f in files:
        text = extract_text(f)
        slides = generate_slide_text(text, model_choice)
        all_slides.extend(slides)

    style = parse_style_from_prompt(design_prompt)
    pptx_bytes = make_ppt(all_slides, style=style, logo_file=logo if logo else None)
    st.download_button("⬇️ Download PPTX", pptx_bytes, file_name="auto_ppt.pptx")
