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

    # Background color
    colors = {"dark blue":"#003366","blue":"#3366CC","dark yellow":"#FFCC00","yellow":"#FFFF00",
              "black":"#000000","white":"#FFFFFF","green":"#008000","red":"#FF0000"}
    for k,v in colors.items():
        if k in prompt_l:
            style["background_color"] = v
            break

    # Font
    fonts = ["arial","calibri","times new roman","helvetica","comic sans ms","verdana"]
    for f in fonts:
        if f.lower() in prompt_l:
            style["font"] = f
            break

    # Font size
    m = re.search(r'size[:= ]?(\d+)', prompt_l)
    if m:
        style["font_size"] = int(m.group(1))
    elif "large" in prompt_l or "big" in prompt_l:
        style["font_size"] = 20
    elif "small" in prompt_l:
        style["font_size"] = 12

    # Font color
    for k,v in colors.items():
        if f"color {k}" in prompt_l:
            style["font_color"] = v
            break

    # Emojis
    if "emoji" in prompt_l or "emojis" in prompt_l:
        style["emoji_in_bullets"] = True

    # Footer
    m = re.search(r'footer[:= ]?(.+)', prompt_l)
    if m:
        style["footer_text"] = m.group(1).strip()

    return style

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
        bullets = [clean_text(l.lstrip("-•* ").strip()) for l in lines[1:]] or [clean_text(out)]
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
    fill = title_slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(bg_color.replace("#",""))

    # Content slides
    for s in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
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
            try:
                p.font.color.rgb = RGBColor.from_string(font_color.replace("#",""))
            except:
                pass

        # Footer
        if footer_text:
            p = tf.add_paragraph()
            p.text = clean_text(footer_text)
            p.font.size = Pt(12)
            p.font.color.rgb = RGBColor(150,150,150)

        # Logo
        if logo_file:
            slide.shapes.add_picture(logo_file, Inches(7), Inches(5), Inches(1.2), Inches(1))

        # Background
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor.from_string(bg_color.replace("#",""))

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()

# ------------------------------
# Streamlit UI
# ------------------------------
st.title("📄 Files to PPT Conversion)")

files = st.file_uploader("Upload PDF / DOCX / TXT", type=["pdf","docx","txt"], accept_multiple_files=True)
design_prompt = st.text_area(
    "Design & Styling Instructions",
    "Example:\n- Background: dark blue (#003366)\n- Font: Calibri, size 20, color white\n- Footer: Company Confidential\n- Add emojis to bullets"
)
logo = st.file_uploader("Upload Logo/Image (optional)", type=["png","jpg","jpeg"])
model_choice = st.selectbox("Groq model", ["llama-3.1-8b-instant","gemma2-9b-it","mixtral-8x7b"])

if files and st.button("Generate PPT"):
    all_slides = []
    for f in files:
        text = extract_text(f)
        slides = generate_slide_text(text, model_choice)
        all_slides.extend(slides)

    style = parse_style_from_prompt(design_prompt)
    pptx_bytes = make_ppt(all_slides, style=style, logo_file=logo if logo else None)
    st.download_button("⬇️ Download PPTX", pptx_bytes, file_name="auto_ppt.pptx")
