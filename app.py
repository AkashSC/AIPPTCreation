import io
import os
import re
import json
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
DEFAULT_MODEL = "llama3-8b-8192"
client = Groq(api_key=GROQ_API_KEY)

# ------------------------------
# Utility: text extraction
# ------------------------------
def extract_text_from_pdf(file_bytes: bytes) -> str:
    text_parts = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text()
            if txt:
                text_parts.append(txt)
    return "\n\n".join(text_parts).strip()

def extract_text_from_docx(file_bytes: bytes) -> str:
    doc = Document(io.BytesIO(file_bytes))
    paras = [p.text for p in doc.paragraphs if p.text and p.text.strip()]
    return "\n\n".join(paras).strip()

def extract_text_from_txt(file_bytes: bytes) -> str:
    return file_bytes.decode("utf-8", errors="replace").strip()

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
# Local summarizer fallback
# ------------------------------
def simple_local_summary(text: str, max_sentences: int = 4) -> str:
    t = re.sub(r"\s+", " ", text).strip()
    sentences = re.split(r'(?<=[.!?])\s+', t)
    return " ".join(sentences[:max_sentences]) or (t[:200] + ("..." if len(t) > 200 else ""))

# ------------------------------
# Style parsing helpers
# ------------------------------
COLOR_MAP = {
    "blue": "#003366", "dark blue": "#003366", "black": "#000000", "white": "#FFFFFF",
    "green": "#008000", "red": "#FF0000", "yellow": "#FFCC00", "gray": "#808080",
    "dark": "#333333", "light": "#F8F8F8", "orange": "#FF8C00", "purple": "#800080"
}
FONT_OPTIONS = ["Arial", "Calibri", "Times New Roman", "Helvetica", "Comic Sans MS", "Verdana"]

def parse_design_prompt(prompt: str):
    prompt_l = (prompt or "").lower()
    style = {"background_color": "#FFFFFF", "font": "Arial", "font_size": 14, "font_color": "#000000", "emoji_in_bullets": False}

    # hex color
    found_hex = re.search(r'#([0-9a-fA-F]{6})', prompt)
    if found_hex:
        style["background_color"] = f"#{found_hex.group(1)}"

    # color words
    for name, hx in COLOR_MAP.items():
        if name in prompt_l:
            style["background_color"] = hx
            style["font_color"] = "#FFFFFF" if name in ("blue","dark blue","black","dark","purple") else "#000000"
            break

    # font
    for f in FONT_OPTIONS:
        if f.lower() in prompt_l:
            style["font"] = f
            break

    # font size hints
    if "large" in prompt_l or "big" in prompt_l:
        style["font_size"] = 20
    if "small" in prompt_l:
        style["font_size"] = 12
    if m := re.search(r'font ?size ?[:= ]?(\d{2})', prompt_l):
        try:
            style["font_size"] = int(m.group(1))
        except:
            pass

    # emojis
    if "emoji" in prompt_l or "emojis" in prompt_l:
        style["emoji_in_bullets"] = True

    return style

def extract_style_json_from_text(s: str):
    m = re.search(r'<STYLE_JSON>(.*?)</STYLE_JSON>', s, re.DOTALL | re.IGNORECASE)
    if m:
        try:
            return json.loads(m.group(1).strip())
        except:
            pass
    return None

# ------------------------------
# Slide parsing
# ------------------------------
def parse_slides_from_output(output: str):
    slides = []
    pattern = re.compile(r'(?:Slide Title:|Title:)\s*(.+?)(?:\n|$)([\s\S]*?)(?=(?:Slide Title:|Title:)|$)', re.IGNORECASE)
    matches = pattern.findall(output)
    if matches:
        for title, body in matches:
            bullets = [re.sub(r'^[-•\*\d\)\.]+\s*', '', l.strip()) for l in body.splitlines() if l.strip()]
            slides.append({"title": title.strip(), "bullets": bullets[:6]})
        return slides
    return [{"title": "Summary", "bullets": re.split(r'(?<=[.!?])\s+', simple_local_summary(output, 4))}]

# ------------------------------
# LLM call
# ------------------------------
def summarize_and_style_with_groq(text: str, design_prompt: str, model: str = DEFAULT_MODEL):
    text_for_prompt = text[:3000] + ("\n\n[TRUNCATED]" if len(text) > 3000 else "")
    prompt = f"""
You are a helpful presentation designer. Summarize the document into slides.

Requirements:
1. For each slide, provide a title and 4-5 concise bullet points.
2. Apply ALL design requests: {design_prompt}
3. After slides, include JSON inside <STYLE_JSON>...</STYLE_JSON> with keys:
   background_color, font, font_size, font_color.

Example output:
Slide Title: Example Slide
- Bullet one
- Bullet two

STYLE_JSON:
<STYLE_JSON>{{"background_color":"#003366","font":"Calibri","font_size":18,"font_color":"#FFFFFF"}}</STYLE_JSON>

Document:
{text_for_prompt}
"""
    try:
        chat = client.chat.completions.create(
            model=model,
            messages=[{"role": "system", "content": "You are a presentation designer."},
                      {"role": "user", "content": prompt}],
            temperature=0.3,
            max_tokens=800
        )
        raw = chat.choices[0].message.content
        slides = parse_slides_from_output(raw)
        style_json = extract_style_json_from_text(raw)
        return slides, style_json, True, None
    except Exception as e:
        return [], None, False, str(e)

# ------------------------------
# PPT construction
# ------------------------------
def hex_to_rgb_obj(hex_color: str):
    try:
        h = hex_color.lstrip("#")
        return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))
    except:
        return RGBColor(255,255,255)

def make_ppt(slides, style):
    prs = Presentation()
    for s in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = hex_to_rgb_obj(style.get("background_color","#FFFFFF"))

        title_shape = slide.shapes.title
        title_shape.text = s.get("title","")
        tf = slide.placeholders[1].text_frame
        tf.clear()
        for b in s.get("bullets", []):
            if style.get("emoji_in_bullets"):
                b = "👉 " + b
            p = tf.add_paragraph()
            p.text = b
            p.font.name = style.get("font","Arial")
            p.font.size = Pt(style.get("font_size",14))
            p.font.color.rgb = hex_to_rgb_obj(style.get("font_color","#000000"))

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()

# ------------------------------
# Streamlit UI
# ------------------------------
st.set_page_config(page_title="Agentic PPT Generator", layout="wide")
st.title("📄 ➜ 🖥️ Multi-doc → PPT (Agentic, design prompts applied)")

design_prompt = st.text_area(
    "Design instructions (free-form).",
    value="Blue background, Calibri font, large text, add emojis"
)
st.markdown("""
**Examples:**  
- `Blue gradient background, Calibri, white bold titles, font size 20, add emojis to bullets`  
- `Minimalist white background, Helvetica, small font, black text`  
- `Dark theme (#0f172a), Sans-serif, font size 18, yellow bullet points`  
- `Playful Comic Sans MS, pastel background, emojis in bullets, large font`  
""")

files = st.file_uploader("Upload PDF / DOCX / TXT (multiple)", accept_multiple_files=True)
model_choice = st.selectbox("Groq model", ["llama3-8b-8192","gemma2-9b-it","mixtral-8x7b"])

if files and st.button("Generate PPT"):
    slides_all = []
    global_style = parse_design_prompt(design_prompt)
    for f in files:
        text = extract_text(f)
        if not text:
            continue
        slides, style_json, used, err = summarize_and_style_with_groq(text, design_prompt, model=model_choice)
        if used and slides:
            slides_all.extend(slides)
            if style_json:
                global_style.update(style_json)
        else:
            st.warning(f"Groq failed for {f.name}, fallback used: {err}")
            bullets = re.split(r'(?<=[.!?])\s+', simple_local_summary(text,4))
            slides_all.append({"title":f.name, "bullets":bullets})

    if slides_all:
        pptx_bytes = make_ppt(slides_all, global_style)
        st.download_button("⬇️ Download PPTX", pptx_bytes, "auto_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    else:
        st.error("No slides generated.")
