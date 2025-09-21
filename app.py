import io, os, re, json
import streamlit as st
import pdfplumber
from docx import Document
from pptx import Presentation
from pptx.util import Pt
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
# LLM-driven summarizer + style
# ------------------------------
def summarize_and_style(text: str, design_prompt: str, model: str = DEFAULT_MODEL):
    """
    Uses LLM to return slides + styling instructions.
    """
    prompt = f"""
    You are a presentation designer.
    Summarize this text into PowerPoint slides.
    - Give 1 short title and 4-5 concise bullet points per slide.
    - Follow these design prompts for styling: {design_prompt}
    - Also return a JSON block with design settings (background_color, font, font_size, font_color).
    
    Example output format:
    ---
    Slide Title: Example
    - Bullet 1
    - Bullet 2
    
    STYLE_JSON:
    {{"background_color":"#003366","font":"Calibri","font_size":18,"font_color":"#FFFFFF"}}
    ---
    Text:
    {text}
    """

    try:
        chat = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.4,
            max_tokens=800
        )
        return chat.choices[0].message.content
    except Exception as e:
        return f"Slide Title: Error\n- Could not summarize\nSTYLE_JSON: {{\"background_color\":\"#FFFFFF\",\"font\":\"Arial\",\"font_size\":14,\"font_color\":\"#000000\"}}"

# ------------------------------
# Extract slides + style JSON
# ------------------------------
def parse_slides_and_style(output: str):
    slides = []
    style = {"background_color": "#FFFFFF", "font": "Arial", "font_size": 14, "font_color": "#000000"}

    parts = output.split("STYLE_JSON:")
    slide_text = parts[0].strip()
    style_text = parts[1].strip() if len(parts) > 1 else ""

    # Parse slides
    for block in re.split(r"Slide Title:", slide_text):
        block = block.strip()
        if not block:
            continue
        lines = block.splitlines()
        title = lines[0].strip()
        bullets = [l.lstrip("-•* ").strip() for l in lines[1:] if l.strip()]
        slides.append({"title": title, "bullets": bullets})

    # Parse style JSON
    try:
        style.update(json.loads(style_text))
    except:
        pass

    return slides, style

# ------------------------------
# PPT generation
# ------------------------------
def hex_to_rgb(hex_color: str):
    hex_color = hex_color.lstrip("#")
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

def make_ppt(slides, style):
    prs = Presentation()
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "Auto-generated PPT"
    title_slide.placeholders[1].text = "via Groq + Agentic AI"

    for s in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])

        # Set background color
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = hex_to_rgb(style.get("background_color", "#FFFFFF"))

        # Title
        title_shape = slide.shapes.title
        title_shape.text = s["title"]
        title_shape.text_frame.paragraphs[0].font.name = style.get("font", "Arial")
        title_shape.text_frame.paragraphs[0].font.size = Pt(style.get("font_size", 14) + 6)
        title_shape.text_frame.paragraphs[0].font.color.rgb = hex_to_rgb(style.get("font_color", "#000000"))

        # Bullets
        tf = slide.placeholders[1].text_frame
        tf.clear()
        for b in s["bullets"]:
            p = tf.add_paragraph()
            p.text = b
            p.level = 0
            p.font.name = style.get("font", "Arial")
            p.font.size = Pt(style.get("font_size", 14))
            p.font.color.rgb = hex_to_rgb(style.get("font_color", "#000000"))

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()

# ------------------------------
# Streamlit UI
# ------------------------------
st.title("📄 ➜ 🖥️ Multi-doc to PPT (Groq Agentic AI)")
st.markdown("""
**Examples for design prompts:**
- "Blue gradient background, white bold titles, Calibri font"
- "Dark theme, yellow bullet points, Comic Sans font, large text"
- "Minimalist white background, Helvetica, black text, small font"
- "Add emojis to bullets, playful design"
""")

files = st.file_uploader("Upload PDF / DOCX / TXT", type=["pdf","docx","txt"], accept_multiple_files=True)
design_prompt = st.text_area("Enter your design instructions", "Blue background, Arial font, large text")
model_choice = st.selectbox("Groq model", ["llama-3.1-8b-instant","gemma2-9b-it","mixtral-8x7b"])

if files and st.button("Generate PPT"):
    all_slides = []
    final_style = None

    for f in files:
        text = extract_text(f)
        raw_output = summarize_and_style(text, design_prompt, model=model_choice)
        slides, style = parse_slides_and_style(raw_output)
        all_slides.extend(slides)
        final_style = style  # take last style block

    pptx_bytes = make_ppt(all_slides, final_style or {})
    st.download_button("⬇️ Download PPTX", pptx_bytes, file_name="auto_ppt.pptx")
