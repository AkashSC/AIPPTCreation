import io, os, re, json
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
    # remove HTML-like tags e.g. <font>, <b>, <i>
    return re.sub(r"<[^>]+>", "", text).strip()

# ------------------------------
# Local fallback summarizer
# ------------------------------
def simple_local_summary(text: str, max_sentences: int = 4) -> str:
    t = re.sub(r"\s+", " ", text).strip()
    sentences = re.split(r'(?<=[.!?])\s+', t)
    return " ".join(sentences[:max_sentences]) or t[:200]

# ------------------------------
# Summarizer + Style extraction
# ------------------------------
def summarize_and_style(text: str, design_prompt: str, model: str = DEFAULT_MODEL, max_chunk_chars=3000):
    slides, style = [], {}
    if not text.strip():
        return [{"title": "Empty Document", "bullets": ["No extractable text"]}], style

    if len(text) > max_chunk_chars:
        chunks = [text[i:i+max_chunk_chars] for i in range(0, len(text), max_chunk_chars)]
    else:
        chunks = [text]

    for idx, chunk in enumerate(chunks, start=1):
        prompt = f"""
        Summarize this text into a PowerPoint slide.
        - Provide 1 short title
        - 4-5 concise bullet points
        Apply these style instructions if relevant: {design_prompt}

        After the slides, output a JSON inside <STYLE_JSON>...</STYLE_JSON> with:
        background_color (hex like "#003366"), font (string), font_size (int), font_color (hex).

        Example output:
        Slide Title: Example
        - Bullet 1
        - Bullet 2

        STYLE_JSON:
        <STYLE_JSON>{{"background_color":"#003366","font":"Calibri","font_size":18,"font_color":"#FFFFFF"}}</STYLE_JSON>

        Text:
        {chunk}
        """
        out = None
        try:
            chat = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3,
                max_tokens=500
            )
            out = chat.choices[0].message.content
        except Exception:
            out = simple_local_summary(chunk)

        # Parse style JSON if present
        if "<STYLE_JSON>" in out and "</STYLE_JSON>" in out:
            style_block = out.split("<STYLE_JSON>")[1].split("</STYLE_JSON>")[0]
            try:
                style_json = json.loads(style_block)
                style.update(style_json)
            except:
                pass

        # Clean slide text
        lines = [clean_text(l.strip()) for l in out.splitlines() if l.strip() and not l.startswith("<STYLE_JSON>")]
        title = clean_text(lines[0]) if lines else f"Part {idx}"
        bullets = [clean_text(l.lstrip("-•* ").strip()) for l in lines[1:] if not l.startswith("STYLE_JSON")] or [clean_text(out)]
        slides.append({"title": title, "bullets": bullets[:6]})

    return slides, style

# ------------------------------
# PPT generation with style + logo
# ------------------------------
def make_ppt(slides, style=None, logo_file=None):
    prs = Presentation()

    # Default style
    bg_color = style.get("background_color", "#FFFFFF") if style else "#FFFFFF"
    font_name = style.get("font", "Arial") if style else "Arial"
    font_size = style.get("font_size", 18) if style else 18
    font_color = style.get("font_color", "#000000") if style else "#000000"

    # Title slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "Auto-generated PPT"
    title_slide.placeholders[1].text = "via Groq + Agentic AI"

    # Apply background to title slide
    fill = title_slide.background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(bg_color.replace("#", ""))

    # Content slides
    for s in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = clean_text(s["title"])

        tf = slide.placeholders[1].text_frame
        tf.clear()
        for b in s["bullets"]:
            p = tf.add_paragraph()
            p.text = clean_text(b)
            p.level = 0
            p.font.size = Pt(font_size)
            p.font.name = font_name
            try:
                p.font.color.rgb = RGBColor.from_string(font_color.replace("#", ""))
            except:
                pass

        # Inject logo bottom-right
        if logo_file:
            slide.shapes.add_picture(logo_file, Inches(7), Inches(5), Inches(1.2), Inches(1))

        # Background color
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor.from_string(bg_color.replace("#", ""))

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()

# ------------------------------
# Streamlit UI
# ------------------------------
st.title("📄 ➜ 🖥️ Multi-doc to PPT (Groq Agentic AI + Custom Design)")

files = st.file_uploader("Upload PDF / DOCX / TXT", type=["pdf","docx","txt"], accept_multiple_files=True)
design_prompt = st.text_area(
    "Design Instructions (optional)",
    "Examples:\n- Use dark blue background and white text\n- Font: Calibri, size 20\n- Add corporate branding feel\n- Make headings bold and professional"
)
logo = st.file_uploader("Upload Logo/Image (optional)", type=["png","jpg","jpeg"])
model_choice = st.selectbox("Groq model", ["llama-3.1-8b-instant","gemma2-9b-it","mixtral-8x7b"])

if files and st.button("Generate PPT"):
    all_slides, final_style = [], {}
    for f in files:
        text = extract_text(f)
        summaries, style = summarize_and_style(text, design_prompt, model=model_choice)
        all_slides.extend(summaries)
        final_style.update(style)
        st.success(f"Processed {f.name} → {len(summaries)} slides")

    pptx_bytes = make_ppt(all_slides, style=final_style, logo_file=logo if logo else None)
    st.download_button("⬇️ Download PPTX", pptx_bytes, file_name="auto_ppt.pptx")
