import io, os, re
import streamlit as st
import pdfplumber
from docx import Document
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from groq import Groq
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

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
# Simple fallback
# ------------------------------
def simple_local_summary(text: str, max_sentences: int = 4) -> str:
    t = re.sub(r"\s+", " ", text).strip()
    sentences = re.split(r'(?<=[.!?])\s+', t)
    return " ".join(sentences[:max_sentences]) or t[:200]

# ------------------------------
# Parse style instructions
# ------------------------------
def parse_styles(instructions: str):
    styles = {"font": "Arial", "font_size": 14, "bg_color": RGBColor(255, 255, 255), "font_color": RGBColor(0, 0, 0)}

    if "blue background" in instructions.lower():
        styles["bg_color"] = RGBColor(0, 102, 204)
    elif "black background" in instructions.lower():
        styles["bg_color"] = RGBColor(0, 0, 0)
        styles["font_color"] = RGBColor(255, 255, 255)
    elif "green background" in instructions.lower():
        styles["bg_color"] = RGBColor(0, 153, 0)

    if "arial" in instructions.lower():
        styles["font"] = "Arial"
    elif "times" in instructions.lower():
        styles["font"] = "Times New Roman"
    elif "calibri" in instructions.lower():
        styles["font"] = "Calibri"

    if "large font" in instructions.lower():
        styles["font_size"] = 20
    elif "small font" in instructions.lower():
        styles["font_size"] = 12

    return styles

# ------------------------------
# Agentic summarizer
# ------------------------------
def summarize_with_agent(text: str, extra_instructions="", model: str = DEFAULT_MODEL, max_chunk_chars=3000):
    slides = []
    if not text.strip():
        return [{"title": "Empty Document", "bullets": ["No extractable text"]}]

    chunks = [text[i:i+max_chunk_chars] for i in range(0, len(text), max_chunk_chars)] if len(text) > max_chunk_chars else [text]

    for idx, chunk in enumerate(chunks, start=1):
        prompt = f"""
        Summarize this text into a PowerPoint slide.
        Provide 1 short title + 4-5 concise bullet points.
        Apply these style instructions if possible: {extra_instructions}
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
            out = simple_local_summary(chunk)

        lines = [l.strip() for l in out.splitlines() if l.strip()]
        title = lines[0] if lines else f"Part {idx}"
        bullets = [l.lstrip("-•* ").strip() for l in lines[1:]] or [out]
        slides.append({"title": title, "bullets": bullets[:6]})

    return slides

# ------------------------------
# PPT generation with styles
# ------------------------------
def make_ppt(slides, styles):
    prs = Presentation()
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "Auto-generated PPT"
    title_slide.placeholders[1].text = "via Groq + Agentic AI"

    for s in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])

        # Set background
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = styles["bg_color"]

        # Title
        title_shape = slide.shapes.title
        title_shape.text = s["title"]
        title_shape.text_frame.paragraphs[0].font.name = styles["font"]
        title_shape.text_frame.paragraphs[0].font.size = Pt(styles["font_size"] + 6)
        title_shape.text_frame.paragraphs[0].font.color.rgb = styles["font_color"]

        # Bullets
        tf = slide.placeholders[1].text_frame
        tf.clear()
        for b in s["bullets"]:
            p = tf.add_paragraph()
            p.text = b
            p.level = 0
            p.font.name = styles["font"]
            p.font.size = Pt(styles["font_size"])
            p.font.color.rgb = styles["font_color"]

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()

# ------------------------------
# PDF merge
# ------------------------------
def make_pdf(all_text: str) -> bytes:
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    y = height - 50
    for line in all_text.splitlines():
        if y < 50:
            c.showPage()
            y = height - 50
        c.drawString(40, y, line[:1000])
        y -= 15
    c.save()
    buffer.seek(0)
    return buffer.read()

# ------------------------------
# Streamlit UI
# ------------------------------
st.title("📄 ➜ Multi-doc to PPT + PDF (Agentic AI)")

files = st.file_uploader("Upload PDF / DOCX / TXT", type=["pdf","docx","txt"], accept_multiple_files=True)
extra_instructions = st.text_area("Style Instructions (e.g., 'blue background, Arial, large font')", "")
model_choice = st.selectbox("Groq model", ["llama-3.1-8b-instant","gemma2-9b-it","mixtral-8x7b"])

if files and st.button("Generate Outputs"):
    all_slides = []
    merged_text = ""
    for f in files:
        text = extract_text(f)
        merged_text += f"\n\n--- {f.name} ---\n\n{text}\n\n"
        summaries = summarize_with_agent(text, extra_instructions, model=model_choice)
        all_slides.extend(summaries)
        st.success(f"Processed {f.name} → {len(summaries)} slides")

    styles = parse_styles(extra_instructions)

    pptx_bytes = make_ppt(all_slides, styles)
    st.download_button("⬇️ Download PPTX", pptx_bytes, file_name="auto_ppt.pptx")

    pdf_bytes = make_pdf(merged_text)
    st.download_button("⬇️ Download Merged PDF", pdf_bytes, file_name="merged.pdf")
