import io, os, re
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
# Simple local fallback
# ------------------------------
def simple_local_summary(text: str, max_sentences: int = 4) -> str:
    t = re.sub(r"\s+", " ", text).strip()
    sentences = re.split(r'(?<=[.!?])\s+', t)
    return " ".join(sentences[:max_sentences]) or t[:200]

# ------------------------------
# Clean unwanted lines
# ------------------------------
def clean_lines(lines):
    cleaned = []
    for l in lines:
        if re.match(r"^\*?Bullet Points[:：]?\**", l, re.IGNORECASE):
            continue
        if "Bullet Points" in l:
            continue
        # remove emojis and stray ** markers
        l = re.sub(r"[^\w\s,.!?-]", "", l)
        l = l.replace("**", "").strip()
        if l:
            cleaned.append(l)
    return cleaned

# ------------------------------
# Agentic summarizer
# ------------------------------
def summarize_with_agent(text: str, model: str = DEFAULT_MODEL, max_chunk_chars=3000, user_prompt: str = ""):
    slides = []
    if not text.strip():
        return [{"title": "Empty Document", "bullets": ["No extractable text"]}]

    chunks = [text[i:i+max_chunk_chars] for i in range(0, len(text), max_chunk_chars)] if len(text) > max_chunk_chars else [text]

    for idx, chunk in enumerate(chunks, start=1):
        base_prompt = f"""
        Summarize this text into a PowerPoint slide.
        Give 1 short title and 4-5 concise bullet points.
        {f"Apply these design/custom instructions: {user_prompt}" if user_prompt else ""}
        Text:
        {chunk}
        """
        out = None
        try:
            chat = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": base_prompt}],
                temperature=0.3,
                max_tokens=400
            )
            out = chat.choices[0].message.content
        except Exception:
            out = simple_local_summary(chunk)

        lines = [l.strip() for l in out.splitlines() if l.strip()]
        lines = clean_lines(lines)
        title = lines[0] if lines else f"Part {idx}"
        bullets = lines[1:] if len(lines) > 1 else [out]
        slides.append({"title": title, "bullets": bullets[:6]})

    return slides

# ------------------------------
# PPT generation
# ------------------------------
def make_ppt(slides, user_logo=None):
    prs = Presentation()
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "Auto-generated PPT"
    title_slide.placeholders[1].text = "via Groq + Agentic AI"

    if user_logo:
        left = top = Pt(10)
        try:
            slide.shapes.add_picture(user_logo, left, top, height=Pt(40))
        except Exception:
            pass

    for s in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = s["title"]
        tf = slide.placeholders[1].text_frame
        tf.clear()
        for b in s["bullets"]:
            p = tf.add_paragraph()
            p.text = b
            p.level = 0
            p.font.size = Pt(14)
            p.font.color.rgb = RGBColor(0, 0, 0)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()

# ------------------------------
# Streamlit UI
# ------------------------------
st.title("📄 ➜ 🖥️ Multi-doc to PPT (Groq Agentic AI)")

files = st.file_uploader("Upload PDF / DOCX / TXT", type=["pdf","docx","txt"], accept_multiple_files=True)
user_prompt = st.text_area("Optional: Enter custom design/content prompts (e.g., 'Dark green background, Arial font, corporate style')", "")
user_logo = st.file_uploader("Optional: Upload a logo", type=["png","jpg","jpeg"])

if files and st.button("Generate PPT"):
    all_slides = []
    for f in files:
        text = extract_text(f)
        summaries = summarize_with_agent(text, user_prompt=user_prompt)
        all_slides.extend(summaries)
        st.success(f"Processed {f.name} → {len(summaries)} slides")

    logo_path = None
    if user_logo:
        logo_path = f"temp_logo.{user_logo.name.split('.')[-1]}"
        with open(logo_path, "wb") as f:
            f.write(user_logo.read())

    pptx_bytes = make_ppt(all_slides, user_logo=logo_path if user_logo else None)
    st.download_button("⬇️ Download PPTX", pptx_bytes, file_name="auto_ppt.pptx")
