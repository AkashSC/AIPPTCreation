import io, os, re
import streamlit as st
import pdfplumber
from docx import Document
from pptx import Presentation
from pptx.util import Pt
from groq import Groq

# ------------------------------
# Config
# ------------------------------
GROQ_API_KEY = os.getenv("GROQ_API_KEY")
DEFAULT_MODEL = "llama3-8b-8192"
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
# Agentic summarizer
# ------------------------------
def summarize_with_agent(text: str, model: str = DEFAULT_MODEL, max_chunk_chars=3000):
    slides = []
    if not text.strip():
        return [{"title": "Empty Document", "bullets": ["No extractable text"]}]

    # Decide chunking
    if len(text) > max_chunk_chars:
        chunks = [text[i:i+max_chunk_chars] for i in range(0, len(text), max_chunk_chars)]
    else:
        chunks = [text]

    for idx, chunk in enumerate(chunks, start=1):
        prompt = f"""
        Summarize this text into a PowerPoint slide.
        Give 1 short title + 4-5 concise bullet points:
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
            # Retry once
            try:
                chat = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0.3,
                    max_tokens=300
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
# PPT generation
# ------------------------------
def make_ppt(slides):
    prs = Presentation()
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "Auto-generated PPT"
    title_slide.placeholders[1].text = "via Groq + Agentic AI"

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

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()

# ------------------------------
# Streamlit UI
# ------------------------------
st.title("📄 ➜ 🖥️ Multi-doc to PPT (Groq Agentic AI)")

files = st.file_uploader("Upload PDF / DOCX / TXT", type=["pdf","docx","txt"], accept_multiple_files=True)
model_choice = st.selectbox("Groq model", ["llama3-8b-8192","gemma2-9b-it","mixtral-8x7b"])

if files and st.button("Generate PPT"):
    all_slides = []
    for f in files:
        text = extract_text(f)
        summaries = summarize_with_agent(text, model=model_choice)
        all_slides.extend(summaries)
        st.success(f"Processed {f.name} → {len(summaries)} slides")

    pptx_bytes = make_ppt(all_slides)
    st.download_button("⬇️ Download PPTX", pptx_bytes, file_name="auto_ppt.pptx")
