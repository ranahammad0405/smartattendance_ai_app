import streamlit as st
import os
import tempfile
import json
from datetime import datetime
from groq import Groq
from fpdf import FPDF
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re

# ──────────────────────────────────────────────
# PAGE CONFIG
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="SmartAttend - AI Attendance & Notes",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────
# CUSTOM CSS  — Light Mode, Polished UI
# ──────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600;700&display=swap');

:root {
    --bg:       #f5f4f0;
    --surface:  #ffffff;
    --card:     #ffffff;
    --card2:    #f0eeea;
    --border:   #e2ddd8;
    --accent:   #2563eb;
    --accent2:  #7c3aed;
    --accent3:  #059669;
    --text:     #1a1814;
    --muted:    #6b6560;
    --success:  #059669;
    --warning:  #d97706;
    --danger:   #dc2626;
    --gold:     #92400e;
    --shadow:   0 1px 3px rgba(0,0,0,0.08), 0 4px 16px rgba(0,0,0,0.06);
    --shadow-lg:0 8px 32px rgba(0,0,0,0.12);
}

html, body, [data-testid="stAppViewContainer"] {
    background-color: var(--bg) !important;
    font-family: 'DM Sans', sans-serif !important;
    color: var(--text) !important;
}

[data-testid="stSidebar"] {
    background-color: var(--surface) !important;
    border-right: 1px solid var(--border) !important;
    box-shadow: 2px 0 12px rgba(0,0,0,0.05) !important;
}

[data-testid="stSidebar"] * {
    color: var(--text) !important;
}

h1, h2, h3, h4 {
    font-family: 'DM Serif Display', serif !important;
    color: var(--text) !important;
    letter-spacing: -0.02em !important;
}

/* ── TABS ── */
[data-testid="stTabs"] [data-baseweb="tab-list"] {
    background: var(--card2) !important;
    border-radius: 12px !important;
    padding: 4px !important;
    border: 1px solid var(--border) !important;
    gap: 2px !important;
}

[data-testid="stTabs"] [data-baseweb="tab"] {
    border-radius: 9px !important;
    padding: 0.5rem 1.2rem !important;
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 500 !important;
    font-size: 0.9rem !important;
    color: var(--muted) !important;
    transition: all 0.15s ease !important;
}

[data-testid="stTabs"] [aria-selected="true"] {
    background: var(--surface) !important;
    color: var(--accent) !important;
    box-shadow: var(--shadow) !important;
}

/* ── BUTTONS ── */
.stButton > button {
    background: var(--accent) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 0.65rem 1.5rem !important;
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    letter-spacing: 0.01em !important;
    transition: all 0.18s ease !important;
    box-shadow: 0 2px 8px rgba(37,99,235,0.25) !important;
}

.stButton > button:hover {
    background: #1d4ed8 !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 20px rgba(37,99,235,0.35) !important;
}

.stButton > button:active {
    transform: translateY(0) !important;
}

/* ── DOWNLOAD BUTTONS ── */
[data-testid="stDownloadButton"] button {
    background: var(--surface) !important;
    color: var(--accent) !important;
    border: 1.5px solid var(--accent) !important;
    border-radius: 10px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 600 !important;
    transition: all 0.18s ease !important;
}

[data-testid="stDownloadButton"] button:hover {
    background: var(--accent) !important;
    color: white !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 20px rgba(37,99,235,0.25) !important;
}

/* ── INPUTS ── */
[data-testid="stTextInput"] input,
[data-testid="stTextArea"] textarea {
    background: var(--surface) !important;
    color: var(--text) !important;
    border: 1.5px solid var(--border) !important;
    border-radius: 10px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.93rem !important;
    transition: border-color 0.15s ease !important;
}

[data-testid="stTextInput"] input:focus,
[data-testid="stTextArea"] textarea:focus {
    border-color: var(--accent) !important;
    box-shadow: 0 0 0 3px rgba(37,99,235,0.12) !important;
}

/* ── FILE UPLOADER ── */
[data-testid="stFileUploaderDropzone"] {
    background: var(--surface) !important;
    border: 2px dashed var(--border) !important;
    border-radius: 14px !important;
    transition: border-color 0.15s ease !important;
}

[data-testid="stFileUploaderDropzone"]:hover {
    border-color: var(--accent) !important;
}

/* ── EXPANDERS ── */
[data-testid="stExpander"] {
    background: var(--surface) !important;
    border: 1px solid var(--border) !important;
    border-radius: 12px !important;
    box-shadow: var(--shadow) !important;
}

/* ── SELECT / DATE ── */
[data-testid="stDateInput"] input {
    background: var(--surface) !important;
    border: 1.5px solid var(--border) !important;
    border-radius: 10px !important;
    color: var(--text) !important;
}

/* ── CHECKBOXES ── */
[data-testid="stCheckbox"] label {
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.95rem !important;
    color: var(--text) !important;
}

/* ── METRICS ── */
[data-testid="stMetric"] {
    background: var(--surface) !important;
    border: 1px solid var(--border) !important;
    border-radius: 12px !important;
    padding: 1rem 1.2rem !important;
    box-shadow: var(--shadow) !important;
}

/* ── CUSTOM COMPONENTS ── */
.hero-card {
    background: linear-gradient(135deg, #1e40af 0%, #3730a3 50%, #6d28d9 100%);
    border-radius: 18px;
    padding: 2rem 2.4rem;
    margin-bottom: 1.8rem;
    box-shadow: 0 8px 40px rgba(37,99,235,0.25);
    position: relative;
    overflow: hidden;
}

.hero-card::before {
    content: '';
    position: absolute;
    top: -40px; right: -40px;
    width: 180px; height: 180px;
    border-radius: 50%;
    background: rgba(255,255,255,0.06);
}

.hero-card::after {
    content: '';
    position: absolute;
    bottom: -60px; left: 30%;
    width: 240px; height: 240px;
    border-radius: 50%;
    background: rgba(255,255,255,0.04);
}

.hero-card h1 {
    color: white !important;
    font-size: 2.1rem !important;
    margin: 0 !important;
    font-family: 'DM Serif Display', serif !important;
}

.hero-card p {
    color: rgba(255,255,255,0.78) !important;
    margin: 0.5rem 0 0 0 !important;
    font-size: 1rem !important;
}

.metric-box {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 1.4rem 1.2rem;
    text-align: center;
    box-shadow: var(--shadow);
    transition: transform 0.15s ease, box-shadow 0.15s ease;
}

.metric-box:hover {
    transform: translateY(-2px);
    box-shadow: var(--shadow-lg);
}

.metric-box .metric-val {
    font-family: 'DM Serif Display', serif;
    font-size: 2.6rem;
    font-weight: 700;
    color: var(--accent);
    line-height: 1;
}

.metric-box .metric-label {
    font-size: 0.78rem;
    color: var(--muted);
    text-transform: uppercase;
    letter-spacing: 0.09em;
    margin-top: 0.4rem;
    font-weight: 600;
}

.section-header {
    font-family: 'DM Serif Display', serif;
    font-size: 1.25rem;
    font-weight: 400;
    color: var(--text);
    margin: 1.6rem 0 0.9rem 0;
    padding-bottom: 0.5rem;
    border-bottom: 2px solid var(--border);
    display: flex;
    align-items: center;
    gap: 0.5rem;
}

.student-chip {
    display: inline-block;
    background: #ecfdf5;
    border: 1px solid #6ee7b7;
    color: #065f46;
    border-radius: 20px;
    padding: 0.25rem 0.9rem;
    font-size: 0.83rem;
    margin: 0.2rem;
    font-weight: 500;
    transition: all 0.15s ease;
}

.student-chip:hover {
    background: #d1fae5;
    border-color: #34d399;
}

.notes-preview {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 1.8rem;
    white-space: pre-wrap;
    font-family: 'DM Sans', sans-serif;
    font-size: 0.93rem;
    line-height: 1.8;
    max-height: 450px;
    overflow-y: auto;
    color: var(--text);
    box-shadow: var(--shadow);
}

.notes-preview::-webkit-scrollbar {
    width: 6px;
}
.notes-preview::-webkit-scrollbar-track {
    background: var(--card2);
    border-radius: 3px;
}
.notes-preview::-webkit-scrollbar-thumb {
    background: var(--border);
    border-radius: 3px;
}

.info-banner {
    background: #eff6ff;
    border: 1px solid #bfdbfe;
    border-left: 4px solid var(--accent);
    border-radius: 10px;
    padding: 0.9rem 1.2rem;
    font-size: 0.9rem;
    color: #1e40af;
    margin-bottom: 1rem;
    display: flex;
    align-items: center;
    gap: 0.6rem;
}

.success-banner {
    background: #f0fdf4;
    border: 1px solid #bbf7d0;
    border-left: 4px solid var(--success);
    border-radius: 10px;
    padding: 0.9rem 1.2rem;
    font-size: 0.9rem;
    color: #14532d;
    margin-bottom: 1rem;
}

.warning-banner {
    background: #fffbeb;
    border: 1px solid #fde68a;
    border-left: 4px solid var(--warning);
    border-radius: 10px;
    padding: 0.9rem 1.2rem;
    font-size: 0.9rem;
    color: #78350f;
    margin-bottom: 1rem;
}

.secure-badge {
    background: #f0fdf4;
    border: 1px solid #86efac;
    border-radius: 8px;
    padding: 0.6rem 1rem;
    font-size: 0.82rem;
    color: #15803d;
    margin-bottom: 1.2rem;
    display: flex;
    align-items: center;
    gap: 0.5rem;
    font-weight: 500;
}

.sidebar-logo {
    font-family: 'DM Serif Display', serif;
    font-size: 1.6rem;
    color: var(--accent) !important;
    margin-bottom: 0.1rem;
}

.sidebar-sub {
    font-size: 0.78rem;
    color: var(--muted) !important;
    margin-top: -0.3rem;
    margin-bottom: 1rem;
    font-weight: 400;
}

.history-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 1.1rem 1.4rem;
    margin-bottom: 0.8rem;
    box-shadow: var(--shadow);
    display: flex;
    align-items: center;
    gap: 1rem;
}

.tag-pill {
    display: inline-block;
    padding: 0.2rem 0.75rem;
    border-radius: 20px;
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.03em;
}

.tag-blue { background: #dbeafe; color: #1d4ed8; }
.tag-green { background: #dcfce7; color: #15803d; }
.tag-purple { background: #f3e8ff; color: #7e22ce; }
.tag-amber { background: #fef3c7; color: #92400e; }

/* Divider */
hr {
    border: none !important;
    border-top: 1px solid var(--border) !important;
    margin: 1.2rem 0 !important;
}

/* Alert colors fix */
[data-testid="stAlert"] {
    border-radius: 10px !important;
}

/* Spinner */
[data-testid="stSpinner"] {
    color: var(--accent) !important;
}
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# API KEY
# ──────────────────────────────────────────────
api_key = st.secrets.get("GROQ_API_KEY", "")
if not api_key:
    st.error(
        "⚠️ **GROQ_API_KEY not found.**\n\n"
        "In Streamlit Cloud go to: **App Settings → Secrets** and add:\n\n"
        "```\nGROQ_API_KEY = \"gsk_your_key_here\"\n```"
    )
    st.stop()

# ──────────────────────────────────────────────
# SESSION STATE
# ──────────────────────────────────────────────
for key, default in {
    "transcript": "",
    "summary": "",
    "lecture_notes": "",
    "attendance_list": [],
    "class_name": "",
    "teacher_name": "",
    "session_date": datetime.today().strftime("%Y-%m-%d"),
    "history": [],
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ──────────────────────────────────────────────
# GROQ CLIENT
# ──────────────────────────────────────────────
@st.cache_resource
def get_groq_client(key: str):
    return Groq(api_key=key)


# ──────────────────────────────────────────────
# SAFE TRUNCATE
# ──────────────────────────────────────────────
def safe_truncate(text: str, max_words: int = 1800) -> str:
    words = text.split()
    if len(words) <= max_words:
        return text
    return " ".join(words[:max_words]) + "\n\n[Transcript truncated to fit model context window]"


# ──────────────────────────────────────────────
# TEXT SANITIZER  — strip non-latin1 chars for fpdf
# ──────────────────────────────────────────────
def sanitize_for_pdf(text: str) -> str:
    """Replace or remove characters outside latin-1 range to prevent fpdf errors."""
    result = []
    for ch in text:
        code = ord(ch)
        if code < 256:
            result.append(ch)
        else:
            # Replace common unicode punctuation with ascii equivalents
            replacements = {
                '\u2018': "'", '\u2019': "'", '\u201c': '"', '\u201d': '"',
                '\u2013': '-', '\u2014': '--', '\u2026': '...', '\u2022': '*',
                '\u00b7': '*', '\u2192': '->', '\u2190': '<-', '\u2022': '-',
                '\u25cf': '*', '\u25b6': '>',
            }
            result.append(replacements.get(ch, '?'))
    return ''.join(result)


def clean_markdown(text: str) -> str:
    """Remove markdown symbols for plain-text rendering."""
    text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
    text = re.sub(r'\*(.*?)\*', r'\1', text)
    text = re.sub(r'`(.*?)`', r'\1', text)
    return text


# ──────────────────────────────────────────────
# LLM HELPERS
# ──────────────────────────────────────────────
def transcribe_audio(client, audio_path: str) -> str:
    with open(audio_path, "rb") as f:
        result = client.audio.transcriptions.create(
            model="whisper-large-v3",
            file=f,
            response_format="text",
        )
    return result


def llm_call(client, system: str, user: str, max_tokens: int = 2048) -> str:
    response = client.chat.completions.create(
        model="llama-3.1-8b-instant",
        messages=[
            {"role": "system", "content": system},
            {"role": "user",   "content": user},
        ],
        temperature=0.4,
        max_tokens=max_tokens,
    )
    return response.choices[0].message.content.strip()


def summarize_transcript(client, transcript: str) -> str:
    system = (
        "You are an expert academic summarizer. "
        "Given a classroom lecture transcript, produce a clear concise summary "
        "(3-5 paragraphs) covering the main topics, key arguments, and conclusions. "
        "Use formal academic language. Use only ASCII-compatible characters."
    )
    return llm_call(client, system, "Transcript:\n\n" + safe_truncate(transcript))


def generate_lecture_notes(client, transcript: str, class_name: str) -> str:
    cn = class_name if class_name else "Class"
    system = (
        "You are a professional academic note-taker. "
        "Convert the following lecture transcript into structured well-formatted lecture notes. "
        "Use this exact structure:\n"
        "# Lecture Notes - " + cn + "\n"
        "## 1. Overview\n"
        "## 2. Key Concepts\n"
        "## 3. Detailed Breakdown\n"
        "   - Sub-topics with bullet points\n"
        "## 4. Important Definitions\n"
        "## 5. Examples and Case Studies (if any)\n"
        "## 6. Summary and Takeaways\n"
        "## 7. Possible Exam Questions\n"
        "Be thorough, organized, and use markdown formatting. "
        "Use only ASCII-compatible characters, no special unicode symbols."
    )
    return llm_call(client, system, "Transcript:\n\n" + safe_truncate(transcript))


def extract_attendance(client, transcript: str) -> list:
    system = (
        "You are an attendance assistant. "
        "Extract the names of all students mentioned as PRESENT in the following transcript. "
        "Return ONLY a valid JSON array of strings. No explanation, no markdown, no code fences. "
        "If no names are found return an empty array []. "
        'Example: ["Alice Smith", "Bob Jones"]'
    )
    raw = llm_call(
        client,
        system,
        "Transcript:\n\n" + safe_truncate(transcript, max_words=1200),
        max_tokens=512,
    )
    try:
        cleaned = raw.strip().replace("```json", "").replace("```", "").strip()
        start = cleaned.find("[")
        end   = cleaned.rfind("]")
        if start != -1 and end != -1 and end > start:
            result = json.loads(cleaned[start : end + 1])
            if isinstance(result, list):
                return [str(n).strip() for n in result if str(n).strip()]
        return []
    except Exception:
        return []


# ──────────────────────────────────────────────
# PDF EXPORT  — FIXED
# ──────────────────────────────────────────────
class SmartAttendPDF(FPDF):
    """Custom PDF class with header/footer."""
    
    def __init__(self, meta):
        super().__init__()
        self.meta = meta
        self.set_margins(20, 20, 20)          # left, top, right — wider margins
        self.set_auto_page_break(auto=True, margin=20)

    def header(self):
        # Blue header bar
        self.set_fill_color(37, 99, 235)
        self.rect(0, 0, 210, 14, 'F')
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(255, 255, 255)
        self.set_y(3)
        self.cell(0, 8, "SmartAttend - AI Classroom Assistant", align="C")
        self.set_text_color(0, 0, 0)
        self.ln(10)

    def footer(self):
        self.set_y(-13)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 8, f"Page {self.page_no()} | Generated {self.meta.get('date', '')}", align="C")


def export_pdf(notes: str, meta: dict) -> bytes:
    pdf = SmartAttendPDF(meta)
    pdf.add_page()

    # ── Title block ──
    pdf.set_fill_color(239, 246, 255)
    pdf.set_draw_color(191, 219, 254)
    pdf.rect(15, pdf.get_y(), 180, 22, 'FD')
    pdf.set_font("Helvetica", "B", 15)
    pdf.set_text_color(30, 64, 175)
    pdf.set_y(pdf.get_y() + 3)
    safe_class = sanitize_for_pdf(meta.get("class_name") or "Lecture")
    pdf.cell(0, 8, safe_class, align="C", ln=True)
    
    info_line = (
        "Teacher: " + sanitize_for_pdf(meta.get("teacher_name") or "-") +
        "   |   Date: " + sanitize_for_pdf(meta.get("date") or "-")
    )
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(100, 116, 139)
    pdf.cell(0, 7, info_line, align="C", ln=True)
    pdf.ln(6)

    # ── Page usable width (important for multi_cell) ──
    # With left=20, right=20 margins on A4 (210mm wide): usable = 170mm
    usable_w = pdf.w - pdf.l_margin - pdf.r_margin  # dynamic calculation

    # ── Render notes line by line ──
    for raw_line in notes.split("\n"):
        line = sanitize_for_pdf(raw_line).strip()

        if not line:
            pdf.ln(3)
            continue

        if line.startswith("# "):
            content = clean_markdown(line[2:]).strip()
            pdf.set_fill_color(239, 246, 255)
            pdf.set_draw_color(191, 219, 254)
            pdf.set_font("Helvetica", "B", 13)
            pdf.set_text_color(30, 64, 175)
            pdf.multi_cell(usable_w, 9, content, border=0, fill=False)
            # Draw underline manually
            pdf.set_draw_color(191, 219, 254)
            pdf.set_line_width(0.4)
            y = pdf.get_y()
            pdf.line(pdf.l_margin, y, pdf.l_margin + usable_w, y)
            pdf.ln(3)

        elif line.startswith("## "):
            content = clean_markdown(line[3:]).strip()
            pdf.set_font("Helvetica", "B", 11)
            pdf.set_text_color(37, 99, 235)
            pdf.multi_cell(usable_w, 8, content, border=0)
            pdf.ln(1)

        elif line.startswith("### "):
            content = clean_markdown(line[4:]).strip()
            pdf.set_font("Helvetica", "B", 10)
            pdf.set_text_color(55, 65, 81)
            pdf.multi_cell(usable_w, 7, content, border=0)

        elif line.startswith("- ") or line.startswith("* "):
            content = clean_markdown(line[2:]).strip()
            pdf.set_font("Helvetica", "", 10)
            pdf.set_text_color(30, 30, 30)
            # Bullet indent: use x offset
            bullet_x = pdf.l_margin + 4
            text_x   = pdf.l_margin + 10
            text_w   = usable_w - 10   # account for indent
            
            y_before = pdf.get_y()
            pdf.set_x(bullet_x)
            pdf.cell(4, 7, chr(149))  # bullet dot
            pdf.set_x(text_x)
            pdf.multi_cell(text_w, 7, content, border=0)

        elif line.startswith("   - ") or line.startswith("   * "):
            content = clean_markdown(line[5:]).strip()
            pdf.set_font("Helvetica", "", 10)
            pdf.set_text_color(55, 65, 81)
            pdf.set_x(pdf.l_margin + 16)
            pdf.multi_cell(usable_w - 16, 6.5, "  - " + content, border=0)

        else:
            content = clean_markdown(line).strip()
            if not content:
                continue
            pdf.set_font("Helvetica", "", 10)
            pdf.set_text_color(30, 30, 30)
            pdf.multi_cell(usable_w, 7, content, border=0)

    return bytes(pdf.output())


# ──────────────────────────────────────────────
# DOCX EXPORT
# ──────────────────────────────────────────────
def export_docx(notes: str, meta: dict) -> bytes:
    doc = Document()

    # Set margins
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.2)
        section.right_margin  = Inches(1.2)

    # Title
    title = doc.add_heading("SmartAttend — Lecture Notes", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if title.runs:
        title.runs[0].font.color.rgb = RGBColor(30, 64, 175)
        title.runs[0].font.size = Pt(24)

    # Meta info
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info = (
        "Class: " + (meta.get("class_name") or "-") +
        "   |   Teacher: " + (meta.get("teacher_name") or "-") +
        "   |   Date: " + (meta.get("date") or "-")
    )
    run = p.add_run(info)
    run.font.color.rgb = RGBColor(100, 116, 139)
    run.font.size = Pt(10)
    doc.add_paragraph()

    for line in notes.split("\n"):
        ls = line.strip()
        if not ls:
            doc.add_paragraph()
            continue
        if ls.startswith("# "):
            h = doc.add_heading(ls[2:], level=1)
            if h.runs:
                h.runs[0].font.color.rgb = RGBColor(30, 64, 175)
                h.runs[0].font.size = Pt(16)
        elif ls.startswith("## "):
            h = doc.add_heading(ls[3:], level=2)
            if h.runs:
                h.runs[0].font.color.rgb = RGBColor(37, 99, 235)
                h.runs[0].font.size = Pt(13)
        elif ls.startswith("### "):
            h = doc.add_heading(ls[4:], level=3)
            if h.runs:
                h.runs[0].font.size = Pt(12)
        elif ls.startswith("- ") or ls.startswith("* "):
            p = doc.add_paragraph(ls[2:], style="List Bullet")
            if p.runs:
                p.runs[0].font.size = Pt(11)
        else:
            p = doc.add_paragraph(ls)
            if p.runs:
                p.runs[0].font.size = Pt(11)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ──────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="sidebar-logo">🎓 SmartAttend</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-sub">AI-Powered Classroom Assistant</div>', unsafe_allow_html=True)
    
    st.markdown(
        '<div class="secure-badge">🔒 API key loaded securely from Secrets</div>',
        unsafe_allow_html=True,
    )

    st.markdown("#### 📋 Session Details")
    st.session_state.class_name   = st.text_input("Class / Subject", placeholder="e.g. Data Structures CS301", key="cn")
    st.session_state.teacher_name = st.text_input("Teacher Name",    placeholder="e.g. Dr. Ahmed Khan",      key="tn")
    date_val = st.date_input("Session Date", value=datetime.today(), key="sd")
    st.session_state.session_date = date_val.strftime("%Y-%m-%d")

    st.divider()

    # Quick stats
    if st.session_state.history:
        st.markdown("#### 📊 Quick Stats")
        total_sessions  = len(st.session_state.history)
        total_students  = sum(s["students_present"] for s in st.session_state.history)
        notes_generated = sum(1 for s in st.session_state.history if s["has_notes"])
        
        cols = st.columns(3)
        cols[0].metric("Sessions",  total_sessions)
        cols[1].metric("Students",  total_students)
        cols[2].metric("Notes",     notes_generated)
        st.divider()

    st.markdown(
        "<p style='font-size:0.76rem;color:#94a3b8;margin-top:0.5rem'>"
        "Powered by Groq · Whisper · LLaMA 3.1"
        "</p>",
        unsafe_allow_html=True,
    )


# ──────────────────────────────────────────────
# HERO SECTION
# ──────────────────────────────────────────────
st.markdown("""
<div class="hero-card">
  <h1>📋 AI Classroom Attendance &amp; Notes</h1>
  <p>Upload a lecture recording → instant transcription · attendance tracking · smart notes · PDF &amp; DOCX export</p>
</div>
""", unsafe_allow_html=True)

tab1, tab2, tab3, tab4 = st.tabs([
    "🎤  Upload & Process",
    "📝  Lecture Notes",
    "👥  Attendance",
    "📚  History",
])

# ══════════════════════════════════════════════
# TAB 1 — Upload & Process
# ══════════════════════════════════════════════
with tab1:
    col_left, col_right = st.columns([1.1, 0.9], gap="large")

    with col_left:
        st.markdown('<div class="section-header">🎙️ Upload Audio / Video</div>', unsafe_allow_html=True)
        audio_file = st.file_uploader(
            "Drop lecture recording here",
            type=["mp3", "mp4", "wav", "m4a", "ogg", "webm"],
            label_visibility="collapsed",
        )
        if audio_file:
            st.audio(audio_file)
            st.markdown(
                f'<div class="success-banner">✅ File ready: <strong>{audio_file.name}</strong> ({audio_file.size // 1024} KB)</div>',
                unsafe_allow_html=True,
            )

        st.markdown('<div class="section-header">📄 Or Paste Transcript</div>', unsafe_allow_html=True)
        manual_text = st.text_area(
            "Manual transcript",
            height=170,
            placeholder="Paste lecture transcript here if you already have one...",
            label_visibility="collapsed",
        )

    with col_right:
        st.markdown('<div class="section-header">⚙️ Processing Options</div>', unsafe_allow_html=True)
        
        st.markdown('<div style="height:0.4rem"></div>', unsafe_allow_html=True)
        do_attendance = st.checkbox("👥  Extract attendance from transcript", value=True)
        do_summary    = st.checkbox("📄  Generate lecture summary",           value=True)
        do_notes      = st.checkbox("📝  Generate full lecture notes",        value=True)

        st.markdown('<div style="height:0.8rem"></div>', unsafe_allow_html=True)

        st.markdown(
            '<div class="info-banner">ℹ️ Processing uses Whisper for transcription and LLaMA 3.1 for analysis.</div>',
            unsafe_allow_html=True,
        )
        
        process_btn = st.button("⚡ Process Lecture", use_container_width=True)

        if st.session_state.transcript:
            st.markdown('<div style="height:1rem"></div>', unsafe_allow_html=True)
            wc = len(st.session_state.transcript.split())
            nl = len(st.session_state.lecture_notes.split("\n")) if st.session_state.lecture_notes else 0
            
            st.markdown('<div class="section-header">📊 Session Metrics</div>', unsafe_allow_html=True)
            mc1, mc2, mc3 = st.columns(3)
            with mc1:
                st.markdown(f'<div class="metric-box"><div class="metric-val">{wc:,}</div><div class="metric-label">Words</div></div>', unsafe_allow_html=True)
            with mc2:
                st.markdown(f'<div class="metric-box"><div class="metric-val">{len(st.session_state.attendance_list)}</div><div class="metric-label">Present</div></div>', unsafe_allow_html=True)
            with mc3:
                st.markdown(f'<div class="metric-box"><div class="metric-val">{nl}</div><div class="metric-label">Note Lines</div></div>', unsafe_allow_html=True)

    # ── Process button logic ──
    if process_btn:
        if not audio_file and not manual_text.strip():
            st.warning("⚠️ Please upload an audio file or paste a transcript before processing.")
        else:
            client = get_groq_client(api_key)

            progress_bar = st.progress(0, text="Starting...")

            # Transcription
            if audio_file and not manual_text.strip():
                progress_bar.progress(10, text="🎙️ Transcribing audio with Whisper...")
                suffix = "." + audio_file.name.split(".")[-1]
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                    tmp.write(audio_file.read())
                    tmp_path = tmp.name
                try:
                    st.session_state.transcript = transcribe_audio(client, tmp_path)
                except Exception as e:
                    st.error(f"Transcription failed: {e}")
                    st.stop()
                finally:
                    try:
                        os.unlink(tmp_path)
                    except Exception:
                        pass
            else:
                st.session_state.transcript = manual_text.strip()

            progress_bar.progress(30, text="✅ Transcript ready")

            # Attendance
            if do_attendance:
                progress_bar.progress(45, text="👥 Extracting attendance...")
                try:
                    st.session_state.attendance_list = extract_attendance(
                        client, st.session_state.transcript
                    )
                except Exception as e:
                    st.warning(f"Attendance extraction failed ({e}). Add names manually in the Attendance tab.")
                    st.session_state.attendance_list = []

            progress_bar.progress(60, text="📄 Summarizing lecture...")

            # Summary
            if do_summary:
                try:
                    st.session_state.summary = summarize_transcript(
                        client, st.session_state.transcript
                    )
                except Exception as e:
                    st.warning(f"Summary failed: {e}")

            progress_bar.progress(80, text="📝 Generating lecture notes...")

            # Notes
            if do_notes:
                try:
                    st.session_state.lecture_notes = generate_lecture_notes(
                        client, st.session_state.transcript, st.session_state.class_name
                    )
                except Exception as e:
                    st.warning(f"Notes generation failed: {e}")

            progress_bar.progress(100, text="🎉 All done!")

            st.session_state.history.append({
                "date":             st.session_state.session_date,
                "class":            st.session_state.class_name or "Unnamed Class",
                "teacher":          st.session_state.teacher_name or "Unknown",
                "students_present": len(st.session_state.attendance_list),
                "has_notes":        bool(st.session_state.lecture_notes),
            })
            st.balloons()
            st.success("✅ Lecture processed successfully! Check the **Lecture Notes** and **Attendance** tabs.")

    # ── Transcript & Summary display ──
    if st.session_state.transcript:
        st.markdown('<div class="section-header">📜 Raw Transcript</div>', unsafe_allow_html=True)
        with st.expander("Click to view full transcript", expanded=False):
            st.markdown(
                '<div class="notes-preview">' + st.session_state.transcript + '</div>',
                unsafe_allow_html=True,
            )

        if st.session_state.summary:
            st.markdown('<div class="section-header">📄 Lecture Summary</div>', unsafe_allow_html=True)
            st.markdown(
                '<div class="notes-preview">' + st.session_state.summary + '</div>',
                unsafe_allow_html=True,
            )


# ══════════════════════════════════════════════
# TAB 2 — Lecture Notes
# ══════════════════════════════════════════════
with tab2:
    if not st.session_state.lecture_notes:
        st.markdown(
            '<div class="info-banner">ℹ️ Process a lecture in the <strong>Upload & Process</strong> tab to generate notes here.</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown('<div class="section-header">📝 Generated Lecture Notes</div>', unsafe_allow_html=True)
        
        meta = {
            "class_name":   st.session_state.class_name,
            "teacher_name": st.session_state.teacher_name,
            "date":         st.session_state.session_date,
        }

        # Export buttons row
        exp_col1, exp_col2, exp_col3 = st.columns([1, 1, 2])
        with exp_col1:
            try:
                pdf_bytes = export_pdf(st.session_state.lecture_notes, meta)
                st.download_button(
                    "📄 Download PDF",
                    data=pdf_bytes,
                    file_name=f"lecture_notes_{st.session_state.session_date}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"PDF export error: {e}")

        with exp_col2:
            try:
                docx_bytes = export_docx(st.session_state.lecture_notes, meta)
                st.download_button(
                    "📝 Download DOCX",
                    data=docx_bytes,
                    file_name=f"lecture_notes_{st.session_state.session_date}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"DOCX export error: {e}")

        with exp_col3:
            st.markdown(
                f'<div style="padding:0.6rem 0"><span class="tag-pill tag-blue">📅 {st.session_state.session_date}</span> '
                f'<span class="tag-pill tag-purple" style="margin-left:0.4rem">🎓 {st.session_state.class_name or "No class set"}</span></div>',
                unsafe_allow_html=True,
            )

        st.markdown('<div style="height:0.5rem"></div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="notes-preview">' +
            st.session_state.lecture_notes.replace("\n", "<br>") +
            '</div>',
            unsafe_allow_html=True,
        )


# ══════════════════════════════════════════════
# TAB 3 — Attendance
# ══════════════════════════════════════════════
with tab3:
    if not st.session_state.transcript:
        st.markdown(
            '<div class="info-banner">ℹ️ Process a lecture first to see attendance data.</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown('<div class="section-header">👥 Attendance Register</div>', unsafe_allow_html=True)

        att_col1, att_col2 = st.columns([1, 1], gap="large")
        
        with att_col1:
            st.markdown("**Edit / add student names** *(one per line)*")
            raw_names = st.text_area(
                "Student names",
                value="\n".join(st.session_state.attendance_list),
                height=200,
                label_visibility="collapsed",
                key="att_edit",
            )
            st.session_state.attendance_list = [n.strip() for n in raw_names.split("\n") if n.strip()]

        with att_col2:
            total = st.session_state.attendance_list
            count = len(total)
            
            st.markdown(
                f'<div class="metric-box" style="margin-bottom:1rem">'
                f'<div class="metric-val" style="color:#059669">{count}</div>'
                f'<div class="metric-label">Students Present</div>'
                f'</div>',
                unsafe_allow_html=True,
            )
            
            if total:
                st.markdown("**Present Students:**")
                chips = "".join(f'<span class="student-chip">✓ {n}</span>' for n in total)
                st.markdown(f'<div style="line-height:2.2">{chips}</div>', unsafe_allow_html=True)
            else:
                st.markdown(
                    '<div class="warning-banner">⚠️ No students detected. Add names manually.</div>',
                    unsafe_allow_html=True,
                )

        if st.session_state.attendance_list:
            st.markdown('<div style="height:0.8rem"></div>', unsafe_allow_html=True)
            csv_data = "Name,Status,Date,Class\n" + "\n".join(
                f"{n},Present,{st.session_state.session_date},{st.session_state.class_name}"
                for n in st.session_state.attendance_list
            )
            dl_cols = st.columns([1, 2])
            with dl_cols[0]:
                st.download_button(
                    "⬇️ Export Attendance CSV",
                    data=csv_data,
                    file_name=f"attendance_{st.session_state.session_date}.csv",
                    mime="text/csv",
                    use_container_width=True,
                )


# ══════════════════════════════════════════════
# TAB 4 — History
# ══════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-header">📚 Session History</div>', unsafe_allow_html=True)
    
    if not st.session_state.history:
        st.markdown(
            '<div class="info-banner">ℹ️ No sessions processed yet. Start by uploading a lecture in the <strong>Upload & Process</strong> tab.</div>',
            unsafe_allow_html=True,
        )
    else:
        # Summary row
        h_col1, h_col2, h_col3, h_col4 = st.columns(4)
        with h_col1:
            st.markdown(f'<div class="metric-box"><div class="metric-val">{len(st.session_state.history)}</div><div class="metric-label">Total Sessions</div></div>', unsafe_allow_html=True)
        with h_col2:
            total_s = sum(s["students_present"] for s in st.session_state.history)
            st.markdown(f'<div class="metric-box"><div class="metric-val">{total_s}</div><div class="metric-label">Total Students</div></div>', unsafe_allow_html=True)
        with h_col3:
            with_notes = sum(1 for s in st.session_state.history if s["has_notes"])
            st.markdown(f'<div class="metric-box"><div class="metric-val">{with_notes}</div><div class="metric-label">Notes Generated</div></div>', unsafe_allow_html=True)
        with h_col4:
            avg_s = round(total_s / len(st.session_state.history), 1) if st.session_state.history else 0
            st.markdown(f'<div class="metric-box"><div class="metric-val">{avg_s}</div><div class="metric-label">Avg. Attendance</div></div>', unsafe_allow_html=True)

        st.markdown('<div style="height:1rem"></div>', unsafe_allow_html=True)

        for i, session in enumerate(reversed(st.session_state.history)):
            session_num = len(st.session_state.history) - i
            notes_tag   = '<span class="tag-pill tag-green">✓ Notes</span>' if session["has_notes"] else '<span class="tag-pill tag-amber">No Notes</span>'
            
            with st.expander(
                f"#{session_num}  |  {session['date']}  —  {session['class']}  ({session['teacher']})",
                expanded=(i == 0),
            ):
                ec1, ec2, ec3, ec4 = st.columns(4)
                ec1.metric("Date",             session["date"])
                ec2.metric("Students Present", session["students_present"])
                ec3.metric("Class",            session["class"])
                ec4.metric("Teacher",          session["teacher"])
                
                st.markdown(
                    f'<div style="margin-top:0.5rem">{notes_tag} '
                    f'<span class="tag-pill tag-blue" style="margin-left:0.4rem">Session #{session_num}</span></div>',
                    unsafe_allow_html=True,
                )
