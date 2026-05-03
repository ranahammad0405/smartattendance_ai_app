import streamlit as st
import os
import tempfile
import json
from datetime import datetime
from groq import Groq
from fpdf import FPDF
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

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
# CUSTOM CSS
# ──────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700&family=Source+Sans+3:wght@300;400;600&display=swap');

:root {
    --bg:      #0f1117;
    --surface: #181c27;
    --card:    #1e2333;
    --border:  #2d3352;
    --accent:  #6c8fff;
    --accent2: #ff7c5c;
    --text:    #e8eaf2;
    --muted:   #8b90a8;
    --success: #52c87a;
    --gold:    #e8c97a;
}
html, body, [data-testid="stAppViewContainer"] {
    background-color: var(--bg) !important;
    font-family: 'Source Sans 3', sans-serif;
    color: var(--text);
}
[data-testid="stSidebar"] {
    background-color: var(--surface) !important;
    border-right: 1px solid var(--border);
}
h1, h2, h3, h4 {
    font-family: 'Playfair Display', serif !important;
    color: var(--text) !important;
}
.stButton > button {
    background: linear-gradient(135deg, var(--accent), #4a6bdf) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 0.6rem 1.4rem !important;
    font-family: 'Source Sans 3', sans-serif !important;
    font-weight: 600 !important;
    letter-spacing: 0.03em !important;
    transition: all 0.2s ease !important;
}
.stButton > button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 20px rgba(108,143,255,0.35) !important;
}
.hero-card {
    background: linear-gradient(135deg, #1a2040 0%, #0f1428 100%);
    border: 1px solid var(--border);
    border-left: 4px solid var(--accent);
    border-radius: 12px;
    padding: 1.6rem 2rem;
    margin-bottom: 1.5rem;
}
.metric-box {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 1.2rem;
    text-align: center;
}
.metric-box .metric-val {
    font-family: 'Playfair Display', serif;
    font-size: 2.4rem;
    font-weight: 700;
    color: var(--accent);
    line-height: 1;
}
.metric-box .metric-label {
    font-size: 0.8rem;
    color: var(--muted);
    text-transform: uppercase;
    letter-spacing: 0.08em;
    margin-top: 0.3rem;
}
.student-chip {
    display: inline-block;
    background: #2a3a1e;
    border: 1px solid var(--success);
    color: var(--success);
    border-radius: 20px;
    padding: 0.2rem 0.85rem;
    font-size: 0.82rem;
    margin: 0.2rem;
}
.notes-preview {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 1.5rem;
    white-space: pre-wrap;
    font-family: 'Source Sans 3', sans-serif;
    font-size: 0.92rem;
    line-height: 1.7;
    max-height: 420px;
    overflow-y: auto;
    color: var(--text);
}
.section-header {
    font-family: 'Playfair Display', serif;
    font-size: 1.3rem;
    font-weight: 600;
    color: var(--gold);
    margin: 1.5rem 0 0.8rem 0;
    padding-bottom: 0.4rem;
    border-bottom: 1px solid var(--border);
}
.secret-badge {
    background: #1e2b1e;
    border: 1px solid #52c87a;
    border-radius: 8px;
    padding: 0.6rem 0.9rem;
    font-size: 0.82rem;
    color: #52c87a;
    margin-bottom: 1rem;
}
[data-testid="stFileUploaderDropzone"] {
    background: var(--card) !important;
    border: 2px dashed var(--border) !important;
    border-radius: 10px !important;
}
[data-testid="stTextInput"] input,
[data-testid="stTextArea"] textarea {
    background: var(--card) !important;
    color: var(--text) !important;
    border-color: var(--border) !important;
    border-radius: 8px !important;
}
[data-testid="stExpander"] {
    background: var(--card) !important;
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
}
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# API KEY - loaded from Streamlit Secrets ONLY
# No key input in the UI - keeps your key hidden
# ──────────────────────────────────────────────
api_key = st.secrets.get("GROQ_API_KEY", "")
if not api_key:
    st.error(
        "GROQ_API_KEY not found. "
        "In Streamlit Cloud go to: App Settings -> Secrets and add:\n\n"
        "GROQ_API_KEY = \"gsk_your_key_here\""
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
# llama3-8b-8192 has ~8192 token context.
# 1800 words ~ 2400 tokens, leaves headroom for
# system prompt + output tokens.
# ──────────────────────────────────────────────
def safe_truncate(text: str, max_words: int = 1800) -> str:
    words = text.split()
    if len(words) <= max_words:
        return text
    return " ".join(words[:max_words]) + "\n\n[Transcript truncated to fit model context window]"


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
        "Use formal academic language."
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
        "Be thorough, organized, and use markdown formatting."
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
# PDF EXPORT
# ──────────────────────────────────────────────
def export_pdf(notes: str, meta: dict) -> bytes:
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 20)
    pdf.set_text_color(30, 50, 120)
    pdf.cell(0, 12, "SmartAttend - Lecture Notes", ln=True, align="C")
    pdf.set_font("Helvetica", "", 11)
    pdf.set_text_color(80, 80, 80)
    info = (
        "Class: " + meta.get("class_name", "-") +
        "   |   Teacher: " + meta.get("teacher_name", "-") +
        "   |   Date: " + meta.get("date", "-")
    )
    pdf.cell(0, 7, info, ln=True, align="C")
    pdf.ln(5)
    pdf.set_draw_color(108, 143, 255)
    pdf.set_line_width(0.8)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(6)
    pdf.set_font("Helvetica", "", 11)
    pdf.set_text_color(30, 30, 30)
    for line in notes.split("\n"):
        line = line.strip()
        if not line:
            pdf.ln(3)
            continue
        if line.startswith("# "):
            pdf.set_font("Helvetica", "B", 15)
            pdf.set_text_color(30, 50, 120)
            pdf.multi_cell(0, 9, line[2:])
            pdf.set_font("Helvetica", "", 11)
            pdf.set_text_color(30, 30, 30)
        elif line.startswith("## "):
            pdf.set_font("Helvetica", "B", 12)
            pdf.set_text_color(60, 90, 180)
            pdf.multi_cell(0, 8, line[3:])
            pdf.set_font("Helvetica", "", 11)
            pdf.set_text_color(30, 30, 30)
        elif line.startswith("- ") or line.startswith("* "):
            pdf.multi_cell(0, 7, "  * " + line[2:])
        else:
            pdf.multi_cell(0, 7, line)
    return bytes(pdf.output())


# ──────────────────────────────────────────────
# DOCX EXPORT
# ──────────────────────────────────────────────
def export_docx(notes: str, meta: dict) -> bytes:
    doc = Document()
    title = doc.add_heading("SmartAttend - Lecture Notes", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if title.runs:
        title.runs[0].font.color.rgb = RGBColor(30, 50, 120)
        title.runs[0].font.size = Pt(22)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info = (
        "Class: " + meta.get("class_name", "-") +
        "   |   Teacher: " + meta.get("teacher_name", "-") +
        "   |   Date: " + meta.get("date", "-")
    )
    run = p.add_run(info)
    run.font.color.rgb = RGBColor(100, 100, 100)
    doc.add_paragraph()
    for line in notes.split("\n"):
        ls = line.strip()
        if not ls:
            doc.add_paragraph()
            continue
        if ls.startswith("# "):
            h = doc.add_heading(ls[2:], level=1)
            if h.runs:
                h.runs[0].font.color.rgb = RGBColor(30, 50, 120)
        elif ls.startswith("## "):
            h = doc.add_heading(ls[3:], level=2)
            if h.runs:
                h.runs[0].font.color.rgb = RGBColor(60, 90, 180)
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
    st.markdown("## 🎓 SmartAttend")
    st.markdown(
        "<p style='color:#8b90a8;font-size:0.82rem;margin-top:-0.5rem'>AI Classroom Assistant</p>",
        unsafe_allow_html=True,
    )
    st.divider()
    st.markdown(
        "<div class='secret-badge'>🔒 API key loaded securely from Streamlit Secrets</div>",
        unsafe_allow_html=True,
    )
    st.markdown("**Session Details**")
    st.session_state.class_name    = st.text_input("Class / Subject", placeholder="e.g. Data Structures CS301")
    st.session_state.teacher_name  = st.text_input("Teacher Name",    placeholder="e.g. Dr. Ahmed Khan")
    st.session_state.session_date  = st.date_input("Date", value=datetime.today()).strftime("%Y-%m-%d")
    st.divider()
    st.markdown(
        "<p style='font-size:0.78rem;color:#8b90a8'>Built with Groq · Whisper · LLaMA3</p>",
        unsafe_allow_html=True,
    )


# ──────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────
st.markdown("""
<div class="hero-card">
  <h1 style="margin:0;font-size:2rem">📋 AI Classroom Attendance &amp; Notes</h1>
  <p style="color:#8b90a8;margin:0.4rem 0 0 0;font-size:0.95rem">
    Upload a lecture recording &rarr; instant transcription, attendance, summary &amp; full lecture notes
  </p>
</div>
""", unsafe_allow_html=True)

tab1, tab2, tab3, tab4 = st.tabs([
    "🎤 Upload & Process",
    "📝 Lecture Notes",
    "👥 Attendance",
    "📚 History",
])

# ══════════ TAB 1 ══════════
with tab1:
    col_upload, col_opts = st.columns([1, 1], gap="large")

    with col_upload:
        st.markdown('<div class="section-header">Upload Audio / Video</div>', unsafe_allow_html=True)
        audio_file = st.file_uploader(
            "Drop lecture recording here",
            type=["mp3", "mp4", "wav", "m4a", "ogg", "webm"],
            label_visibility="collapsed",
        )
        if audio_file:
            st.audio(audio_file)

        st.markdown('<div class="section-header">Or Paste Transcript</div>', unsafe_allow_html=True)
        manual_text = st.text_area(
            "Manual transcript",
            height=160,
            placeholder="Paste lecture transcript here if you already have one...",
            label_visibility="collapsed",
        )

    with col_opts:
        st.markdown('<div class="section-header">Processing Options</div>', unsafe_allow_html=True)
        do_attendance = st.checkbox("Extract attendance from transcript", value=True)
        do_summary    = st.checkbox("Generate lecture summary",           value=True)
        do_notes      = st.checkbox("Generate full lecture notes",        value=True)
        st.markdown("")
        process_btn = st.button("⚡ Process Lecture", use_container_width=True)

    if process_btn:
        if not audio_file and not manual_text.strip():
            st.warning("Please upload an audio file or paste a transcript.")
        else:
            client = get_groq_client(api_key)

            # Transcription
            if audio_file and not manual_text.strip():
                with st.spinner("🎙️ Transcribing audio with Whisper..."):
                    suffix = "." + audio_file.name.split(".")[-1]
                    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                        tmp.write(audio_file.read())
                        tmp_path = tmp.name
                    try:
                        st.session_state.transcript = transcribe_audio(client, tmp_path)
                    finally:
                        os.unlink(tmp_path)
            else:
                st.session_state.transcript = manual_text.strip()

            st.success("Transcript ready")

            # Attendance
            if do_attendance:
                with st.spinner("👥 Extracting attendance..."):
                    try:
                        st.session_state.attendance_list = extract_attendance(
                            client, st.session_state.transcript
                        )
                    except Exception as e:
                        st.warning(f"Attendance extraction failed ({e}). Add names manually in the Attendance tab.")
                        st.session_state.attendance_list = []

            # Summary
            if do_summary:
                with st.spinner("📄 Summarizing lecture..."):
                    try:
                        st.session_state.summary = summarize_transcript(
                            client, st.session_state.transcript
                        )
                    except Exception as e:
                        st.warning(f"Summary failed: {e}")

            # Notes
            if do_notes:
                with st.spinner("📝 Generating lecture notes..."):
                    try:
                        st.session_state.lecture_notes = generate_lecture_notes(
                            client, st.session_state.transcript, st.session_state.class_name
                        )
                    except Exception as e:
                        st.warning(f"Notes generation failed: {e}")

            st.session_state.history.append({
                "date":             st.session_state.session_date,
                "class":            st.session_state.class_name or "Unnamed Class",
                "teacher":          st.session_state.teacher_name or "Unknown",
                "students_present": len(st.session_state.attendance_list),
                "has_notes":        bool(st.session_state.lecture_notes),
            })
            st.balloons()

    if st.session_state.transcript:
        with st.expander("📜 View Raw Transcript", expanded=False):
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
        st.markdown("")
        m1, m2, m3 = st.columns(3)
        wc = len(st.session_state.transcript.split())
        nl = len(st.session_state.lecture_notes.split("\n")) if st.session_state.lecture_notes else 0
        with m1:
            st.markdown(f'<div class="metric-box"><div class="metric-val">{wc:,}</div><div class="metric-label">Words Transcribed</div></div>', unsafe_allow_html=True)
        with m2:
            st.markdown(f'<div class="metric-box"><div class="metric-val">{len(st.session_state.attendance_list)}</div><div class="metric-label">Students Present</div></div>', unsafe_allow_html=True)
        with m3:
            st.markdown(f'<div class="metric-box"><div class="metric-val">{nl}</div><div class="metric-label">Notes Lines</div></div>', unsafe_allow_html=True)

# ══════════ TAB 2 ══════════
with tab2:
    if not st.session_state.lecture_notes:
        st.info("Process a lecture in the Upload & Process tab to generate notes here.")
    else:
        st.markdown('<div class="section-header">📝 Generated Lecture Notes</div>', unsafe_allow_html=True)
        meta = {
            "class_name":   st.session_state.class_name,
            "teacher_name": st.session_state.teacher_name,
            "date":         st.session_state.session_date,
        }
        dl1, dl2, _ = st.columns([1, 1, 2])
        with dl1:
            st.download_button(
                "⬇️ Download PDF",
                data=export_pdf(st.session_state.lecture_notes, meta),
                file_name=f"lecture_notes_{st.session_state.session_date}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        with dl2:
            st.download_button(
                "⬇️ Download DOCX",
                data=export_docx(st.session_state.lecture_notes, meta),
                file_name=f"lecture_notes_{st.session_state.session_date}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        st.markdown(
            '<div class="notes-preview">' +
            st.session_state.lecture_notes.replace("\n", "<br>") +
            '</div>',
            unsafe_allow_html=True,
        )

# ══════════ TAB 3 ══════════
with tab3:
    if not st.session_state.transcript:
        st.info("Process a lecture first to see attendance.")
    else:
        st.markdown('<div class="section-header">👥 Attendance Register</div>', unsafe_allow_html=True)
        raw_names = st.text_area(
            "Edit / add student names (one per line)",
            value="\n".join(st.session_state.attendance_list),
            height=140,
        )
        st.session_state.attendance_list = [n.strip() for n in raw_names.split("\n") if n.strip()]
        total = st.session_state.attendance_list
        if total:
            st.markdown(f"**{len(total)} students marked present**")
            st.markdown(
                "".join(f'<span class="student-chip">{n}</span>' for n in total),
                unsafe_allow_html=True,
            )
            csv_data = "Name,Status,Date,Class\n" + "\n".join(
                f"{n},Present,{st.session_state.session_date},{st.session_state.class_name}"
                for n in total
            )
            st.download_button(
                "⬇️ Export Attendance CSV",
                data=csv_data,
                file_name=f"attendance_{st.session_state.session_date}.csv",
                mime="text/csv",
            )
        else:
            st.warning("No students detected. Add names manually above.")

# ══════════ TAB 4 ══════════
with tab4:
    st.markdown('<div class="section-header">📚 Session History</div>', unsafe_allow_html=True)
    if not st.session_state.history:
        st.info("No sessions processed yet.")
    else:
        for i, session in enumerate(reversed(st.session_state.history)):
            with st.expander(f"📅 {session['date']} — {session['class']} ({session['teacher']})"):
                c1, c2, c3 = st.columns(3)
                c1.metric("Students Present", session["students_present"])
                c2.metric("Notes Generated",  "Yes" if session["has_notes"] else "No")
                c3.metric("Session #",         len(st.session_state.history) - i)
