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

# ──────────────────────────────────────────────
# PAGE CONFIG
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="SmartAttend – AI Attendance & Notes",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────
# CUSTOM CSS  (dark academia / refined academic)
# ──────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700&family=Source+Sans+3:wght@300;400;600&display=swap');

:root {
    --bg:       #0f1117;
    --surface:  #181c27;
    --card:     #1e2333;
    --border:   #2d3352;
    --accent:   #6c8fff;
    --accent2:  #ff7c5c;
    --text:     #e8eaf2;
    --muted:    #8b90a8;
    --success:  #52c87a;
    --gold:     #e8c97a;
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
    box-shadow: 0 6px 20px rgba(108, 143, 255, 0.35) !important;
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
.student-chip.absent {
    background: #3a1e1e;
    border-color: var(--accent2);
    color: var(--accent2);
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

.tag {
    display: inline-block;
    background: rgba(108,143,255,0.15);
    color: var(--accent);
    border-radius: 5px;
    padding: 0.1rem 0.55rem;
    font-size: 0.78rem;
    font-weight: 600;
    margin-right: 0.4rem;
    letter-spacing: 0.04em;
}

[data-testid="stFileUploaderDropzone"] {
    background: var(--card) !important;
    border: 2px dashed var(--border) !important;
    border-radius: 10px !important;
}
[data-testid="stTextInput"] input,
[data-testid="stTextArea"] textarea,
.stSelectbox select {
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
.stAlert {
    background: var(--card) !important;
    border-radius: 8px !important;
}
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────
# SESSION STATE INIT
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
def get_groq_client(api_key: str):
    return Groq(api_key=api_key)


def transcribe_audio(client, audio_path: str) -> str:
    with open(audio_path, "rb") as f:
        result = client.audio.transcriptions.create(
            model="whisper-large-v3",
            file=f,
            response_format="text",
        )
    return result


def llm_call(client, system: str, user: str) -> str:
    response = client.chat.completions.create(
        model="llama3-8b-8192",
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        temperature=0.4,
        max_tokens=2048,
    )
    return response.choices[0].message.content.strip()


def summarize_transcript(client, transcript: str) -> str:
    system = (
        "You are an expert academic summarizer. "
        "Given a classroom lecture transcript, produce a clear, concise summary "
        "(3-5 paragraphs) covering the main topics, key arguments, and conclusions. "
        "Use formal academic language."
    )
    return llm_call(client, system, f"Transcript:\n\n{transcript}")


def generate_lecture_notes(client, transcript: str, class_name: str) -> str:
    system = (
        "You are a professional academic note-taker. "
        "Convert the following lecture transcript into structured, well-formatted lecture notes. "
        "Use this exact structure:\n"
        "# Lecture Notes – {class_name}\n"
        "## 1. Overview\n"
        "## 2. Key Concepts\n"
        "## 3. Detailed Breakdown\n"
        "   - Sub-topics with bullet points\n"
        "## 4. Important Definitions\n"
        "## 5. Examples & Case Studies (if any)\n"
        "## 6. Summary & Takeaways\n"
        "## 7. Possible Exam Questions\n"
        "Be thorough, organized, and use markdown formatting."
    ).replace("{class_name}", class_name or "Class")
    return llm_call(client, system, f"Transcript:\n\n{transcript}")


def extract_attendance(client, transcript: str) -> list:
    system = (
        "You are an attendance assistant. "
        "Extract the names of all students mentioned as PRESENT in the following transcript. "
        "Return ONLY a JSON array of strings — no explanation, no markdown, no extra text. "
        'Example: ["Alice Smith", "Bob Jones"]'
    )
    raw = llm_call(client, system, f"Transcript:\n\n{transcript}")
    try:
        raw = raw.strip().lstrip("```json").lstrip("```").rstrip("```").strip()
        return json.loads(raw)
    except Exception:
        return []


# ──────────────────────────────────────────────
# PDF EXPORT
# ──────────────────────────────────────────────
def export_pdf(notes: str, meta: dict) -> bytes:
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Header
    pdf.set_font("Helvetica", "B", 20)
    pdf.set_text_color(30, 50, 120)
    pdf.cell(0, 12, "SmartAttend – Lecture Notes", ln=True, align="C")
    pdf.set_font("Helvetica", "", 11)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(0, 7, f"Class: {meta.get('class_name','—')}   |   Teacher: {meta.get('teacher_name','—')}   |   Date: {meta.get('date','—')}", ln=True, align="C")
    pdf.ln(5)
    pdf.set_draw_color(108, 143, 255)
    pdf.set_line_width(0.8)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(6)

    # Body
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
            pdf.multi_cell(0, 7, f"  • {line[2:]}")
        else:
            pdf.multi_cell(0, 7, line)

    return bytes(pdf.output())


# ──────────────────────────────────────────────
# DOCX EXPORT
# ──────────────────────────────────────────────
def export_docx(notes: str, meta: dict) -> bytes:
    doc = Document()

    # Title
    title = doc.add_heading("SmartAttend – Lecture Notes", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.runs[0]
    run.font.color.rgb = RGBColor(30, 50, 120)
    run.font.size = Pt(22)

    # Meta
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(
        f"Class: {meta.get('class_name','—')}   |   "
        f"Teacher: {meta.get('teacher_name','—')}   |   "
        f"Date: {meta.get('date','—')}"
    ).font.color.rgb = RGBColor(100, 100, 100)

    doc.add_paragraph()

    # Body
    for line in notes.split("\n"):
        line_s = line.strip()
        if not line_s:
            doc.add_paragraph()
            continue
        if line_s.startswith("# "):
            h = doc.add_heading(line_s[2:], level=1)
            h.runs[0].font.color.rgb = RGBColor(30, 50, 120)
        elif line_s.startswith("## "):
            h = doc.add_heading(line_s[3:], level=2)
            h.runs[0].font.color.rgb = RGBColor(60, 90, 180)
        elif line_s.startswith("- ") or line_s.startswith("* "):
            p = doc.add_paragraph(line_s[2:], style="List Bullet")
            p.runs[0].font.size = Pt(11)
        else:
            p = doc.add_paragraph(line_s)
            p.runs[0].font.size = Pt(11)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ──────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🎓 SmartAttend")
    st.markdown("<p style='color:#8b90a8;font-size:0.82rem;margin-top:-0.5rem'>AI Classroom Assistant</p>", unsafe_allow_html=True)
    st.divider()

    api_key = st.text_input("🔑 Groq API Key", type="password", placeholder="gsk_...")
    st.divider()

    st.markdown("**Session Details**")
    st.session_state.class_name = st.text_input("Class / Subject", placeholder="e.g. Data Structures CS301")
    st.session_state.teacher_name = st.text_input("Teacher Name", placeholder="e.g. Dr. Ahmed Khan")
    st.session_state.session_date = st.date_input("Date", value=datetime.today()).strftime("%Y-%m-%d")
    st.divider()

    st.markdown("<p style='font-size:0.78rem;color:#8b90a8'>Built with Groq · Whisper · LLaMA3</p>", unsafe_allow_html=True)


# ──────────────────────────────────────────────
# MAIN LAYOUT
# ──────────────────────────────────────────────
st.markdown("""
<div class="hero-card">
  <h1 style="margin:0;font-size:2rem">📋 AI Classroom Attendance & Notes</h1>
  <p style="color:#8b90a8;margin:0.4rem 0 0 0;font-size:0.95rem">
    Upload a lecture recording → get instant transcription, attendance, summary &amp; full lecture notes
  </p>
</div>
""", unsafe_allow_html=True)

# ── TABS ──────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(["🎤 Upload & Process", "📝 Lecture Notes", "👥 Attendance", "📚 History"])

# ══════════════════════════════════════════════
# TAB 1 – UPLOAD & PROCESS
# ══════════════════════════════════════════════
with tab1:
    col_upload, col_text = st.columns([1, 1], gap="large")

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
            "Manual transcript input",
            height=160,
            placeholder="Paste lecture transcript here if you already have one…",
            label_visibility="collapsed",
        )

    with col_text:
        st.markdown('<div class="section-header">Processing Options</div>', unsafe_allow_html=True)
        do_attendance = st.checkbox("Extract attendance from transcript", value=True)
        do_summary    = st.checkbox("Generate lecture summary", value=True)
        do_notes      = st.checkbox("Generate full lecture notes", value=True)
        do_pdf        = st.checkbox("Auto-generate PDF notes", value=True)
        do_docx       = st.checkbox("Auto-generate DOCX notes", value=True)

        st.markdown("")
        process_btn = st.button("⚡ Process Lecture", use_container_width=True)

    # ── PROCESS ───────────────────────────────
    if process_btn:
        if not api_key:
            st.error("Please enter your Groq API key in the sidebar.")
        elif not audio_file and not manual_text.strip():
            st.warning("Please upload an audio file or paste a transcript.")
        else:
            client = get_groq_client(api_key)

            # Step 1 – Transcription
            if audio_file and not manual_text.strip():
                with st.spinner("🎙️ Transcribing audio with Whisper…"):
                    with tempfile.NamedTemporaryFile(delete=False, suffix=f".{audio_file.name.split('.')[-1]}") as tmp:
                        tmp.write(audio_file.read())
                        tmp_path = tmp.name
                    st.session_state.transcript = transcribe_audio(client, tmp_path)
                    os.unlink(tmp_path)
            else:
                st.session_state.transcript = manual_text.strip()

            st.success("✅ Transcript ready")

            # Step 2 – Attendance
            if do_attendance:
                with st.spinner("👥 Extracting attendance…"):
                    st.session_state.attendance_list = extract_attendance(client, st.session_state.transcript)

            # Step 3 – Summary
            if do_summary:
                with st.spinner("📄 Summarizing lecture…"):
                    st.session_state.summary = summarize_transcript(client, st.session_state.transcript)

            # Step 4 – Notes
            if do_notes:
                with st.spinner("📝 Generating lecture notes…"):
                    st.session_state.lecture_notes = generate_lecture_notes(
                        client, st.session_state.transcript, st.session_state.class_name
                    )

            # Save to history
            st.session_state.history.append({
                "date": st.session_state.session_date,
                "class": st.session_state.class_name or "Unnamed Class",
                "teacher": st.session_state.teacher_name or "Unknown",
                "students_present": len(st.session_state.attendance_list),
                "has_notes": bool(st.session_state.lecture_notes),
            })

            st.balloons()

    # ── TRANSCRIPT PREVIEW ────────────────────
    if st.session_state.transcript:
        with st.expander("📜 View Raw Transcript", expanded=False):
            st.markdown(f'<div class="notes-preview">{st.session_state.transcript}</div>', unsafe_allow_html=True)

        if st.session_state.summary:
            st.markdown('<div class="section-header">📄 Lecture Summary</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="notes-preview">{st.session_state.summary}</div>', unsafe_allow_html=True)

        # Metrics row
        st.markdown("")
        m1, m2, m3 = st.columns(3)
        word_count = len(st.session_state.transcript.split())
        with m1:
            st.markdown(f'<div class="metric-box"><div class="metric-val">{word_count:,}</div><div class="metric-label">Words Transcribed</div></div>', unsafe_allow_html=True)
        with m2:
            st.markdown(f'<div class="metric-box"><div class="metric-val">{len(st.session_state.attendance_list)}</div><div class="metric-label">Students Present</div></div>', unsafe_allow_html=True)
        with m3:
            notes_lines = len(st.session_state.lecture_notes.split("\n")) if st.session_state.lecture_notes else 0
            st.markdown(f'<div class="metric-box"><div class="metric-val">{notes_lines}</div><div class="metric-label">Notes Lines Generated</div></div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════
# TAB 2 – LECTURE NOTES
# ══════════════════════════════════════════════
with tab2:
    if not st.session_state.lecture_notes:
        st.info("Process a lecture in the **Upload & Process** tab to generate notes here.")
    else:
        st.markdown('<div class="section-header">📝 Generated Lecture Notes</div>', unsafe_allow_html=True)

        meta = {
            "class_name": st.session_state.class_name,
            "teacher_name": st.session_state.teacher_name,
            "date": st.session_state.session_date,
        }

        col_dl1, col_dl2, col_dl3 = st.columns([1, 1, 2])

        with col_dl1:
            pdf_bytes = export_pdf(st.session_state.lecture_notes, meta)
            st.download_button(
                "⬇️ Download PDF",
                data=pdf_bytes,
                file_name=f"lecture_notes_{st.session_state.session_date}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        with col_dl2:
            docx_bytes = export_docx(st.session_state.lecture_notes, meta)
            st.download_button(
                "⬇️ Download DOCX",
                data=docx_bytes,
                file_name=f"lecture_notes_{st.session_state.session_date}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        st.markdown('<div class="notes-preview">' + st.session_state.lecture_notes.replace("\n", "<br>") + '</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════
# TAB 3 – ATTENDANCE
# ══════════════════════════════════════════════
with tab3:
    if not st.session_state.transcript:
        st.info("Process a lecture first to see attendance.")
    else:
        st.markdown('<div class="section-header">👥 Attendance Register</div>', unsafe_allow_html=True)

        # Manual override
        raw_names = st.text_area(
            "Edit / add student names (one per line)",
            value="\n".join(st.session_state.attendance_list),
            height=140,
        )
        st.session_state.attendance_list = [n.strip() for n in raw_names.split("\n") if n.strip()]

        total = st.session_state.attendance_list
        if total:
            st.markdown(f"**{len(total)} students marked present**")
            chips = "".join(f'<span class="student-chip">{n}</span>' for n in total)
            st.markdown(chips, unsafe_allow_html=True)
        else:
            st.warning("No students detected. Try editing the list above.")

        # Export attendance CSV
        if total:
            csv_lines = ["Name,Status,Date,Class"]
            for name in total:
                csv_lines.append(f"{name},Present,{st.session_state.session_date},{st.session_state.class_name}")
            csv_data = "\n".join(csv_lines)
            st.download_button(
                "⬇️ Export Attendance CSV",
                data=csv_data,
                file_name=f"attendance_{st.session_state.session_date}.csv",
                mime="text/csv",
            )


# ══════════════════════════════════════════════
# TAB 4 – HISTORY
# ══════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-header">📚 Session History</div>', unsafe_allow_html=True)
    if not st.session_state.history:
        st.info("No sessions processed yet.")
    else:
        for i, session in enumerate(reversed(st.session_state.history)):
            with st.expander(f"📅 {session['date']} — {session['class']} ({session['teacher']})"):
                c1, c2, c3 = st.columns(3)
                c1.metric("Students Present", session["students_present"])
                c2.metric("Notes Generated", "✅" if session["has_notes"] else "❌")
                c3.metric("Session #", len(st.session_state.history) - i)
