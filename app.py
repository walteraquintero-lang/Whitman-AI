import streamlit as st
import google.generativeai as genai
import io
import os
import re
from docx import Document
from pptx import Presentation
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Whitman AI - Liceo Moderno Walt Whitman",
    page_icon="🎓",
    layout="wide"
)

# --- ESTILOS CSS ---
st.markdown("""
<style>
    .header-box {
        background: linear-gradient(135deg, #0f2042 0%, #1e3a8a 100%);
        border-left: 6px solid #f59e0b;
        border-radius: 12px;
        padding: 16px 22px;
        margin-bottom: 20px;
        color: white;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25);
    }
    .header-box h1 { margin: 0; font-size: 1.8rem; color: #ffffff; font-weight: 800; }
    .header-box p { margin: 3px 0 0 0; color: #93c5fd; font-size: 0.9rem; }
    .badge-role {
        display: inline-block;
        background-color: #f59e0b;
        color: #0f2042;
        font-weight: 800;
        padding: 3px 12px;
        border-radius: 15px;
        font-size: 0.75rem;
        margin-top: 6px;
        text-transform: uppercase;
    }
    .stDownloadButton > button {
        width: 100%;
        background-color: #0f2042;
        border: 1px solid #f59e0b;
        color: #ffffff;
        font-weight: 600;
        border-radius: 6px;
        padding: 6px;
    }
    .stDownloadButton > button:hover {
        background-color: #f59e0b;
        color: #0f2042;
    }
    .docs-btn {
        display: inline-block;
        width: 100%;
        text-align: center;
        background-color: #1e3a8a;
        color: #ffffff !important;
        border: 1px solid #60a5fa;
        border-radius: 6px;
        padding: 6px;
        font-weight: 600;
        text-decoration: none;
    }
    .docs-btn:hover { background-color: #2563eb; }
</style>
""", unsafe_allow_html=True)

# --- CONEXIÓN API ---
api_key = st.secrets.get("GEMINI_API_KEY")
if not api_key:
    st.error("Falta configurar GEMINI_API_KEY en Secrets de Streamlit.")
    st.stop()

genai.configure(api_key=api_key)

# Lista completa de modelos disponibles en Google
ALL_CANDIDATE_MODELS = [
    "models/gemini-3.8-flash",
    "gemini-3.8-flash",
    "models/gemini-2.0-flash",
    "gemini-2.0-flash",
    "models/gemini-1.5-pro",
    "gemini-1.5-pro",
    "models/gemini-1.5-flash-8b",
    "gemini-1.5-flash-8b"
]

# --- FILTRO DE METADATOS ---
def guillotine_clean(text):
    if not text:
        return ""
    lines = text.split("\n")
    start_index = 0
    for i, line in enumerate(lines):
        clean_l = line.strip().lower()
        if any(clean_l.startswith(k) for k in ["* *identity:", "* *training:", "* *context:", "* *user request:", "* *audience:", "* *topic:", "* *structure:"]):
            start_index = i + 1
        elif clean_l.startswith("#") or clean_l.startswith("**área") or clean_l.startswith("**título") or "guía de" in clean_l or "circular" in clean_l or "el liceo" in clean_l:
            start_index = i
            break
    clean_text = "\n".join(lines[start_index:]).strip()
    return clean_text if clean_text else text.strip()

# --- GENERADORES EN CACHÉ ---
@st.cache_data(show_spinner=False)
def generate_pdf_reportlab(text):
    try:
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle('Title', fontName='Helvetica-Bold', fontSize=12, textColor=colors.HexColor("#0F2042"), alignment=1, spaceAfter=2)
        sub_style = ParagraphStyle('Sub', fontName='Helvetica-Bold', fontSize=8, textColor=colors.HexColor("#F59E0B"), alignment=1, spaceAfter=8)
        h1_style = ParagraphStyle('H1', fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor("#0F2042"), spaceBefore=8, spaceAfter=4)
        body_style = ParagraphStyle('Body', fontName='Helvetica', fontSize=8.5, leading=11, textColor=colors.HexColor("#1E293B"))
        th_style = ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.white, alignment=1)
        td_style = ParagraphStyle('TD', fontName='Helvetica', fontSize=7, leading=8.5, textColor=colors.HexColor("#0F172A"))
        
        story = []
        logo_left = "logo_liceo.png" if os.path.exists("logo_liceo.png") else None
        logo_right = "logo_studio.png" if os.path.exists("logo_studio.png") else None
        
        header_text = [
            Paragraph("LICEO MODERNO WALT WHITMAN", title_style),
            Paragraph("ECOSISTEMA VITAL 26 • WHITMAN STUDIO", sub_style)
        ]
        header_data = [[
            RLImage(logo_left, width=50, height=50) if logo_left else "",
            header_text,
            RLImage(logo_right, width=50, height=50) if logo_right else ""
        ]]
        
        header_table = Table(header_data, colWidths=[60, 420, 60])
        header_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (0,0), (0,0), 'LEFT'),
            ('ALIGN', (1,0), (1,0), 'CENTER'),
            ('ALIGN', (2,0), (2,0), 'RIGHT'),
            ('LINEBELOW', (0,0), (-1,-1), 1.5, colors.HexColor("#F59E0B")),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6)
        ]))
        story.append(header_table)
        story.append(Spacer(1, 8))
        
        clean_text = guillotine_clean(text)
        clean_text = re.sub(r'[^\x00-\x7F\xC0-\xFF]+', ' ', clean_text)
        lines = clean_text.split("\n")
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if line.startswith("|") and line.endswith("|"):
                table_raw = []
                while i < len(lines) and lines[i].strip().startswith("|") and lines[i].strip().endswith("|"):
                    row_str = lines[i].strip()
                    if not re.match(r"^\|(\s*[-:]+\s*\|)+$", row_str):
                        table_raw.append([c.strip() for c in row_str.split("|")[1:-1]])
                    i += 1
                if table_raw:
                    num_cols = max(len(r) for r in table_raw)
                    col_w = 540 / max(num_cols, 1)
                    table_data = []
                    for r_idx, r in enumerate(table_raw):
                        row_cells = [Paragraph(c[:180], th_style if r_idx == 0 else td_style) for c in r]
                        while len(row_cells) < num_cols:
                            row_cells.append(Paragraph("", td_style))
                        table_data.append(row_cells)
                    t = Table(table_data, colWidths=[col_w] * num_cols)
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0F2042")),
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                        ('TOPPADDING', (0, 0), (-1, -1), 3),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                    ]))
                    story.append(t)
                    story.append(Spacer(1, 4))
                continue
            if line.startswith("#"):
                story.append(Spacer(1, 4))
                story.append(Paragraph(re.sub(r"^#+\s*", "", line), h1_style))
            elif line:
                story.append(Paragraph(line, body_style))
                story.append(Spacer(1, 2))
            i += 1
        doc.build(story)
        return buffer.getvalue()
    except Exception:
        return None

@st.cache_data(show_spinner=False)
def generate_docx(text):
    try:
        clean = guillotine_clean(text)
        doc = Document()
        doc.add_heading("Liceo Moderno Walt Whitman", level=0)
        doc.add_paragraph("Ecosistema Vital 26 • Whitman Studio")
        doc.add_paragraph("--------------------------------------------------")
        for line in clean.split("\n"):
            l = line.strip()
            if l.startswith("### "):
                doc.add_heading(l.replace("### ", ""), level=2)
            elif l.startswith("## ") or l.startswith("# "):
                doc.add_heading(l.replace("#", "").strip(), level=1)
            elif l.startswith("- ") or l.startswith("* "):
                doc.add_paragraph(l[2:], style='List Bullet')
            elif l and not l.startswith("|"):
                doc.add_paragraph(l)
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    except Exception:
        return None

@st.cache_data(show_spinner=False)
def generate_pptx(text):
    try:
        clean = guillotine_clean(text)
        prs = Presentation()
        title_slide = prs.slides.add_slide(prs.slide_layouts[0])
        title_slide.shapes.title.text = "Whitman AI - Presentación"
        title_slide.placeholders[1].text = "Liceo Moderno Walt Whitman\nEcosistema Vital 26"
        sections = clean.split("##")
        for sec in sections[1:7]:
            lines = [l.strip() for l in sec.split("\n") if l.strip() and not l.startswith("|")]
            if not lines:
                continue
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = lines[0].replace("#", "").strip()[:50]
            tf = slide.placeholders[1].text_frame
            tf.word_wrap = True
            for body_line in lines[1:5]:
                p = tf.add_paragraph()
                p.text = body_line[:120]
        bio = io.BytesIO()
        prs.save(bio)
        return bio.getvalue()
    except Exception:
        return None

# --- BIENVENIDAS DIRECTAS ---
PROFILES = {
    "wq26": {
        "badge": "ADMINISTRACIÓN / INNOVACIÓN",
        "welcome": "👋 ¡Hola Walter! Centro de Control activo. ¿En qué planeación, circular o proyecto de Whitman Studio trabajamos hoy?"
    },
    "docente": {
        "badge": "DOCENTE LICEÍSTA",
        "welcome": """👋 ¡Bienvenido/a profe! Tu entorno de planeación está listo.
👉 Fórmula rápida: `planear/[grado]/[área]/[tema]/[sesiones]/`
¿Qué clase preparamos hoy?"""
    },
    "sena": {
        "badge": "APRENDIZ SENA",
        "welcome": "🎨 ¡Hola Aprendiz! Laboratorio de Illustrator listo. ¿Revisamos tu Buyer Persona, Reto 5 o una guía técnica?"
    },
    "padre": {
        "badge": "PADRES DE FAMILIA",
        "welcome": "🏡 ¡Bienvenidos Padres y Madres Liceístas! Espacio escolar. ¿En qué información o evento les colaboro?"
    }
}

# Verificación estricta (solo si es únicamente un login de 1 o 2 palabras)
def check_identity(text):
    clean = text.lower().strip()
    if len(clean.split()) > 3:
        return None
    if clean in ["wq26", "soy wq26"]:
        return PROFILES["wq26"]
    elif clean in ["sena", "sena1", "sena2", "soy sena", "soy sena 2"]:
        return PROFILES["sena"]
    elif clean in ["profe", "docente", "profesor", "profesora"]:
        return PROFILES["docente"]
    elif clean in ["padre", "mama", "papa", "acudiente", "soy padre", "familia"]:
        return PROFILES["padre"]
    return None

# --- ESTADOS DE SESIÓN ---
if "messages" not in st.session_state:
    st.session_state.messages = []
if "user_role" not in st.session_state:
    st.session_state.user_role = "DOCENTE LICEÍSTA"

# --- CABECERA PRINCIPAL ---
st.markdown(f"""
<div class="header-box">
    <h1>🎓 Whitman AI • Ecosistema Vital 26</h1>
    <p>Liceo Moderno Walt Whitman & Whitman Studio | Centro Pedagógico e Institucional</p>
    <div class="badge-role">{st.session_state.user_role}</div>
</div>
""", unsafe_allow_html=True)

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("🏫 Acceso Rápido")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("👑 Walter Q."):
            st.session_state.user_role = PROFILES["wq26"]["badge"]
            st.session_state.messages.append({"role": "assistant", "content": PROFILES["wq26"]["welcome"]})
            st.rerun()
        if st.button("🚀 SENA"):
            st.session_state.user_role = PROFILES["sena"]["badge"]
            st.session_state.messages.append({"role": "assistant", "content": PROFILES["sena"]["welcome"]})
            st.rerun()
    with c2:
        if st.button("👩‍🏫 Docentes"):
            st.session_state.user_role = PROFILES["docente"]["badge"]
            st.session_state.messages.append({"role": "assistant", "content": PROFILES["docente"]["welcome"]})
            st.rerun()
        if st.button("🏡 Familias"):
            st.session_state.user_role = PROFILES["padre"]["badge"]
            st.session_state.messages.append({"role": "assistant", "content": PROFILES["padre"]["welcome"]})
            st.rerun()
    st.divider()
    if st.button("🗑️ Nueva Consulta"):
        st.session_state.messages = []
        st.rerun()

# --- HISTORIAL ---
for idx, msg in enumerate(st.session_state.messages):
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])
        
        if msg["role"] == "assistant" and len(msg["content"]) > 90:
            st.markdown("""
            <div style="margin-top: 15px; border-top: 1px solid #1e3a8a; padding-top: 8px;">
                <b style="color: #93c5fd; font-size: 0.85rem;">📥 Exportar Material Oficial:</b>
            </div>
            """, unsafe_allow_html=True)
            
            b1, b2, b3, b4 = st.columns(4)
            docx_data = generate_docx(msg["content"])
            if docx_data:
                with b1:
                    st.download_button("📄 Word (.docx)", docx_data, f"Whitman_{idx}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"docx_{idx}")
            pdf_data = generate_pdf_reportlab(msg["content"])
            if pdf_data:
                with b2:
                    st.download_button("📑 PDF Oficial", pdf_data, f"Whitman_{idx}.pdf", "application/pdf", key=f"pdf_{idx}")
            pptx_data = generate_pptx(msg["content"])
            if pptx_data:
                with b3:
                    st.download_button("📊 Diapositivas (.pptx)", pptx_data, f"Whitman_{idx}.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", key=f"pptx_{idx}")
            with b4:
                st.markdown('<a href="https://docs.new" target="_blank" class="docs-btn">🌐 Abrir en Google Docs</a>', unsafe_allow_html=True)

# --- ENTRADA ---
if prompt := st.chat_input("Escribe tu solicitud de planeación o consulta..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    ident = check_identity(prompt)
    if ident:
        st.session_state.user_role = ident["badge"]
        w = ident["welcome"]
        with st.chat_message("assistant"):
            st.markdown(w)
        st.session_state.messages.append({"role": "assistant", "content": w})
        st.rerun()
    else:
        with st.chat_message("assistant"):
            master_query = f"""
Eres Whitman AI, asistente pedagógico e institucional oficial del Liceo Moderno Walt Whitman (Bogotá, Villa Mayor) y Whitman Studio.
Entrenado por Walter Quintero (Coord. Innovación y Tecnología).
Rol del usuario: {st.session_state.user_role}

REGLAS:
1. Comienza DIRECTAMENTE con la respuesta o planeación. Cero notas previas en inglés.
2. Si es una familia/padre: sé cálido, orientativo e institucional (explica beneficios, enfoque y valores del Liceo).
3. Si es planeación: usa la estructura de DBA, 40-45 min, páginas del libro y PIAR.
4. Finaliza siempre con el lema institucional de Walt Whitman.

SOLICITUD:
{prompt}
"""
            # INTENTO MULTI-MODELO MOSTRANDO EL ESTADO REAL
            answered = False
            last_error = ""
            
            for candidate in ALL_CANDIDATE_MODELS:
                try:
                    m = genai.GenerativeModel(candidate, generation_config={"temperature": 0.2})
                    response = m.generate_content(master_query, stream=True)
                    
                    def stream_chunks():
                        for chunk in response:
                            if chunk.text:
                                yield chunk.text

                    raw_answer = st.write_stream(stream_chunks())
                    clean_answer = guillotine_clean(raw_answer)
                    
                    st.session_state.messages.append({
                        "role": "assistant",
                        "content": clean_answer
                    })
                    answered = True
                    st.rerun()
                    break
                except Exception as e:
                    last_error = str(e)
                    continue
                    
            
