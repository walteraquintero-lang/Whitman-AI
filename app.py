import streamlit as st
import google.generativeai as genai
import io
import re
from docx import Document
from pptx import Presentation
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Whitman AI - Liceo Moderno Walt Whitman",
    page_icon="🎓",
    layout="wide"
)

# --- ESTILOS CSS INSTITUCIONALES ---
st.markdown("""
<style>
    .header-box {
        background: linear-gradient(135deg, #0f2042 0%, #1e3a8a 100%);
        border-left: 6px solid #f59e0b;
        border-radius: 10px;
        padding: 20px;
        margin-bottom: 20px;
        color: white;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25);
    }
    .header-box h1 { margin: 0; font-size: 2rem; color: #ffffff; }
    .header-box p { margin: 4px 0 0 0; color: #93c5fd; font-size: 0.95rem; }
    .badge-role {
        display: inline-block;
        background-color: #f59e0b;
        color: #0f2042;
        font-weight: 800;
        padding: 4px 12px;
        border-radius: 15px;
        font-size: 0.8rem;
        margin-top: 8px;
        text-transform: uppercase;
    }
    .export-box {
        background-color: #111e38;
        border: 1px solid #1e3a8a;
        border-radius: 8px;
        padding: 15px;
        margin-top: 15px;
    }
    .stDownloadButton > button {
        width: 100%;
        background-color: #0f2042;
        border: 1px solid #f59e0b;
        color: #ffffff;
        font-weight: 600;
        border-radius: 6px;
        padding: 8px;
    }
    .stDownloadButton > button:hover {
        background-color: #f59e0b;
        color: #0f2042;
        border-color: #ffffff;
    }
</style>
""", unsafe_allow_html=True)

# --- FILTRO PROFUNDO DE NOTAS INTERNAS / PENSAMIENTOS ---
def clean_ai_response(text):
    if not text:
        return ""
    # Si detecta notas preliminares en inglés o autocrítica
    thought_triggers = [
        r"\*Target Audience", r"\*Context:", r"\*Requirements:", 
        r"\*Persona:", r"\(Self-Correction", r"\*Final Polish"
    ]
    has_thoughts = any(re.search(trigger, text, re.IGNORECASE) for trigger in thought_triggers)
    
    if has_thoughts:
        # Buscar el inicio formal: primer título oficial, saludo o tabla
        start_markers = [
            r"\n(?=#[^#])", r"\n(?=##[^#])", r"\n(?=###[^#])",
            r"\n(?=\*\*Título)", r"\n(?=📖)", r"\n(?=✅)", r"\n(?=📋)"
        ]
        for marker in start_markers:
            match = re.search(marker, text)
            if match and match.start() > 50:
                return text[match.start():].strip()
    return text.strip()

# --- GENERADOR DE PDF INSTITUCIONAL CON TABLAS REALES ---
def generate_pdf_reportlab(text):
    try:
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer, pagesize=letter,
            rightMargin=36, leftMargin=36,
            topMargin=36, bottomMargin=36
        )
        
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'LiceoTitle', fontName='Helvetica-Bold', fontSize=14,
            textColor=colors.HexColor("#0F2042"), alignment=1, spaceAfter=4
        )
        sub_style = ParagraphStyle(
            'LiceoSub', fontName='Helvetica', fontSize=8,
            textColor=colors.HexColor("#64748B"), alignment=1, spaceAfter=14
        )
        h1_style = ParagraphStyle(
            'LiceoH1', fontName='Helvetica-Bold', fontSize=11,
            textColor=colors.HexColor("#0F2042"), spaceBefore=10, spaceAfter=5
        )
        body_style = ParagraphStyle(
            'LiceoBody', fontName='Helvetica', fontSize=8.5,
            leading=11.5, textColor=colors.HexColor("#1E293B")
        )
        th_style = ParagraphStyle(
            'LiceoTH', fontName='Helvetica-Bold', fontSize=7.5,
            leading=9, textColor=colors.white, alignment=1
        )
        td_style = ParagraphStyle(
            'LiceoTD', fontName='Helvetica', fontSize=7,
            leading=8.5, textColor=colors.HexColor("#0F172A")
        )
        
        story = [
            Paragraph("LICEO MODERNO WALT WHITMAN", title_style),
            Paragraph("Ecosistema Vital 26 • Formando Líderes Transformadores", sub_style),
            Spacer(1, 8)
        ]
        
        clean_text = re.sub(r'[^\x00-\x7F\xC0-\xFF]+', ' ', text)
        lines = clean_text.split("\n")
        i = 0
        
        while i < len(lines):
            line = lines[i].strip()
            
            # Detectar tablas de Markdown
            if line.startswith("|") and line.endswith("|"):
                table_raw = []
                while i < len(lines) and lines[i].strip().startswith("|") and lines[i].strip().endswith("|"):
                    row_str = lines[i].strip()
                    # Omitir línea divisoria Markdown (|---|---|)
                    if not re.match(r"^\|(\s*[-:]+\s*\|)+$", row_str):
                        cells = [c.strip() for c in row_str.split("|")[1:-1]]
                        table_raw.append(cells)
                    i += 1
                
                if table_raw:
                    # Determinar ancho por columna
                    num_cols = max(len(r) for r in table_raw)
                    available_width = 540 # Ancho imprimible carta
                    col_w = available_width / max(num_cols, 1)
                    
                    table_data = []
                    for r_idx, r in enumerate(table_raw):
                        row_cells = []
                        for c in r:
                            style_use = th_style if r_idx == 0 else td_style
                            row_cells.append(Paragraph(c[:180], style_use))
                        # Rellenar columnas faltantes si la fila es corta
                        while len(row_cells) < num_cols:
                            row_cells.append(Paragraph("", td_style))
                        table_data.append(row_cells)
                    
                    t = Table(table_data, colWidths=[col_w] * num_cols)
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0F2042")),
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                        ('TOPPADDING', (0, 0), (-1, -1), 4),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                    ]))
                    story.append(Spacer(1, 4))
                    story.append(t)
                    story.append(Spacer(1, 6))
                continue
            
            # Títulos y texto normal
            if line.startswith("#"):
                clean_title = re.sub(r"^#+\s*", "", line)
                story.append(Spacer(1, 4))
                story.append(Paragraph(clean_title, h1_style))
            elif line:
                story.append(Paragraph(line, body_style))
                story.append(Spacer(1, 2))
            else:
                story.append(Spacer(1, 4))
            i += 1
            
        doc.build(story)
        return buffer.getvalue()
    except Exception:
        return None

# --- GENERADOR DE WORD ---
def generate_docx(text):
    try:
        doc = Document()
        doc.add_heading("Liceo Moderno Walt Whitman", level=0)
        doc.add_paragraph("Ecosistema Vital 26 • Formación Pedagógica Institucional")
        doc.add_paragraph("--------------------------------------------------")
        for line in text.split("\n"):
            clean = line.strip()
            if clean.startswith("### "):
                doc.add_heading(clean.replace("### ", ""), level=2)
            elif clean.startswith("## ") or clean.startswith("# "):
                doc.add_heading(clean.replace("#", "").strip(), level=1)
            elif clean.startswith("- ") or clean.startswith("* "):
                doc.add_paragraph(clean[2:], style='List Bullet')
            elif clean and not clean.startswith("|"):
                doc.add_paragraph(clean)
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    except Exception:
        return None

# --- GENERADOR DE POWERPOINT ---
def generate_pptx(text):
    try:
        prs = Presentation()
        title_slide = prs.slides.add_slide(prs.slide_layouts[0])
        title_slide.shapes.title.text = "Whitman AI - Presentación"
        title_slide.placeholders[1].text = "Liceo Moderno Walt Whitman\nEcosistema Vital 26"
        sections = text.split("##")
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

# --- CONEXIÓN IA ---
api_key = st.secrets.get("GEMINI_API_KEY")
if not api_key:
    st.error("Falta configurar la clave GEMINI_API_KEY en Secrets.")
    st.stop()

genai.configure(api_key=api_key)

SYSTEM_PROMPT = r"""
CONFIGURACIÓN AVANZADA PARA GEMA WHITMAN AI - V2 (ECOSISTEMA VITAL 26)

[DIRECTIVA ESTRICTA DE SALIDA LIMPIA Y PROFESIONAL]:
- NUNCA generes ni expongas procesos de pensamiento, "Target Audience:", "Context:", "Persona:", "Requirements:", ni "Self-Correction".
- Responde DIRECTAMENTE con el contenido solicitado, estructurado con títulos claros, tablas y viñetas pedagógicas.
- Si el usuario solicita un guion, póster, planeación o reto, entrégalo listo para producción o impresión.

IDENTIDAD:
Eres Whitman AI, asistente pedagógico del Liceo Moderno Walt Whitman (Villa Mayor).
Diseñado y entrenado por: Walter Quintero (Coordinador de Innovación y Tecnología).

VALIDACIÓN DE ROL:
- Administrador: wq26 -> Walter Quintero (Coord. Innovación y Tecnología)
- Dirección: ILGACHA26, LFMENDEZ26, CMENDEZ26, LKMENDEZ26, DFMENDEZ26
- Coordinaciones: NAACOSTA26, ISALGADO26, LMCUBIDES26, AMCAMARGO26, JCPRADA26
- Docentes: SLATORRE26 (Science), NATAC26, MSANTANA26, TPARRA26, SBUSTOS26, SPINEROS26, etc.
- SENA y Studio: "sena1", "sena2", "WS26_STUDIO".
- Estudiantes: "soy sexto", "soy noveno", "soy decimo", "soy once".

Respuesta de validación:
"✅ **Perfil Validado:** [Nombre del Rol / Docente - Nodo]
¡Bienvenido/a a Whitman AI! Tu entorno pedagógico está activo. ¿En qué proyecto o planeación trabajaremos hoy?"

CIERRE OBLIGATORIO:
---
> *"El arte de la educación es el arte de despertar la curiosidad."* — Walt Whitman

**Resultados potenciados por Whitman AI | Entrenado por: Walter Quintero (Coord. Innovación y Tecnología)**
"""

@st.cache_resource
def get_working_engine():
    candidate_names = []
    try:
        for m in genai.list_models():
            if "generateContent" in m.supported_generation_methods:
                candidate_names.append(m.name)
    except Exception:
        pass
    for fb in ["gemini-2.0-flash", "gemini-1.5-pro", "gemini-pro"]:
        if fb not in candidate_names:
            candidate_names.append(fb)
    for model_name in candidate_names:
        try:
            test_m = genai.GenerativeModel(model_name=model_name)
            test_m.generate_content("ping")
            return model_name, genai.GenerativeModel(model_name=model_name, system_instruction=SYSTEM_PROMPT)
        except Exception:
            continue
    raise RuntimeError("No se encontró ningún modelo compatible habilitado.")

active_model_name, model = get_working_engine()

# --- SESIÓN ---
if "chat" not in st.session_state:
    st.session_state.chat = model.start_chat(history=[])
if "messages" not in st.session_state:
    st.session_state.messages = []
if "active_role" not in st.session_state:
    st.session_state.active_role = "Comunidad Educativa General"

def detect_role_badge(text):
    t = text.lower()
    if "wq26" in t:
        return "👑 Administrador / Innovación (Walter Quintero)"
    elif "slatorre26" in t or "science" in t:
        return "🔬 Science & Bilingüismo (Docente)"
    elif "sena1" in t:
        return "🎨 Aprendiz SENA - Fase 1 (10°)"
    elif "sena2" in t:
        return "🚀 Aprendiz SENA - Fase 2 (11°)"
    elif "ws26_studio" in t:
        return "🎬 Whitman Studio (Creativo)"
    return None

# --- HEADER VISUAL ---
st.markdown(f"""
<div class="header-box">
    <h1>🎓 Whitman AI • Ecosistema Vital 26</h1>
    <p>Plataforma Pedagógica e Institucional | Liceo Moderno Walt Whitman</p>
    <div class="badge-role">{st.session_state.active_role}</div>
</div>
""", unsafe_allow_html=True)

# --- SIDEBAR ---
with st.sidebar:
    st.header("🏫 Centro de Mando")
    st.caption(f"Motor IA Activo: `{active_model_name}`")
    st.markdown("### 🔐 Acceso Seguro Docente")
    secret_code = st.text_input("Código institucional:", type="password")
    if st.button("Validar Credencial"):
        if secret_code.strip():
            b = detect_role_badge(secret_code.strip())
            if b:
                st.session_state.active_role = b
            st.session_state.messages.append({"role": "user", "content": "Código validado por el usuario"})
            with st.spinner("Autenticando..."):
                resp = st.session_state.chat.send_message(secret_code.strip())
                st.session_state.messages.append({"role": "assistant", "content": clean_ai_response(resp.text)})
            st.rerun()

    st.divider()
    if st.button("🗑️ Iniciar Nueva Consulta"):
        st.session_state.chat = model.start_chat(history=[])
        st.session_state.messages = []
        st.session_state.active_role = "Comunidad Educativa General"
        st.rerun()

# --- MOSTRAR CONVERSACIÓN ---
for idx, msg in enumerate(st.session_state.messages):
    with st.chat_message(msg["role"]):
        content_text = msg["content"]
        
        # 1. MOSTRAR EL DOCUMENTO COMPLETO PRIMERO
        st.markdown(content_text)
        
        # 2. BOTONES AL FINAL DEL MENSAJE (SOLO SI ES UNA RESPUESTA DEL ASISTENTE)
        if msg["role"] == "assistant" and len(content_text) > 120:
            st.markdown("""
            <div style="margin-top: 20px; border-top: 1px solid #1e3a8a; padding-top: 10px;">
                <b style="color: #93c5fd;">📥 Centro de Exportación Institucional:</b>
            </div>
            """, unsafe_allow_html=True)
            
            c1, c2, c3 = st.columns(3)
            docx_bytes = generate_docx(content_text)
            if docx_bytes:
                with c1:
                    st.download_button(
                        label="📄 Descargar Word (.docx)",
                        data=docx_bytes,
                        file_name=f"Whitman_Guia_{idx}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"docx_{idx}"
                    )
            pdf_bytes = generate_pdf_reportlab(content_text)
            if pdf_bytes:
                with c2:
                    st.download_button(
                        label="📑 Descargar PDF Oficial",
                        data=pdf_bytes,
                        file_name=f"Whitman_Oficial_{idx}.pdf",
                        mime="application/pdf",
                        key=f"pdf_{idx}"
                    )
            pptx_bytes = generate_pptx(content_text)
            if pptx_bytes:
                with c3:
                    st.download_button(
                        label="📊 Descargar Presentación (.pptx)",
                        data=pptx_bytes,
                        file_name=f"Whitman_Diapositivas_{idx}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key=f"pptx_{idx}"
                    )

# --- ENTRADA ---
if prompt := st.chat_input("Escribe tu solicitud, comando de planeación o proyecto..."):
    detected = detect_role_badge(prompt)
    if detected:
        st.session_state.active_role = detected
        
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        with st.spinner("Whitman AI está estructurando tu material institucional..."):
            try:
                response = st.session_state.chat.send_message(prompt)
                clean_text = clean_ai_response(response.text)
                st.session_state.messages.append({"role": "assistant", "content": clean_text})
                st.rerun()
            except Exception:
                st.session_state.chat = model.start_chat(history=[])
                response = st.session_state.chat.send_message(prompt)
                clean_text = clean_ai_response(response.text)
                st.session_state.messages.append({"role": "assistant", "content": clean_text})
                st.rerun()
