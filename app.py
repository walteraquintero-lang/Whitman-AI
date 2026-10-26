import streamlit as st
import google.generativeai as genai
import io
import re
import urllib.parse
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
        padding: 18px 22px;
        margin-bottom: 20px;
        color: white;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25);
    }
    .header-box h1 { margin: 0; font-size: 1.8rem; color: #ffffff; }
    .header-box p { margin: 3px 0 0 0; color: #93c5fd; font-size: 0.9rem; }
    .badge-role {
        display: inline-block;
        background-color: #f59e0b;
        color: #0f2042;
        font-weight: 800;
        padding: 4px 14px;
        border-radius: 15px;
        font-size: 0.8rem;
        margin-top: 8px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .welcome-card {
        background-color: #111e38;
        border: 1px solid #1e3a8a;
        border-radius: 8px;
        padding: 18px;
        margin: 10px 0;
        border-left: 5px solid #3b82f6;
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
    }
</style>
""", unsafe_allow_html=True)

# --- MATRIZ DE AUTENTICACIÓN Y PERFILES (EN PYTHON PURO - CERO FUGAS DE IA) ---
PROFILES = {
    # Administrador
    "wq26": {
        "role_name": "👑 Coordinador de Innovación & Tecnología",
        "badge": "ADMINISTRACIÓN / INNOVACIÓN",
        "welcome": "¡Bienvenido, Ing. Walter Quintero! Centro de Control y Gestión Tecnológica activo. Tienes acceso completo a la configuración pedagógica, herramientas audiovisuales de Whitman Studio y protocolos de IA institucional.",
        "prompt_mode": "Eres el copiloto técnico del Coordinador Walter Quintero. Tono profesional, estratégico, técnico y de alta innovación educativa."
    },
    # Docentes Clave
    "slatorre26": {
        "role_name": "🔬 Teacher Sonia Latorre (Science Primaria)",
        "badge": "DOCENTE BILINGÜE - SCIENCE",
        "welcome": "¡Welcome, Teacher Sonia! Tus planeaciones de Science (1° a 5°) están configuradas al estándar bilingüe de Ciudadanos del Mundo (Modelo Natalia Cruz). Todo el contenido se generará 100% en Technical Science English bajo la metodología 70/30 con vinculación a ODS.",
        "prompt_mode": "Usuario es Sonia Latorre. Asignatura: Science. Todas las planeaciones DEBEN ser 100% en inglés, metodología 70/30, formato Ciudadanos del Mundo, ODS e IA task."
    },
    "natac26": {
        "role_name": "🌍 Docente Natalia Cruz (Ciudadanos del Mundo)",
        "badge": "DOCENTE BILINGÜE - CIUDADANOS",
        "welcome": "¡Bienvenida, Teacher Natalia! Tu entorno para Ciudadanos del Mundo (5°, 6° y 7°A) está activo con la matriz de 6 momentos bilingüe y producción oral activa.",
        "prompt_mode": "Usuario es Natalia Cruz. Enfoque: Ciudadanos del Mundo bilingüe, 70/30 y desarrollo global."
    },
    # Aprendices SENA y Whitman Studio
    "sena1": {
        "role_name": "🎨 Aprendiz SENA - Fase 1 (Grado 10°)",
        "badge": "APRENDIZ SENA - FASE 1",
        "welcome": "¡Acceso Concedido, Aprendiz SENA! 🎨 Tu laboratorio técnico de Diseño Multimedia está activo. Recuerda los estándares del Coordinador Walter Quintero: proyectos vectoriales puros en Adobe Illustrator (.AI), mesas de trabajo organizadas en 3 capas ([Fondos], [Vectores], [Textos]) y fuentes convertidas a curvas.",
        "prompt_mode": "Usuario es Aprendiz SENA Fase 1 (10°). Foco en diseño en Adobe Illustrator, capas, curvas, teoría del color y guías de póster."
    },
    "sena2": {
        "role_name": "🚀 Aprendiz SENA - Fase 2 (Grado 11° / Proyecto Productivo)",
        "badge": "APRENDIZ SENA - FASE 2 (PROYECTO)",
        "welcome": "¡Acceso Concedido, Aprendiz SENA! 🚀 Tu entorno para el Proyecto Productivo (Fase 2) está activo. Estoy listo para asesorar la investigación de mercados, la matriz de segmentación del Reto 5 (Demográfica, Geográfica, Psicográfica y Conductual) y la ficha del Buyer Persona.",
        "prompt_mode": "Usuario es Aprendiz SENA Fase 2 (11°). Asesoría en Proyecto Productivo, Reto 5, matriz de segmentación para Mesa 1 y Buyer Persona para Mesa 2 en Illustrator."
    },
    "ws26_studio": {
        "role_name": "🎬 Integrante de Whitman Studio",
        "badge": "WHITMAN STUDIO CREATIVO",
        "welcome": "¡Acceso al Semillero Creativo Concedido! 🎥 Tu espacio para storytelling, hooks virales, guiones audiovisuales, diseño de marca y producción digital está listo.",
        "prompt_mode": "Usuario es integrante de Whitman Studio. Rol de Director de Arte y mentor audiovisual escolar."
    },
    # Padres de Familia
    "padre": {
        "role_name": "🏡 Familia Whitman",
        "badge": "PADRES DE FAMILIA",
        "welcome": "¡Estimada Familia del Liceo Moderno Walt Whitman, bienvenidos! 🌟 Este es su espacio de orientación institucional. Estoy aquí para acompañarles con información sobre el modelo pedagógico, actividades escolares, convivencia y apoyo en el proceso formativo de sus hijos.",
        "prompt_mode": "Hablas con un padre de familia. Sé extraordinariamente cálido, respetuoso, empático y claro. No uses jerga técnica de planeación docente. Orienta sobre la formación integral de los estudiantes."
    }
}

# --- FUNCIÓN DE DETECCIÓN INTELIGENTE ---
def check_user_identity(text):
    clean = text.lower().strip()
    # Coincidencias directas
    if "wq26" in clean:
        return PROFILES["wq26"]
    elif "slatorre26" in clean or ("sonia" in clean and "latorre" in clean):
        return PROFILES["slatorre26"]
    elif "natac26" in clean:
        return PROFILES["natac26"]
    elif any(k in clean for k in ["sena 2", "sena2", "soy sena 2", "fase 2", "once sena"]):
        return PROFILES["sena2"]
    elif any(k in clean for k in ["sena 1", "sena1", "soy sena 1", "fase 1", "decimo sena"]):
        return PROFILES["sena1"]
    elif "ws26" in clean or "studio" in clean:
        return PROFILES["ws26_studio"]
    elif any(k in clean for k in ["padre", "mama", "mamá", "papa", "papá", "acudiente", "familia"]):
        return PROFILES["padre"]
    elif any(k in clean for k in ["noveno", "9°", "soy noveno", "grado 9"]):
        return {
            "role_name": "🎒 Estudiante de Grado 9°",
            "badge": "ESTUDIANTE NOVENO",
            "welcome": "¡Hola, estudiante de Noveno! 📘 Tu espacio de aprendizaje está activo. ¿En qué temática o proyecto te acompaño hoy?",
            "prompt_mode": "Tutor pedagógico socrático para grado noveno. Explica paso a paso con paciencia, sin dar respuestas directas a tareas."
        }
    return None

# --- CONEXIÓN IA ---
api_key = st.secrets.get("GEMINI_API_KEY")
if not api_key:
    st.error("Falta configurar la clave GEMINI_API_KEY en Secrets.")
    st.stop()

genai.configure(api_key=api_key)

SYSTEM_BASE = """
Eres Whitman AI, asistente pedagógico e institucional oficial del Liceo Moderno Walt Whitman (Villa Mayor, Bogotá).
Diseñado y entrenado por: Walter Quintero (Coordinador de Innovación y Tecnología).

REGLA CRÍTICA DE SALIDA:
- NUNCA generes notas preliminares, ni "User input:", ni "Goal:", ni "Context:", ni listas de chequeo en inglés.
- Responde DIRECTAMENTE con el contenido solicitado, estructurado con títulos limpios, tablas o viñetas.

CIERRE OBLIGATORIO:
---
> *"El arte de la educación es el arte de despertar la curiosidad."* — Walt Whitman

**Resultados potenciados por Whitman AI | Entrenado por: Walter Quintero (Coord. Innovación y Tecnología)**
"""

@st.cache_resource
def get_engine():
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
            m = genai.GenerativeModel(model_name=model_name, system_instruction=SYSTEM_BASE)
            m.generate_content("ping")
            return model_name, m
        except Exception:
            continue
    raise RuntimeError("No se encontró ningún modelo compatible habilitado.")

active_model_name, base_model = get_engine()

# --- ESTADOS DE SESIÓN ---
if "messages" not in st.session_state:
    st.session_state.messages = []
if "active_profile" not in st.session_state:
    st.session_state.active_profile = {
        "role_name": "Comunidad Educativa General",
        "badge": "VISITANTE GENERAL",
        "prompt_mode": "Asistente institucional del colegio para toda la comunidad."
    }

# --- GENERADORES DE ARCHIVOS ---
def generate_pdf_reportlab(text):
    try:
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('LiceoTitle', fontName='Helvetica-Bold', fontSize=14, textColor=colors.HexColor("#0F2042"), alignment=1, spaceAfter=4)
        sub_style = ParagraphStyle('LiceoSub', fontName='Helvetica', fontSize=8, textColor=colors.HexColor("#64748B"), alignment=1, spaceAfter=14)
        h1_style = ParagraphStyle('LiceoH1', fontName='Helvetica-Bold', fontSize=10.5, textColor=colors.HexColor("#0F2042"), spaceBefore=10, spaceAfter=4)
        body_style = ParagraphStyle('LiceoBody', fontName='Helvetica', fontSize=8.5, leading=11, textColor=colors.HexColor("#1E293B"))
        th_style = ParagraphStyle('LiceoTH', fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.white, alignment=1)
        td_style = ParagraphStyle('LiceoTD', fontName='Helvetica', fontSize=7, leading=8.5, textColor=colors.HexColor("#0F172A"))
        
        story = [
            Paragraph("LICEO MODERNO WALT WHITMAN", title_style),
            Paragraph("Ecosistema Vital 26 • Formando Líderes Transformadores", sub_style),
            Spacer(1, 6)
        ]
        
        clean_text = re.sub(r'[^\x00-\x7F\xC0-\xFF]+', ' ', text)
        lines = clean_text.split("\n")
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if line.startswith("|") and line.endswith("|"):
                table_raw = []
                while i < len(lines) and lines[i].strip().startswith("|") and lines[i].strip().endswith("|"):
                    row_str = lines[i].strip()
                    if not re.match(r"^\|(\s*[-:]+\s*\|)+$", row_str):
                        cells = [c.strip() for c in row_str.split("|")[1:-1]]
                        table_raw.append(cells)
                    i += 1
                if table_raw:
                    num_cols = max(len(r) for r in table_raw)
                    col_w = 540 / max(num_cols, 1)
                    table_data = []
                    for r_idx, r in enumerate(table_raw):
                        row_cells = []
                        for c in r:
                            row_cells.append(Paragraph(c[:180], th_style if r_idx == 0 else td_style))
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
                    story.append(Spacer(1, 5))
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

# --- HEADER PRINCIPAL ---
st.markdown(f"""
<div class="header-box">
    <h1>🎓 Whitman AI • Ecosistema Vital 26</h1>
    <p>Liceo Moderno Walt Whitman | Asistente Integral Escolar</p>
    <div class="badge-role">{st.session_state.active_profile['badge']}</div>
</div>
""", unsafe_allow_html=True)

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("🏫 Selector de Perfil")
    st.caption("Selecciona tu rol para una experiencia adaptada:")
    
    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("👑 Administrador"):
            st.session_state.active_profile = PROFILES["wq26"]
            st.session_state.messages.append({"role": "assistant", "content": PROFILES["wq26"]["welcome"]})
            st.rerun()
        if st.button("🚀 SENA 2 (11°)"):
            st.session_state.active_profile = PROFILES["sena2"]
            st.session_state.messages.append({"role": "assistant", "content": PROFILES["sena2"]["welcome"]})
            st.rerun()
    with col_b:
        if st.button("🔬 Science Docente"):
            st.session_state.active_profile = PROFILES["slatorre26"]
            st.session_state.messages.append({"role": "assistant", "content": PROFILES["slatorre26"]["welcome"]})
            st.rerun()
        if st.button("🏡 Padre de Familia"):
            st.session_state.active_profile = PROFILES["padre"]
            st.session_state.messages.append({"role": "assistant", "content": PROFILES["padre"]["welcome"]})
            st.rerun()
            
    st.divider()
    if st.button("🗑️ Reiniciar Sesión"):
        st.session_state.messages = []
        st.session_state.active_profile = {
            "role_name": "Comunidad Educativa General",
            "badge": "VISITANTE GENERAL",
            "prompt_mode": "Asistente institucional del colegio para toda la comunidad."
        }
        st.rerun()

# --- MOSTRAR CONVERSACIÓN ---
for idx, msg in enumerate(st.session_state.messages):
    with st.chat_message(msg["role"]):
        content_text = msg["content"]
        st.markdown(content_text)
        
        # Botones de exportación al final de respuestas largas
        if msg["role"] == "assistant" and len(content_text) > 150:
            st.markdown("""
            <div style="margin-top: 15px; border-top: 1px solid #1e3a8a; padding-top: 8px;">
                <b style="color: #93c5fd; font-size: 0.85rem;">📥 Descargar Documento Formateado:</b>
            </div>
            """, unsafe_allow_html=True)
            
            c1, c2, c3 = st.columns(3)
            docx_bytes = generate_docx(content_text)
            if docx_bytes:
                with c1:
                    st.download_button("📄 Word (.docx)", docx_bytes, f"Whitman_{idx}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"docx_{idx}")
            pdf_bytes = generate_pdf_reportlab(content_text)
            if pdf_bytes:
                with c2:
                    st.download_button("📑 PDF Oficial", pdf_bytes, f"Whitman_{idx}.pdf", "application/pdf", key=f"pdf_{idx}")
            pptx_bytes = generate_pptx(content_text)
            if pptx_bytes:
                with c3:
                    st.download_button("📊 Diapositivas (.pptx)", pptx_bytes, f"Whitman_{idx}.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", key=f"pptx_{idx}")

# --- ENTRADA DE MENSAJES ---
if prompt := st.chat_input("Escribe tu pregunta, solicitud o identifícate..."):
    # 1. Mostrar mensaje del usuario
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # 2. VALIDACIÓN DIRECTA EN PYTHON (CERO PENSAMIENTOS DE IA)
    matched_identity = check_user_identity(prompt)
    
    if matched_identity:
        # Si el usuario solo se estaba identificando (ej. "Soy Sena 2", "wq26", "soy padre")
        st.session_state.active_profile = matched_identity
        welcome_msg = matched_identity["welcome"]
        
        with st.chat_message("assistant"):
            st.markdown(welcome_msg)
        st.session_state.messages.append({"role": "assistant", "content": welcome_msg})
        st.rerun()

    else:
        # 3. CONSULTA REGULAR A LA IA (PERSONALIZADA AL PERFIL ACTIVO)
        with st.chat_message("assistant"):
            with st.spinner("Whitman AI está estructurando tu respuesta..."):
                try:
                    role_context = st.session_state.active_profile.get("prompt_mode", "")
                    custom_prompt = f"""
[PERFIL ACTIVO DEL USUARIO: {st.session_state.active_profile['role_name']}]
DIRECTIVA DE ROL: {role_context}

SOLICITUD: {prompt}

Responde directamente en español al usuario según su perfil, sin notas internas ni prefacios.
"""
                    res = base_model.generate_content(custom_prompt)
                    ans = res.text.strip()
                    
                    # Filtro de seguridad adicional contra fugas
                    if "User input:" in ans or "The instruction says:" in ans:
                        idx_cut = ans.rfind("---")
                        if idx_cut != -1:
                            ans = ans[idx_cut:].strip()
                            
                    st.markdown(ans)
                    st.session_state.messages.append({"role": "assistant", "content": ans})
                    st.rerun()
                except Exception as e:
                    err_msg = f"No pudimos conectar temporalmente con el servidor. Por favor reintenta tu pregunta. ({str(e)})"
                    st.markdown(err_msg)
                    st.session_state.messages.append({"role": "assistant", "content": err_msg})
