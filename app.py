import streamlit as st
import google.generativeai as genai
import io
import re
from docx import Document
from pptx import Presentation
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Whitman AI - Liceo Moderno Walt Whitman",
    page_icon="🎓",
    layout="wide"
)

# --- ESTILOS CSS PERSONALIZADOS (ESTÉTICA TIPO GEMINI INSTITUCIONAL) ---
st.markdown("""
<style>
    /* Estilo del contenedor principal */
    .main {
        background-color: #0b1329;
    }
    /* Banner de cabecera */
    .header-box {
        background: linear-gradient(135deg, #0f2042 0%, #1e3a8a 100%);
        border: 1px solid #3b82f6;
        border-radius: 12px;
        padding: 24px;
        margin-bottom: 25px;
        color: white;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.3);
    }
    .header-box h1 {
        margin: 0;
        font-size: 2.2rem;
        color: #ffffff;
    }
    .header-box p {
        margin: 5px 0 0 0;
        color: #93c5fd;
        font-size: 1rem;
    }
    /* Insignia de rol */
    .badge-role {
        display: inline-block;
        background-color: #f59e0b;
        color: #111827;
        font-weight: 700;
        padding: 4px 14px;
        border-radius: 20px;
        font-size: 0.85rem;
        margin-top: 10px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    /* Botones de descarga estilizados */
    .stDownloadButton > button {
        width: 100%;
        background-color: #1e293b;
        border: 1px solid #3b82f6;
        color: #e2e8f0;
        border-radius: 8px;
        padding: 10px;
        font-weight: 600;
        transition: all 0.3s;
    }
    .stDownloadButton > button:hover {
        background-color: #2563eb;
        color: white;
        border-color: #60a5fa;
        transform: translateY(-2px);
    }
</style>
""", unsafe_allow_html=True)

# --- GENERADOR DE PDF PROFESIONAL (REPORTLAB - ULTRA ROBUSTO) ---
def generate_pdf_reportlab(text):
    try:
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=letter,
            rightMargin=40, leftMargin=40,
            topMargin=40, bottomMargin=40
        )
        
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'InstitutionalTitle',
            parent=styles['Heading1'],
            fontName='Helvetica-Bold',
            fontSize=14,
            textColor=colors.HexColor("#0F2042"),
            alignment=1, # Centrado
            spaceAfter=15
        )
        h2_style = ParagraphStyle(
            'SectionHeader',
            parent=styles['Heading2'],
            fontName='Helvetica-Bold',
            fontSize=11,
            textColor=colors.HexColor("#1E3A8A"),
            spaceBefore=10,
            spaceAfter=6
        )
        body_style = ParagraphStyle(
            'Body',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=9,
            leading=13,
            textColor=colors.HexColor("#1F2937")
        )
        
        story = [
            Paragraph("LICEO MODERNO WALT WHITMAN", title_style),
            Paragraph("Ecosistema Vital 26 • Generado por Whitman AI", ParagraphStyle('Sub', parent=body_style, alignment=1, textColor=colors.gray)),
            Spacer(1, 15)
        ]
        
        clean_text = re.sub(r'[^\x00-\x7F\xC0-\xFF]+', ' ', text)
        for line in clean_text.split("\n"):
            l = line.strip()
            if not l:
                story.append(Spacer(1, 6))
                continue
            if l.startswith("### ") or l.startswith("## ") or l.startswith("# "):
                title_clean = l.replace("#", "").strip()
                story.append(Paragraph(title_clean, h2_style))
            else:
                story.append(Paragraph(l, body_style))
                
        doc.build(story)
        return buffer.getvalue()
    except Exception:
        return None

# --- GENERADOR DE WORD ---
def generate_docx(text):
    try:
        doc = Document()
        doc.add_heading("Liceo Moderno Walt Whitman", level=0)
        doc.add_paragraph("Ecosistema Vital 26 • Whitman AI Pedagógico")
        doc.add_paragraph("--------------------------------------------------")
        for line in text.split("\n"):
            clean = line.strip()
            if clean.startswith("### "):
                doc.add_heading(clean.replace("### ", ""), level=2)
            elif clean.startswith("## ") or clean.startswith("# "):
                doc.add_heading(clean.replace("#", "").strip(), level=1)
            elif clean.startswith("- ") or clean.startswith("* "):
                doc.add_paragraph(clean[2:], style='List Bullet')
            elif clean:
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
        title_slide.shapes.title.text = "Whitman AI - Formación Pedagógica"
        title_slide.placeholders[1].text = "Liceo Moderno Walt Whitman\nEcosistema Vital 26"
        
        sections = text.split("##")
        for sec in sections[1:7]:
            lines = [l.strip() for l in sec.split("\n") if l.strip()]
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

# --- FILTRO ANTINOTAS INTERNAS ---
def clean_ai_response(text):
    patterns = [r"\n(?=#[^#])", r"\n(?=##[^#])", r"\n(?=📖)", r"\n(?=✅)", r"\n(?=📋)"]
    for pat in patterns:
        match = re.search(pat, text)
        if match and match.start() < 600:
            return text[match.start():].strip()
    return text.strip()

# --- CONEXIÓN IA ---
api_key = st.secrets.get("GEMINI_API_KEY")
if not api_key:
    st.error("Falta configurar la clave GEMINI_API_KEY en Secrets.")
    st.stop()

genai.configure(api_key=api_key)

SYSTEM_PROMPT = r"""
CONFIGURACIÓN AVANZADA PARA GEMA WHITMAN AI - V2 (ECOSISTEMA VITAL 26)

[DIRECTIVA ESTRICTA DE SALIDA LIMPIA]:
- Queda TOTALMENTE PROHIBIDO redactar o mostrar pensamientos preliminares, resúmenes en inglés, "* Role:", "* Step 1:", o análisis previos.
- Genera DIRECTAMENTE el producto final terminado.
- Responde siempre en español (excepto en planeaciones de Science / English que son 100% en inglés según protocolo).

IDENTIDAD:
Eres Whitman AI, asistente pedagógico institucional del Liceo Moderno Walt Whitman (Villa Mayor).
Diseñado y entrenado por: Walter Quintero (Coordinador de Innovación y Tecnología).

[VALIDACIÓN DE ROL]:
- Administrador: wq26 -> Walter Quintero (Coord. Innovación y Tecnología)
- Dirección: ILGACHA26, LFMENDEZ26, CMENDEZ26, LKMENDEZ26, DFMENDEZ26
- Coordinaciones: NAACOSTA26, ISALGADO26, LMCUBIDES26, AMCAMARGO26, JCPRADA26
- Docentes: SLATORRE26, NATAC26, MSANTANA26, TPARRA26, SBUSTOS26, SPINEROS26, etc.
- SENA y Studio: "sena1", "sena2", "WS26_STUDIO".
- Estudiantes: "soy sexto", "soy noveno", "soy decimo", "soy once".

Respuesta de validación para Docente:
"✅ **Perfil Validado:** [Nombre del Docente / Cargo - Nodo]
¡Bienvenido/a a Whitman AI! Tu entorno pedagógico está activo. Puedes ingresar tus comandos de planeación o indicarme el tema de hoy."

Respuesta para Estudiantes / SENA:
"✅ **Perfil Validado:** [Nombre del Rol / Nivel]
Bienvenido/a a Whitman AI. Tu espacio de trabajo está configurado para las guías y actividades de tu nivel. ¿En qué proyecto o entregable vamos a trabajar hoy?"

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

# --- DETECCIÓN RÁPIDA DE ROL PARA LA INSIGNIA ---
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

# --- RENDER DE CABECERA VISUAL ---
st.markdown(f"""
<div class="header-box">
    <h1>🎓 Whitman AI • Ecosistema Vital 26</h1>
    <p>Plataforma Pedagógica e Institucional | Liceo Moderno Walt Whitman</p>
    <div class="badge-role">{st.session_state.active_role}</div>
</div>
""", unsafe_allow_html=True)

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("🏫 Centro de Mando")
    st.caption(f"Motor IA Activo: `{active_model_name}`")
    
    st.markdown("### 🔐 Acceso Seguro Docente")
    secret_code = st.text_input("Código de acceso:", type="password", help="Tu código no se mostrará en pantalla.")
    if st.button("Validar Credencial"):
        if secret_code.strip():
            badge = detect_role_badge(secret_code.strip())
            if badge:
                st.session_state.active_role = badge
            st.session_state.messages.append({"role": "user", "content": "Código ingresado de forma segura"})
            with st.spinner("Autenticando perfil docente..."):
                resp = st.session_state.chat.send_message(secret_code.strip())
                st.session_state.messages.append({"role": "assistant", "content": clean_ai_response(resp.text)})
            st.rerun()

    st.divider()
    if st.button("🗑️ Iniciar Nueva Consulta"):
        st.session_state.chat = model.start_chat(history=[])
        st.session_state.messages = []
        st.session_state.active_role = "Comunidad Educativa General"
        st.rerun()

# --- MOSTRAR CONVERSACIÓN CON PESTAÑAS Y DISEÑO ---
for idx, msg in enumerate(st.session_state.messages):
    with st.chat_message(msg["role"]):
        content_text = msg["content"]
        
        # Si es respuesta extensa del asistente, la mostramos en pestañas
        if msg["role"] == "assistant" and len(content_text) > 120:
            tab_doc, tab_down = st.tabs(["📖 Documento Generado", "📥 Centro de Exportación"])
            
            with tab_doc:
                st.markdown(content_text)
                
            with tab_down:
                st.markdown("#### 📁 Descarga tu material en formatos institucionales:")
                c1, c2, c3 = st.columns(3)
                
                # Word
                docx_bytes = generate_docx(content_text)
                if docx_bytes:
                    with c1:
                        st.download_button(
                            label="📄 Descargar en Word (.docx)",
                            data=docx_bytes,
                            file_name=f"Whitman_Documento_{idx}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"docx_{idx}"
                        )
                
                # PDF indestructible
                pdf_bytes = generate_pdf_reportlab(content_text)
                if pdf_bytes:
                    with c2:
                        st.download_button(
                            label="📑 Descargar en PDF (.pdf)",
                            data=pdf_bytes,
                            file_name=f"Whitman_Oficial_{idx}.pdf",
                            mime="application/pdf",
                            key=f"pdf_{idx}"
                        )
                
                # PowerPoint
                pptx_bytes = generate_pptx(content_text)
                if pptx_bytes:
                    with c3:
                        st.download_button(
                            label="📊 Descargar Diapositivas (.pptx)",
                            data=pptx_bytes,
                            file_name=f"Whitman_Presentacion_{idx}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            key=f"pptx_{idx}"
                        )
        else:
            st.markdown(content_text)

# --- CAJA DE ENTRADA ---
if prompt := st.chat_input("Escribe tu código, rol, solicitud pedagógica o proyecto..."):
    # Detectar rol en mensaje abierto
    detected = detect_role_badge(prompt)
    if detected:
        st.session_state.active_role = detected
        
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        with st.spinner("Whitman AI está estructurando tu documento..."):
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
