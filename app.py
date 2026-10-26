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
        padding: 16px 20px;
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
        padding: 4px 12px;
        border-radius: 15px;
        font-size: 0.75rem;
        margin-top: 8px;
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
    .docs-btn:hover {
        background-color: #2563eb;
    }
</style>
""", unsafe_allow_html=True)

# --- PERFILES ---
PROFILES = {
    "wq26": {
        "role_name": "👑 Coordinador de Innovación & Tecnología",
        "badge": "ADMINISTRACIÓN / INNOVACIÓN",
        "welcome": "¡Bienvenido, Ing. Walter Quintero! Centro de Control activo. Acceso a diseño pedagógico, proyectos SENA y herramientas audiovisuales de Whitman Studio.",
        "prompt_mode": "Usuario: Coordinador Walter Quintero. Tono estratégico, directivo, técnico y de innovación escolar."
    },
    "sena2": {
        "role_name": "🚀 Aprendiz SENA - Fase 2 (11° / Proyecto)",
        "badge": "APRENDIZ SENA - FASE 2",
        "welcome": "¡Acceso Concedido, Aprendiz SENA! 🚀 Tu entorno para el Proyecto Productivo (Fase 2) está listo. Podemos estructurar tarjetas de invitación, piezas gráficas para las sustentaciones o matrices para Illustrator.",
        "prompt_mode": "Usuario: Aprendiz SENA Fase 2 (11°). Foco en Proyecto Productivo, sustentaciones, Adobe Illustrator y piezas de difusión."
    },
    "padre": {
        "role_name": "🏡 Familia Whitman",
        "badge": "PADRES DE FAMILIA",
        "welcome": "¡Estimada Familia del Liceo Moderno Walt Whitman, bienvenidos! 🌟 Este es su espacio de orientación escolar. ¿En qué información sobre eventos o sustentaciones podemos orientarle hoy?",
        "prompt_mode": "Hablas con un padre de familia. Tono institucional, cálido, formal y claro. Orienta sobre eventos, sustentaciones y vida escolar sin jerga técnica."
    }
}

def check_user_identity(text):
    clean = text.lower().strip()
    action_words = ["crea", "creemos", "haz", "diseña", "redacta", "planea", "escribe", "genera", "comunicado", "tarjeta", "invitacion", "invitación"]
    if any(w in clean for w in action_words) or len(clean.split()) > 5:
        return None
    if "wq26" in clean:
        return PROFILES["wq26"]
    elif any(k in clean for k in ["sena 2", "sena2", "soy sena 2", "fase 2"]):
        return PROFILES["sena2"]
    elif clean in ["padre", "papa", "papá", "mama", "mamá", "soy padre", "acudiente"]:
        return PROFILES["padre"]
    return None

def generate_image_url(prompt_text):
    clean = re.sub(r'[^\w\s]', ' ', prompt_text).strip()
    encoded = urllib.parse.quote(clean[:200])
    return f"https://image.pollinations.ai/prompt/{encoded}?width=900&height=500&nologo=true"

# --- CONEXIÓN DIRECTA CON EL MOTOR QUE FUNCIONA ---
api_key = st.secrets.get("GEMINI_API_KEY")
if not api_key:
    st.error("Falta configurar GEMINI_API_KEY en Secrets.")
    st.stop()

genai.configure(api_key=api_key)

SYSTEM_BASE = """
Eres Whitman AI, asistente pedagógico e institucional oficial del Liceo Moderno Walt Whitman (Villa Mayor, Bogotá).
Diseñado y entrenado por: Walter Quintero (Coordinador de Innovación y Tecnología).

DIRECTIVAS INSTITUCIONALES:
1. COMUNICADOS: Redáctalos formalmente con membrete, número de circular, fecha, lugar (Teatro Villa Mayor), orden del día, recomendaciones y firma de Rectoría / Coordinación.
2. TARJETAS DE INVITACIÓN: Diseña la estructura visual y textual completa (Titular, Fecha, Hora, Lugar, Protocolo de vestimenta, Código de confirmación).
3. GUÍAS Y PLANEACIONES: Usa tablas estructuradas con momentos de clase (DBA, Ecosistema Vital).
4. NUNCA generes notas preliminares en inglés ni textos de pensamiento. Responde directamente.

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
    for fb in ["gemini-2.0-flash", "gemini-1.5-pro", "gemini-1.5-flash"]:
        if fb not in candidate_names:
            candidate_names.append(fb)
    for model_name in candidate_names:
        try:
            test_m = genai.GenerativeModel(model_name=model_name, system_instruction=SYSTEM_BASE)
            test_m.generate_content("ping")
            return model_name, test_m
        except Exception:
            continue
    raise RuntimeError("No se encontró ningún modelo compatible habilitado.")

active_model_name, active_model = get_working_engine()

# --- GENERADOR DE PDF ---
def generate_pdf_reportlab(text):
    try:
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('Title', fontName='Helvetica-Bold', fontSize=13, textColor=colors.HexColor("#0F2042"), alignment=1, spaceAfter=4)
        sub_style = ParagraphStyle('Sub', fontName='Helvetica', fontSize=8, textColor=colors.HexColor("#64748B"), alignment=1, spaceAfter=12)
        h1_style = ParagraphStyle('H1', fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor("#0F2042"), spaceBefore=8, spaceAfter=4)
        body_style = ParagraphStyle('Body', fontName='Helvetica', fontSize=8.5, leading=11, textColor=colors.HexColor("#1E293B"))
        th_style = ParagraphStyle('TH', fontName='Helvetica-Bold', fontSize=7.5, textColor=colors.white, alignment=1)
        td_style = ParagraphStyle('TD', fontName='Helvetica', fontSize=7, leading=8.5, textColor=colors.HexColor("#0F172A"))
        
        story = [
            Paragraph("LICEO MODERNO WALT WHITMAN", title_style),
            Paragraph("Ecosistema Vital 26 • Formando Líderes Transformadores", sub_style),
            Spacer(1, 4)
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

def generate_docx(text):
    try:
        doc = Document()
        doc.add_heading("Liceo Moderno Walt Whitman", level=0)
        doc.add_paragraph("Ecosistema Vital 26 • Formación Institucional")
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

# --- ESTADOS DE SESIÓN ---
if "messages" not in st.session_state:
    st.session_state.messages = []
if "active_profile" not in st.session_state:
    st.session_state.active_profile = {
        "role_name": "Comunidad Educativa General",
        "badge": "VISITANTE GENERAL",
        "prompt_mode": "Asistente institucional del colegio para toda la comunidad."
    }

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
    st.header("🏫 Acceso Rápido")
    st.caption(f"Motor IA: `{active_model_name}`")
    if st.button("👑 Administrador (wq26)"):
        st.session_state.active_profile = PROFILES["wq26"]
        st.session_state.messages.append({"role": "assistant", "content": PROFILES["wq26"]["welcome"]})
        st.rerun()
    if st.button("🚀 SENA 2 (Grado 11°)"):
        st.session_state.active_profile = PROFILES["sena2"]
        st.session_state.messages.append({"role": "assistant", "content": PROFILES["sena2"]["welcome"]})
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
        
        # Mostrar imagen generada si la tiene
        if "image_url" in msg and msg["image_url"]:
            st.image(msg["image_url"], caption="🎨 Propuesta Visual Diseñada para el Evento", use_container_width=True)
            
        # Botones de exportación al final
        if msg["role"] == "assistant" and len(content_text) > 140:
            st.markdown("""
            <div style="margin-top: 15px; border-top: 1px solid #1e3a8a; padding-top: 8px;">
                <b style="color: #93c5fd; font-size: 0.85rem;">📥 Exportar y Guardar Documento:</b>
            </div>
            """, unsafe_allow_html=True)
            
            c1, c2, c3, c4 = st.columns(4)
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
            with c4:
                st.markdown('<a href="https://docs.new" target="_blank" class="docs-btn">🌐 Abrir en Google Docs</a>', unsafe_allow_html=True)

# --- ENTRADA DE MENSAJES ---
if prompt := st.chat_input("Escribe tu solicitud, redacta un comunicado, planeación o tarjeta..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # 1. Comprobar si es un login estricto
    matched_identity = check_user_identity(prompt)
    if matched_identity:
        st.session_state.active_profile = matched_identity
        w_msg = matched_identity["welcome"]
        with st.chat_message("assistant"):
            st.markdown(w_msg)
        st.session_state.messages.append({"role": "assistant", "content": w_msg})
        st.rerun()

    # 2. Es una solicitud de trabajo real
    else:
        with st.chat_message("assistant"):
            with st.spinner("Whitman AI está redactando y formateando tu documento..."):
                try:
                    p_lower = prompt.lower()
                    needs_img = any(k in p_lower for k in ["tarjeta", "invitacion", "invitación", "poster", "logo", "imagen", "diseño visual"])
                    
                    # Llamada DIRECTA al motor verificado de tu cuenta
                    full_query = f"Perfil: {st.session_state.active_profile['role_name']}\nDirectiva: {st.session_state.active_profile['prompt_mode']}\n\nSolicitud del usuario: {prompt}"
                    response = active_model.generate_content(full_query)
                    ai_text = response.text.strip()
                    
                    # Generar imagen visual si corresponde
                    img_link = None
                    if needs_img:
                        img_link = generate_image_url("Elegant institutional invitation card for student project defense at Villa Mayor Theater, navy blue and gold accents, official formal graphic design, high resolution")
                    
                    st.markdown(ai_text)
                    if img_link:
                        st.image(img_link, caption="🎨 Propuesta Visual Diseñada para el Evento", use_container_width=True)
                        
                    st.session_state.messages.append({
                        "role": "assistant",
                        "content": ai_text,
                        "image_url": img_link
                    })
                    st.rerun()
                except Exception as e:
                    err = f"Error al generar respuesta: {str(e)}"
                    st.error(err)
                    st.session_state.messages.append({"role": "assistant", "content": err})
