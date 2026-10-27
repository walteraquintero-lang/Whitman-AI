import streamlit as st
import google.generativeai as genai
import io
import re
import requests
import base64
from docx import Document
from pptx import Presentation
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import streamlit.components.v1 as components

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
    st.error("Falta configurar GEMINI_API_KEY en Secrets.")
    st.stop()

genai.configure(api_key=api_key)

# Cargar modelo activo con validación
@st.cache_resource
def load_llm():
    for name in ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro", "gemini-pro"]:
        try:
            m = genai.GenerativeModel(name)
            m.generate_content("ok")
            return name, m
        except Exception:
            continue
    return "gemini-1.5-flash", genai.GenerativeModel("gemini-1.5-flash")

model_name, active_model = load_llm()

# --- GENERADOR NATIVO DE IMÁGENES GOOGLE (SIN MANDAR A DALL-E) ---
def generate_image_google(prompt_text, key):
    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/imagen-3.0-generate-002:predict?key={key}"
        payload = {
            "instances": [{"prompt": f"{prompt_text}, high quality graphic design, clean institutional colors navy blue and gold"}],
            "parameters": {"sampleCount": 1, "aspectRatio": "16:9"}
        }
        res = requests.post(url, json=payload, timeout=20)
        if res.status_code == 200:
            data = res.json()
            b64 = data["predictions"][0]["bytesBase64Encoded"]
            return base64.b64decode(b64)
    except Exception:
        pass
    return None

# --- PARSEADOR MERMAID CON PROTECCIÓN ANTI-ERRORES ---
def render_mermaid_safe(text):
    matches = re.findall(r"```mermaid\n([\s\S]*?)\n```", text)
    for code in matches:
        # Sanitizar código
        safe_code = code.replace('"', "'")
        html = f"""
        <div id="mermaid-wrap" style="background:#0f2042; padding:12px; border-radius:8px; border:1px solid #f59e0b; margin:10px 0;">
            <div class="mermaid">{safe_code}</div>
        </div>
        <script src="https://cdn.jsdelivr.net/npm/mermaid/dist/mermaid.min.js"></script>
        <script>
            try {{
                mermaid.initialize({{startOnLoad:true, theme:'dark', suppressErrorRendering:true}});
            }} catch(e) {{
                document.getElementById('mermaid-wrap').style.display = 'none';
            }}
        </script>
        """
        components.html(html, height=320, scrolling=True)

# --- GENERADORES DE ARCHIVOS (WORD, PDF, PPT) ---
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
        
        clean_text = re.sub(r'```mermaid[\s\S]*?```', '', text)
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
            elif clean and not clean.startswith("|") and not clean.startswith("```"):
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
            lines = [l.strip() for l in sec.split("\n") if l.strip() and not l.startswith("|") and not l.startswith("```")]
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

# --- PERFILES CON BIENVENIDAS CORTAS Y DIRECTAS ---
PROFILES = {
    "wq26": {
        "role_name": "👑 Coordinador Walter Quintero",
        "badge": "ADMINISTRACIÓN / INNOVACIÓN",
        "welcome": "¡Hola Walter! Centro de Control activo. ¿Qué planeación, circular o proyecto de Whitman Studio estructuramos hoy?"
    },
    "docente": {
        "role_name": "👩‍🏫 Docente Liceísta",
        "badge": "DOCENTE LICEÍSTA",
        "welcome": """¡Bienvenido/a profe! Tu entorno de planeación está listo.
Recuerda que puedes usar la estructura rápida:
👉 `planear/[grado]/[área]/[tema]/[sesiones]/`
O simplemente decirme: *"Ayúdame a planear una clase de [tema] para grado [X]"*. ¿En qué te colaboro hoy?"""
    },
    "sena2": {
        "role_name": "🚀 Aprendiz SENA (11°)",
        "badge": "APRENDIZ SENA - FASE 2",
        "welcome": "¡Hola Aprendiz! Laboratorio de Illustrator y Proyecto Productivo listo. ¿Trabajamos en tu Buyer Persona, matrices del Reto 5 o una pieza gráfica?"
    },
    "padre": {
        "role_name": "🏡 Familia Whitman",
        "badge": "PADRES DE FAMILIA",
        "welcome": "¡Bienvenidos Padres y Madres Liceístas! Espacio de orientación escolar. ¿En qué información sobre eventos, horarios o vida escolar les puedo colaborar?"
    }
}

def check_identity(text):
    clean = text.lower().strip()
    action_words = ["crea", "creemos", "haz", "diseña", "redacta", "planea", "escribe", "genera", "tarjeta", "invitacion", "invitación", "comunicado", "imagen"]
    if any(w in clean for w in action_words) or len(clean.split()) > 4:
        return None
    if "wq26" in clean:
        return PROFILES["wq26"]
    elif any(k in clean for k in ["sena 2", "sena2", "soy sena 2"]):
        return PROFILES["sena2"]
    elif any(k in clean for k in ["profe", "docente", "profesor", "profesora"]):
        return PROFILES["docente"]
    elif clean in ["padre", "soy padre", "acudiente", "mama", "papa"]:
        return PROFILES["padre"]
    return None

# --- ESTADOS DE SESIÓN ---
if "messages" not in st.session_state:
    st.session_state.messages = []
if "user_role" not in st.session_state:
    st.session_state.user_role = "Comunidad Educativa General"

# --- CABECERA PRINCIPAL ---
st.markdown(f"""
<div class="header-box">
    <h1>🎓 Whitman AI • Ecosistema Vital 26</h1>
    <p>Liceo Moderno Walt Whitman | Asistente Integral Pedagógico y Multimodal</p>
    <div class="badge-role">{st.session_state.user_role}</div>
</div>
""", unsafe_allow_html=True)

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("🏫 Acceso Directo")
    st.caption(f"Motor Conectado: `{model_name}`")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("👑 Walter Q."):
            st.session_state.user_role = PROFILES["wq26"]["badge"]
            st.session_state.messages.append({"role": "assistant", "content": PROFILES["wq26"]["welcome"]})
            st.rerun()
        if st.button("🚀 SENA 11°"):
            st.session_state.user_role = PROFILES["sena2"]["badge"]
            st.session_state.messages.append({"role": "assistant", "content": PROFILES["sena2"]["welcome"]})
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
        st.session_state.user_role = "Comunidad Educativa General"
        st.rerun()

# --- HISTORIAL ---
for idx, msg in enumerate(st.session_state.messages):
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])
        
        # Renderizar diagrama si lo hay
        if "```mermaid" in msg["content"]:
            render_mermaid_safe(msg["content"])
            
        # Mostrar imagen generada directamente si existe
        if "image_bytes" in msg and msg["image_bytes"]:
            st.image(msg["image_bytes"], caption="🎨 Propuesta Gráfica Generada", use_container_width=True)
            
        # Botones de exportación
        if msg["role"] == "assistant" and len(msg["content"]) > 100:
            st.markdown("""
            <div style="margin-top: 15px; border-top: 1px solid #1e3a8a; padding-top: 8px;">
                <b style="color: #93c5fd; font-size: 0.85rem;">📥 Exportar Documento:</b>
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
if prompt := st.chat_input("Escribe tu solicitud (planeaciones, comunicados, imágenes, mapas conceptuales...)..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # 1. Comprobar si es un login breve
    ident = check_identity(prompt)
    if ident:
        st.session_state.user_role = ident["badge"]
        w = ident["welcome"]
        with st.chat_message("assistant"):
            st.markdown(w)
        st.session_state.messages.append({"role": "assistant", "content": w})
        st.rerun()

    # 2. Solicitud directa al motor
    else:
        with st.chat_message("assistant"):
            with st.spinner("Whitman AI está procesando..."):
                try:
                    p_low = prompt.lower()
                    wants_image = any(k in p_low for k in ["imagen", "logo", "tarjeta visual", "afiche", "poster", "dibujo"])
                    
                    master_query = f"""
Eres Whitman AI, asistente pedagógico e institucional oficial del Liceo Moderno Walt Whitman (Bogotá, Villa Mayor).
Entrenado por Walter Quintero (Coord. Innovación y Tecnología).
Rol actual: {st.session_state.user_role}

DIRECTIVAS ESTRICTAS:
1. Responde DIRECTAMENTE con la tarea solicitada en español (o inglés si es Science/Bilingüe).
2. Si piden imágenes o diseños: NUNCA digas "ve a DALL-E" o "usa Midjourney". Describe la propuesta gráfica y entrega el contenido listo. La aplicación generará la imagen automáticamente.
3. Si piden una planeación: usa la estructura oficial del Liceo (DBA, 40-45 min, páginas del libro, PIAR).
4. Si piden un diagrama: genera un bloque ```mermaid ... ``` con sintaxis limpia y sencilla.
5. Cero notas de pensamiento ("Target audience:", "Context:").

Finaliza siempre con:
---
> "El arte de la educación es el arte de despertar la curiosidad." — Walt Whitman
Resultados potenciados por Whitman AI | Entrenado por: Walter Quintero (Coord. Innovación y Tecnología)

SOLICITUD:
{prompt}
"""
                    res = active_model.generate_content(master_query)
                    answer_text = res.text.strip()
                    
                    # Generar imagen directa si se solicitó
                    img_data = None
                    if wants_image:
                        img_data = generate_image_google(prompt, api_key)
                    
                    st.markdown(answer_text)
                    if "```mermaid" in answer_text:
                        render_mermaid_safe(answer_text)
                    if img_data:
                        st.image(img_data, caption="🎨 Propuesta Visual Generada en Vivo", use_container_width=True)
                        
                    st.session_state.messages.append({
                        "role": "assistant",
                        "content": answer_text,
                        "image_bytes": img_data
                    })
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al generar respuesta: {str(e)}")
