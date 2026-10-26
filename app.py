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

# --- ESTILOS INSTITUCIONALES ---
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
    .docs-btn:hover { background-color: #2563eb; }
</style>
""", unsafe_allow_html=True)

# --- CONEXIÓN A GOOGLE ---
api_key = st.secrets.get("GEMINI_API_KEY")
if not api_key:
    st.error("Falta configurar GEMINI_API_KEY en Secrets.")
    st.stop()

genai.configure(api_key=api_key)

# --- DETECCIÓN DEL MOTOR 3.8 / MODERNO ---
@st.cache_resource
def get_clean_model():
    candidates = ["models/gemini-3.8-flash", "gemini-3.8-flash", "models/gemini-2.0-flash", "gemini-2.0-flash"]
    try:
        for m in genai.list_models():
            if "generateContent" in m.supported_generation_methods and m.name not in candidates:
                candidates.append(m.name)
    except Exception:
        pass
        
    for mod_name in candidates:
        try:
            m = genai.GenerativeModel(mod_name)
            m.generate_content("ping")
            return mod_name, m
        except Exception:
            continue
    raise RuntimeError("No se pudo iniciar el modelo de Google.")

model_name, active_model = get_clean_model()

# --- GENERADOR DE IMÁGENES OFICIAL GOOGLE IMAGEN 3 ---
def generate_google_image(prompt_text):
    try:
        imagen = genai.ImageGenerationModel("imagen-3.0-generate-002")
        result = imagen.generate_images(
            prompt=prompt_text,
            number_of_images=1,
            aspect_ratio="16:9"
        )
        for img in result.images:
            return img._image_bytes
    except Exception:
        return None

# --- GENERADOR DE TARJETA VISUAL INSTITUCIONAL (HTML/CSS) ---
def render_invitation_card(title, subtitle, date_info, location, protocol):
    return f"""
    <div style="background: linear-gradient(135deg, #0f2042 0%, #1e3a8a 100%); 
                border: 3px solid #f59e0b; border-radius: 12px; padding: 24px; 
                color: white; text-align: center; max-width: 680px; margin: 15px auto; 
                box-shadow: 0 8px 25px rgba(0,0,0,0.4);">
        <p style="color: #f59e0b; font-weight: 800; font-size: 0.85rem; letter-spacing: 2px; margin: 0;">LICEO MODERNO WALT WHITMAN</p>
        <p style="color: #93c5fd; font-size: 0.8rem; margin: 2px 0 12px 0;">ECOSISTEMA VITAL 26 • PROYECTO PRODUCTIVO SENA</p>
        <hr style="border: 0; border-top: 1px solid rgba(245, 158, 11, 0.4); margin: 10px 0;">
        <h2 style="color: #ffffff; font-size: 1.4rem; margin: 8px 0;">{title}</h2>
        <p style="color: #e2e8f0; font-size: 0.95rem; margin-bottom: 15px;">{subtitle}</p>
        <div style="background: rgba(15, 32, 66, 0.7); border-radius: 8px; padding: 12px; margin: 10px 0; border: 1px solid rgba(59, 130, 246, 0.4); text-align: left;">
            <p style="margin: 4px 0; font-size: 0.85rem;">📅 <b>Fecha:</b> {date_info}</p>
            <p style="margin: 4px 0; font-size: 0.85rem;">📍 <b>Lugar:</b> {location}</p>
            <p style="margin: 4px 0; font-size: 0.85rem;">👔 <b>Protocolo:</b> {protocol}</p>
        </div>
        <p style="color: #f59e0b; font-size: 0.75rem; margin-top: 10px; font-style: italic;">
            "El arte de la educación es el arte de despertar la curiosidad." — Walt Whitman
        </p>
    </div>
    """

# --- GENERADORES DE ARCHIVOS ---
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
if "active_role" not in st.session_state:
    st.session_state.active_role = "ADMINISTRACIÓN / INNOVACIÓN"

# --- HEADER PRINCIPAL ---
st.markdown(f"""
<div class="header-box">
    <h1>🎓 Whitman AI • Ecosistema Vital 26</h1>
    <p>Liceo Moderno Walt Whitman | Asistente Integral Escolar</p>
    <div class="badge-role">{st.session_state.active_role}</div>
</div>
""", unsafe_allow_html=True)

# --- SIDEBAR ---
with st.sidebar:
    st.header("🏫 Centro de Control")
    st.caption(f"Motor activo: `{model_name}`")
    st.divider()
    if st.button("🗑️ Reiniciar Sesión"):
        st.session_state.messages = []
        st.rerun()

# --- MOSTRAR CONVERSACIÓN ---
for idx, msg in enumerate(st.session_state.messages):
    with st.chat_message(msg["role"]):
        content_text = msg["content"]
        st.markdown(content_text)
        
        # Tarjeta visual maquetada si aplica
        if "card_html" in msg and msg["card_html"]:
            st.markdown(msg["card_html"], unsafe_allow_html=True)
            
        # Imagen binaria de Google Imagen 3 si aplica
        if "img_bytes" in msg and msg["img_bytes"]:
            st.image(msg["img_bytes"], caption="🎨 Ilustración Generada por Google Imagen 3", use_container_width=True)
            
        # Botones de exportación al final
        if msg["role"] == "assistant" and len(content_text) > 120:
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

    with st.chat_message("assistant"):
        with st.spinner("Whitman AI está estructurando tu documento oficial..."):
            try:
                master_query = f"""
Eres Whitman AI, asistente pedagógico e institucional oficial del Liceo Moderno Walt Whitman (Villa Mayor, Bogotá).
Entrenado por Walter Quintero (Coord. Innovación y Tecnología).

DIRECTIVAS INSTITUCIONALES:
- Responde DIRECTAMENTE con el documento o diseño solicitado. Cero notas de pensamiento preliminares.
- Si te piden una tarjeta o comunicado para las sustentaciones, incluye:
  * Fecha: 29 de Octubre / Noviembre
  * Lugar: Teatro Villa Mayor
  * Asunto: Sustentación de Proyectos Productivos SENA Grado 11°
  * Orden del día / Agenda en formato de tabla
  * Protocolo formal de asistencia
- Finaliza siempre con:
---
> "El arte de la educación es el arte de despertar la curiosidad." — Walt Whitman
Resultados potenciados por Whitman AI | Entrenado por: Walter Quintero (Coord. Innovación y Tecnología)

SOLICITUD:
{prompt}
"""
                res = active_model.generate_content(master_query)
                ai_text = res.text.strip()
                
                # Comprobar si requiere elementos visuales
                p_low = prompt.lower()
                is_invitation = any(k in p_low for k in ["tarjeta", "invitacion", "invitación"])
                wants_artwork = any(k in p_low for k in ["imagen", "ilustracion", "dibujo", "poster"])
                
                card_html = None
                img_bytes = None
                
                # 1. Renderizar tarjeta gráfica si es invitación
                if is_invitation:
                    card_html = render_invitation_card(
                        title="SUSTENTACIÓN DE PROYECTOS PRODUCTIVOS SENA",
                        subtitle="Promoción Grado 11° • Ecosistema Vital 26",
                        date_info="29 de Octubre | 8:00 AM - 1:00 PM",
                        location="Teatro Villa Mayor (Bogotá D.C.)",
                        protocol="Traje Formal / Uniforme de Gala"
                    )
                
                # 2. Generar imagen real con Google Imagen 3 si pidieron arte gráfico
                if wants_artwork:
                    img_bytes = generate_google_image("Official luxury invitation poster for student project defense at theater, navy blue and gold accents, minimal clean vector layout, 8k")
                
                st.markdown(ai_text)
                if card_html:
                    st.markdown(card_html, unsafe_allow_html=True)
                if img_bytes:
                    st.image(img_bytes, caption="🎨 Ilustración Generada por Google Imagen 3", use_container_width=True)
                    
                st.session_state.messages.append({
                    "role": "assistant",
                    "content": ai_text,
                    "card_html": card_html,
                    "img_bytes": img_bytes
                })
                st.rerun()
            except Exception as e:
                st.error(f"Error al procesar: {str(e)}")
