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
        padding: 20px 25px;
        margin-bottom: 20px;
        color: white;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.3);
    }
    .header-box h1 { margin: 0; font-size: 2rem; color: #ffffff; font-weight: 800; }
    .header-box p { margin: 4px 0 0 0; color: #93c5fd; font-size: 0.95rem; }
    .badge-role {
        display: inline-block;
        background-color: #f59e0b;
        color: #0f2042;
        font-weight: 800;
        padding: 4px 14px;
        border-radius: 20px;
        font-size: 0.8rem;
        margin-top: 10px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
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

# --- CONEXIÓN DIRECTA Y ROBUSTA A GEMINI ---
api_key = st.secrets.get("GEMINI_API_KEY")
if not api_key:
    st.error("Falta configurar GEMINI_API_KEY en Secrets de Streamlit.")
    st.stop()

genai.configure(api_key=api_key)

# Inicialización directa del modelo sin bloqueos de bucle
@st.cache_resource
def load_llm():
    # Modelos autorizados en orden de preferencia
    for m_name in ["gemini-1.5-flash", "gemini-2.0-flash", "gemini-1.5-pro", "gemini-pro"]:
        try:
            m = genai.GenerativeModel(m_name)
            return m_name, m
        except Exception:
            continue
    # Fallback directo
    return "gemini-1.5-flash", genai.GenerativeModel("gemini-1.5-flash")

model_name, active_model = load_llm()

# --- BASE DE CONOCIMIENTO INSTITUCIONAL MAESTRA ---
INSTITUTIONAL_CORE = r"""
ERES WHITMAN AI: EL ASISTENTE PEDAGÓGICO, ESTRATÉGICO Y MULTIMODAL DEL LICEO MODERNO WALT WHITMAN (BOGOTÁ, BARRIO VILLA MAYOR).
Entrenado por: Walter Quintero (Coordinador de Innovación y Tecnología).

1. IDENTIDAD INSTITUCIONAL:
- Lema: "El arte de la educación es el arte de despertar la curiosidad." — Walt Whitman
- Misión: Formar estudiantes que crezcan cada día, sean felices y emprendan sus sueños con habilidades para el siglo XXI.
- Ubicación / Aula Extendida: Calle 38 Bis Sur # 34D-29, Villa Mayor. Articulación con Parque Villa Mayor, Centro Comercial Centro Mayor y entorno barrial.

2. ECOSISTEMA VITAL 2026 (LOS 4 MOTORES ESTRATÉGICOS):
- Motor 1: Transformación Académica ("Tejiendo Saberes"): Enfoque Reggio Emilia (Cien Lenguajes de la infancia, el ambiente como Tercer Maestro), Educación Holística y Aprendizaje Basado en Proyectos (ABP).
- Motor 2: Bienestar Integral ("Cátedra de Emocionalidad"): Ley 2503/2025. Malla curricular por ciclos (Ciclo 1: Reconocimiento/Calma; Ciclo 2: Frustración/Asertividad; Ciclo 3: Riesgos/Decisiones; Ciclo 4: Liderazgo/Proyecto de Vida). Justicia Restaurativa y Rincones de Calma.
- Motor 3: Ciudadanía Global (ODS y Psicoorientación): Conexión con Agenda 2030 (ODS 2, 6, 12, 16). Inmersión bilingüe y proyecto de vida STEM/Verde (ODS 8).
- Motor 4: Innovación Tecnológica (IA como Atelier Digital): La IA no es sustituto sino copiloto y socio metodológico. Metodología Sándwich (Pan humano - Jamón IA - Pan humano).

3. PROYECTO DE BILINGÜISMO (CIUDADANOS DEL MUNDO):
- Liderado por Marcela Santana y Natalia Cruz. Inmersión internacional en Toronto, Canadá.
- Pruebas internacionales TOEFL Junior (Nivel B1/B2 en 11°). Singing Fest y Spelling Bee.
- En Science (1° a 5°) y Ciudadanos del Mundo: 100% en inglés técnico (Technical Science English), metodología 70/30 (30% Explicación, 70% Speaking/Listening) e integración de ODS.

4. ARTICULACIÓN MEDIA TÉCNICA SENA Y WHITMAN STUDIO:
- Programa: Técnico en Integración de Contenidos Digitales (Grados 10° y 11°). Doble titulación.
- 9 Fases del Proyecto Productivo (Reto 5: Investigación de Mercados, Buyer Persona, etc.).
- Estándar Walter Quintero: Proyectos vectoriales estrictos en Adobe Illustrator (.AI) en 3 capas ([Fondos], [Vectores], [Textos]), tipografías a curvas, sin bitmaps pixelados.
- Whitman Studio: Semillero creativo de producción audiovisual, branding, hooks virales y fotografía escolar.

5. SIEE (EVALUACIÓN) Y MANUAL DE CONVIVENCIA:
- Evaluación cualitativa: Superior, Alto, Básico y Bajo (Cero promedios numéricos).
- La sustentación oral vale el 60% de la evaluación; el entregable escrito el 40%.
- Faltas Tipo I (Leves), Tipo II (Graves) y Tipo III (Gravísimas) con estricto debido proceso.
- Póliza de accidentes escolares: SURA (Código convenio: 2005341).

6. DIRECTIVAS DE GENERACIÓN MULTIMODAL:
- Si el usuario solicita un MAPA, DIAGRAMA O FLUJO: Genera un bloque de código Mermaid (```mermaid ... ```) para que se renderice visualmente.
- Si solicita AUDIO O SONIDO: Diseña el guion técnico con efectos de sonido (SFX), tiempos, y prompts musicales listos para Suno o Udio.
- Si solicita COMUNICADOS O TARJETAS: Redáctalos con toda la formalidad institucional, orden del día, fecha, lugar (Teatro Villa Mayor) y firmas correspondientes.
- Si solicita PLANEACIONES: Estructura encabezado, DBA, momentos de clase de 40-45 min, páginas del libro y módulo PIAR.
- NUNCA generes notas preliminares en inglés ("Target audience", "Context", "Self-correction"). Entrega directamente el producto solicitado.
"""

# --- PARSEADOR DE DIAGRAMAS MERMAID ---
def render_mermaid(text):
    mermaid_matches = re.findall(r"```mermaid\n([\s\S]*?)\n```", text)
    for code in mermaid_matches:
        html_code = f"""
        <div class="mermaid" style="background:#0f2042; padding:15px; border-radius:10px; border:1px solid #f59e0b;">
            {code}
        </div>
        <script src="https://cdn.jsdelivr.net/npm/mermaid/dist/mermaid.min.js"></script>
        <script>mermaid.initialize({{startOnLoad:true, theme:'dark'}});</script>
        """
        components.html(html_code, height=350, scrolling=True)

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
        title_slide.shapes.title.text = "Whitman AI - Material Institucional"
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
if "user_role" not in st.session_state:
    st.session_state.user_role = "Comunidad Educativa General"

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
    st.header("🏫 Centro de Mando")
    st.caption(f"Motor Conectado: `{model_name}`")
    st.markdown("### Selecciona o valida tu rol:")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("👑 Walter Q."):
            st.session_state.user_role = "Coordinador de Innovación y Tecnología"
            st.session_state.messages.append({"role": "assistant", "content": "¡Bienvenido, Walter! Entorno de Administración, Whitman Studio y Ecosistema Vital activo."})
            st.rerun()
        if st.button("🚀 SENA 11°"):
            st.session_state.user_role = "Aprendiz SENA - Fase 2 (11°)"
            st.session_state.messages.append({"role": "assistant", "content": "¡Bienvenido, Aprendiz SENA! Laboratorio de Proyectos Productivos, Retos y Adobe Illustrator listo."})
            st.rerun()
    with c2:
        if st.button("🔬 Science"):
            st.session_state.user_role = "Docente Science / Ciudadanos del Mundo"
            st.session_state.messages.append({"role": "assistant", "content": "¡Welcome Teacher! Formato bilingüe 70/30, ODS y estándares MCER activos."})
            st.rerun()
        if st.button("🏡 Familias"):
            st.session_state.user_role = "Padres de Familia"
            st.session_state.messages.append({"role": "assistant", "content": "¡Bienvenidos Padres y Madres Liceístas! Espacio de orientación, convivencia y eventos escolares."})
            st.rerun()
    st.divider()
    if st.button("🗑️ Nueva Conversación"):
        st.session_state.messages = []
        st.session_state.user_role = "Comunidad Educativa General"
        st.rerun()

# --- HISTORIAL DE CONVERSACIÓN ---
for idx, msg in enumerate(st.session_state.messages):
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])
        
        # Renderizar diagramas Mermaid si los hay
        if "```mermaid" in msg["content"]:
            render_mermaid(msg["content"])
            
        # Botones de exportación al final
        if msg["role"] == "assistant" and len(msg["content"]) > 130:
            st.markdown("""
            <div style="margin-top: 15px; border-top: 1px solid #1e3a8a; padding-top: 8px;">
                <b style="color: #93c5fd; font-size: 0.85rem;">📥 Exportar a Formatos Oficiales:</b>
            </div>
            """, unsafe_allow_html=True)
            
            b1, b2, b3, b4 = st.columns(4)
            docx_data = generate_docx(msg["content"])
            if docx_data:
                with b1:
                    st.download_button("📄 Word (.docx)", docx_data, f"Whitman_Doc_{idx}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"docx_{idx}")
            pdf_data = generate_pdf_reportlab(msg["content"])
            if pdf_data:
                with b2:
                    st.download_button("📑 PDF Oficial", pdf_data, f"Whitman_Doc_{idx}.pdf", "application/pdf", key=f"pdf_{idx}")
            pptx_data = generate_pptx(msg["content"])
            if pptx_data:
                with b3:
                    st.download_button("📊 Diapositivas (.pptx)", pptx_data, f"Whitman_Doc_{idx}.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", key=f"pptx_{idx}")
            with b4:
                st.markdown('<a href="https://docs.new" target="_blank" class="docs-btn">🌐 Abrir en Google Docs</a>', unsafe_allow_html=True)

# --- CAJA DE ENTRADA UNIVERSAL ---
if prompt := st.chat_input("Escribe cualquier solicitud (planeaciones, mapas, circulares, guías SENA, etc.)..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        with st.spinner("Whitman AI está estructurando tu solicitud con los estándares del Liceo..."):
            try:
                # Consulta contextualizada con la base de conocimiento institucional completa
                query = f"""
{INSTITUTIONAL_CORE}

ROL DEL USUARIO ACTUAL: {st.session_state.user_role}

SOLICITUD:
{prompt}
"""
                response = active_model.generate_content(query)
                answer_text = response.text.strip()
                
                # Limpieza de notas internas si hubiese
                if "User input:" in answer_text or "*Target Audience:" in answer_text:
                    cut = answer_text.rfind("---")
                    if cut != -1:
                        answer_text = answer_text[cut:].strip()
                
                st.markdown(answer_text)
                
                # Renderizar diagrama si la respuesta generó uno
                if "```mermaid" in answer_text:
                    render_mermaid(answer_text)
                    
                st.session_state.messages.append({"role": "assistant", "content": answer_text})
                st.rerun()
            except Exception as e:
                err_msg = f"Error al generar respuesta: {str(e)}"
                st.error(err_msg)
                st.session_state.messages.append({"role": "assistant", "content": err_msg})
