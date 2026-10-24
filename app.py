import streamlit as st
import google.generativeai as genai
import io
import re
from docx import Document
from pptx import Presentation
from fpdf import FPDF

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(
    page_title="Whitman AI - Ecosistema Vital 26",
    page_icon="🎓",
    layout="wide"
)

# --- FILTRO ANTIBOCETOS / SALIDA LIMPIA ---
def clean_ai_response(text):
    """Elimina notas preliminares o borradores en inglés que la IA genere al inicio"""
    # Si detecta encabezados de borrador interno, buscar dónde inicia la guía real
    patterns = [r"\n(?=#[^#])", r"\n(?=##[^#])", r"\n(?=📖)", r"\n(?=✅)", r"\n(?=📋)"]
    for pat in patterns:
        match = re.search(pat, text)
        if match and match.start() < 600:
            return text[match.start():].strip()
    return text.strip()

# --- GENERADORES SEGUROS DE ARCHIVOS (CON PROTECCIÓN DE ERRORES) ---
def generate_docx(text):
    try:
        doc = Document()
        doc.add_heading("Liceo Moderno Walt Whitman", level=0)
        doc.add_paragraph("Ecosistema Vital 26 • Generado por Whitman AI")
        doc.add_paragraph("--------------------------------------------------")
        
        for line in text.split("\n"):
            clean = line.strip()
            if clean.startswith("### "):
                doc.add_heading(clean.replace("### ", ""), level=2)
            elif clean.startswith("## "):
                doc.add_heading(clean.replace("## ", ""), level=1)
            elif clean.startswith("# "):
                doc.add_heading(clean.replace("# ", ""), level=1)
            elif clean.startswith("- ") or clean.startswith("* "):
                doc.add_paragraph(clean[2:], style='List Bullet')
            elif clean:
                doc.add_paragraph(clean)
                
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    except Exception:
        return None

def generate_pdf(text):
    try:
        # Remover emojis y caracteres especiales que rompen fuentes básicas de PDF
        clean_text = re.sub(r'[^\x00-\x7F\xC0-\xFF]+', ' ', text)
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Helvetica", "B", 13)
        pdf.cell(0, 10, "Liceo Moderno Walt Whitman - Whitman AI", align="C")
        pdf.ln(10)
        pdf.set_font("Helvetica", size=9)
        
        for line in clean_text.split("\n"):
            l = line.strip()
            if not l:
                pdf.ln(2)
                continue
            if l.startswith("#"):
                pdf.set_font("Helvetica", "B", 10)
                pdf.multi_cell(0, 6, l.replace("#", "").strip()[:100])
                pdf.set_font("Helvetica", size=9)
            else:
                pdf.multi_cell(0, 5, l[:150])
        return bytes(pdf.output())
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

# --- CABECERA INSTITUCIONAL ---
st.title("🎓 Whitman AI - v2.5")
st.subheader("Asistente Pedagógico Institucional | Liceo Moderno Walt Whitman")
st.caption("Ecosistema Vital 26 • Entrenado por Walter Quintero (Coord. Innovación y Tecnología)")

# --- CONEXIÓN CON GEMINI API ---
api_key = st.secrets.get("GEMINI_API_KEY")
if not api_key:
    st.error("Falta configurar la clave GEMINI_API_KEY en Secrets de Streamlit.")
    st.stop()

genai.configure(api_key=api_key)

# --- CEREBRO MAESTRO DE WHITMAN AI ---
SYSTEM_PROMPT = r"""
CONFIGURACIÓN AVANZADA PARA GEMA WHITMAN AI - V2 (ECOSISTEMA VITAL 26)

[DIRECTIVA ESTRICTA DE SALIDA LIMPIA]:
- Queda TOTALMENTE PROHIBIDO redactar o mostrar pensamientos preliminares, resúmenes en inglés, "* Role:", "* Step 1:", o análisis previos.
- Genera DIRECTAMENTE el producto final, comenzando de inmediato por el título oficial o la validación del perfil.
- Responde siempre en español (excepto en las planeaciones de Science / English que son 100% en inglés según el protocolo).

IDENTIDAD Y ROL:
Eres Whitman AI, asistente pedagógico del Liceo Moderno Walt Whitman.
Diseñado y entrenado por: Walter Quintero (Coordinador de Innovación y Tecnología).

[VALIDACIÓN DE ROL]:
- Administrador: wq26 -> Walter Quintero / Coordinador de Tecnología e Innovación
- Dirección: ILGACHA26, LFMENDEZ26, CMENDEZ26, LKMENDEZ26, DFMENDEZ26
- Coordinaciones: NAACOSTA26, ISALGADO26, LMCUBIDES26, AMCAMARGO26, JCPRADA26
- Docentes: SLATORRE26, NATAC26, MSANTANA26, TPARRA26, SBUSTOS26, SPINEROS26 y demás docentes de la matriz.
- SENA y Studio: "sena1", "sena2", "WS26_STUDIO".
- Estudiantes: "soy sexto", "soy noveno", "soy decimo", "soy once", etc.

Respuesta de validación para Docente:
"✅ **Perfil Validado:** [Nombre del Docente / Cargo - Nodo]
¡Bienvenido/a a Whitman AI! Tu entorno pedagógico está activo. Puedes ingresar tus comandos de planeación o indicarme el tema de hoy."

Respuesta para Estudiantes / SENA:
"✅ **Perfil Validado:** [Nombre del Rol / Nivel]
Bienvenido/a a Whitman AI. Tu espacio de trabajo está configurado para las guías y actividades de tu nivel. ¿En qué proyecto o entregable vamos a trabajar hoy?"

FORMATOS Y CIERRE:
Toda respuesta debe finalizar con:
---
> *"El arte de la educación es el arte de despertar la curiosidad."* — Walt Whitman

**Resultados potenciados por Whitman AI | Entrenado por: Walter Quintero (Coord. Innovación y Tecnología)**
"""

# --- MOTOR DE IA ---
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
    raise RuntimeError("No se encontró ningún modelo compatible habilitado en la API Key.")

active_model_name, model = get_working_engine()

# --- MEMORIA Y SESIÓN ---
if "chat" not in st.session_state:
    st.session_state.chat = model.start_chat(history=[])
if "messages" not in st.session_state:
    st.session_state.messages = []

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("🏫 Centro de Mando")
    st.markdown("**Liceo Moderno Walt Whitman**")
    st.caption(f"Motor activo: `{active_model_name}`")
    
    st.markdown("### 🔐 Acceso Seguro Docente")
    secret_code = st.text_input("Código institucional:", type="password")
    if st.button("Validar Credencial"):
        if secret_code.strip():
            st.session_state.messages.append({"role": "user", "content": "Código validado por el usuario"})
            with st.spinner("Autenticando..."):
                resp = st.session_state.chat.send_message(secret_code.strip())
                st.session_state.messages.append({"role": "assistant", "content": clean_ai_response(resp.text)})
            st.rerun()

    st.divider()
    if st.button("🗑️ Reiniciar Sesión / Nueva Consulta"):
        st.session_state.chat = model.start_chat(history=[])
        st.session_state.messages = []
        st.rerun()

# --- MOSTRAR CONVERSACIÓN Y DESCARGAS ---
for idx, msg in enumerate(st.session_state.messages):
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])
        
        # Botones de exportación protegidos
        if msg["role"] == "assistant" and len(msg["content"]) > 100:
            content_text = msg["content"]
            col1, col2, col3 = st.columns(3)
            
            docx_data = generate_docx(content_text)
            if docx_data:
                with col1:
                    st.download_button(
                        label="📄 Descargar en Word (.docx)",
                        data=docx_data,
                        file_name=f"Whitman_Guia_{idx}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"docx_{idx}"
                    )
                    
            pdf_data = generate_pdf(content_text)
            if pdf_data:
                with col2:
                    st.download_button(
                        label="📑 Descargar en PDF",
                        data=pdf_data,
                        file_name=f"Whitman_Documento_{idx}.pdf",
                        mime="application/pdf",
                        key=f"pdf_{idx}"
                    )
                    
            pptx_data = generate_pptx(content_text)
            if pptx_data:
                with col3:
                    st.download_button(
                        label="📊 Descargar diapositivas (.pptx)",
                        data=pptx_data,
                        file_name=f"Whitman_Presentacion_{idx}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key=f"pptx_{idx}"
                    )

# --- CAJA DE ENTRADA ---
if prompt := st.chat_input("Escribe tu solicitud, comando de planeación o pide un recurso visual..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        with st.spinner("Whitman AI está elaborando el contenido..."):
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
