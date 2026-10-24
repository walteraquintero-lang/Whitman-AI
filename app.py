import streamlit as st
import google.generativeai as genai

st.set_page_config(
    page_title="Whitman AI - Liceo Moderno Walt Whitman",
    page_icon="🎓",
    layout="wide"
)

# --- CABECERA INSTITUCIONAL ---
st.title("🎓 Whitman AI")
st.subheader("Asistente Pedagógico Oficial - Liceo Moderno Walt Whitman")
st.caption("Herramienta de apoyo para docentes, planeación curricular, proyectos y comunidad educativa.")

# --- CONEXIÓN ---
api_key = st.secrets.get("GEMINI_API_KEY")
if not api_key:
    st.error("Falta configurar GEMINI_API_KEY en Secrets.")
    st.stop()

genai.configure(api_key=api_key)

# --- CEREBRO E INSTRUCCIONES DEL LICEO (PERSONALÍZALO AQUÍ) ---
SYSTEM_PROMPT = """
Eres WHITMAN AI, el asistente pedagógico e institucional oficial del LICEO MODERNO WALT WHITMAN.

TUS FUNCIONES PRINCIPALES:
1. APOYO A DOCENTES Y PLANEACIONES:
   - Ayudas a los profesores a estructurar planeaciones de clase, guías didácticas, proyectos de aula y evaluaciones.
   - Te riges por los estándares educativos, Derechos Básicos de Aprendizaje (DBA) y las orientaciones del Liceo.
   - Cuando te pidan una planeación, estructura siempre:
     * Tema y Estándar/DBA
     * Objetivo de aprendizaje
     * Inicio / Exploración (saberes previos)
     * Desarrollo / Conceptualización
     * Cierre / Evaluación formativa y retroalimentación

2. IDENTIDAD Y TONO:
   - Tu tono es pedagógico, colaborativo, empático, profesional e inspirador (haciendo honor al espíritu de Walt Whitman y la filosofía del liceo).
   - Siempre dispuesto a apoyar la labor docente y el aprendizaje integral de los estudiantes.
"""

# Configuración del modelo (utiliza el modelo disponible)
@st.cache_resource
def load_model():
    return genai.GenerativeModel(
        model_name="gemini-1.5-flash",
        system_instruction=SYSTEM_PROMPT
    )

try:
    model = load_model()
except Exception:
    model = genai.GenerativeModel(
        model_name="gemini-1.5-pro",
        system_instruction=SYSTEM_PROMPT
    )

# --- GESTIÓN DE SESIÓN Y CHAT ---
if "chat" not in st.session_state:
    st.session_state.chat = model.start_chat(history=[])
if "messages" not in st.session_state:
    st.session_state.messages = []

# Barra lateral para profesores
with st.sidebar:
    st.header("🏫 Liceo Moderno Walt Whitman")
    st.write("Bienvenido profesor/a. Whitman AI está listo para apoyarle en sus labores pedagógicas.")
    st.divider()
    if st.button("🗑️ Iniciar nueva consulta"):
        st.session_state.chat = model.start_chat(history=[])
        st.session_state.messages = []
        st.rerun()

# Mostrar mensajes anteriores
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# Entrada de mensaje
if prompt := st.chat_input("Escribe tu solicitud (ej: 'Ayúdame a planear una clase de ciencias para grado 5°')..."):
    # Mostrar mensaje usuario
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # Respuesta de Whitman AI
    with st.chat_message("assistant"):
        with st.spinner("Whitman AI está estructurando tu solicitud..."):
            try:
                response = st.session_state.chat.send_message(prompt)
                st.markdown(response.text)
                st.session_state.messages.append({"role": "assistant", "content": response.text})
            except Exception as e:
                st.error(f"Ocurrió un error al responder: {str(e)}")
