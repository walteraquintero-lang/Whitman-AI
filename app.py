import streamlit as st
import google.generativeai as genai

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(
    page_title="Whitman AI - Asistente Escolar",
    page_icon="🎓",
    layout="centered"
)

st.title("🎓 Whitman AI")
st.caption("Tu asistente escolar de cabecera. ¡Pregúntame sobre tus clases, proyectos o tareas!")

# --- CONEXIÓN CON LA IA ---
api_key = st.secrets.get("GEMINI_API_KEY")

if not api_key:
    st.error("Falta configurar la clave GEMINI_API_KEY en los secretos.")
    st.stop()

genai.configure(api_key=api_key)

# Personalidad de Whitman AI
SYSTEM_PROMPT = """
Eres Whitman AI, el asistente virtual oficial del colegio.
- Tu tono es entusiasta, pedagógico, respetuoso y muy claro.
- Tu objetivo es guiar a los estudiantes para que aprendan y piensen por sí mismos, no solo darles las respuestas directas.
- Ayuda a profesores y padres con dudas sobre actividades y proyectos escolares.
- Mantén siempre un entorno seguro para menores.
"""

# Usamos el mode
