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

SYSTEM_PROMPT = """
Eres Whitman AI, el asistente virtual oficial del colegio.
- Tu tono es entusiasta, pedagógico, respetuoso y muy claro.
- Tu objetivo es guiar a los estudiantes para que aprendan y piensen por sí mismos, no solo darles las respuestas directas.
- Ayuda a profesores y padres con dudas sobre actividades y proyectos escolares.
- Mantén siempre un entorno seguro para menores.
"""

# Detectar y asignar el modelo estable de Google
@st.cache_resource
def get_model():
    try:
        available = [
            m.name for m in genai.list_models()
            if "generateContent" in m.supported_generation_methods
        ]
        # Priorizar gemini-2.0-flash, gemini-1.5-pro o el primero disponible
        for preferred in ["gemini-2.0-flash", "gemini-1.5-pro", "gemini-pro"]:
            found = next((m for m in available if preferred in m), None)
            if found:
                return found, genai.GenerativeModel(model_name=found, system_instruction=SYSTEM_PROMPT)
        # Si no, tomar el primero de la lista
        first_model = available[0]
        return first_model, genai.GenerativeModel(model_name=first_model, system_instruction=SYSTEM_PROMPT)
    except Exception as e:
        # Fallback estándar
        return "gemini-1.5-pro", genai.GenerativeModel(model_name="gemini-1.5-pro", system_instruction=SYSTEM_PROMPT)

model_name, model = get_model()

# --- HISTORIAL DE CONVERSACIÓN ---
if "messages" not in st.session_state:
    st.session_state.messages = []

# Barra lateral para soporte y reinicio
with st.sidebar:
    st.markdown(f"**Conectado a:** `{model_name}`")
    if st.button("🗑️ Reiniciar conversación"):
        st.session_state.messages = []
        st.rerun()

# Mostrar mensajes anteriores
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# --- ENTRADA DE MENSAJES ---
if prompt := st.chat_input("Escribe tu pregunta aquí..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        with st.spinner("Whitman AI está pensando..."):
            try:
                # Construir el historial compatible de forma segura
                chat_context = [{"role": "user", "parts": [SYSTEM_PROMPT]}]
                for m in st.session_state.messages:
                    role = "user" if m["role"] == "user" else "model"
                    chat_context.append({"role": role, "parts": [m["content"]]})
                
                response = model.generate_content(chat_context)
                
                if response and response.text:
                    st.markdown(response.text)
                    st.session_state.messages.append({"role": "assistant", "content": response.text})
                else:
                    st.warning("No se recibió respuesta del modelo.")
            except Exception as e:
                st.error(f"Error al responder: {str(e)}")
