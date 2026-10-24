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

# --- DETECCIÓN AUTOMÁTICA DEL MODELO DISPONIBLE ---
@st.cache_resource
def get_model():
    # Buscar qué modelos tiene habilitados tu cuenta de Google
    available_models = [
        m.name for m in genai.list_models() 
        if "generateContent" in m.supported_generation_methods
    ]
    
    # Priorizar modelos Flash
    chosen_model = next((m for m in available_models if "flash" in m.lower()), None)
    
    # Si no hay flash, tomar el primero que sirva (ej. gemini-pro)
    if not chosen_model and available_models:
        chosen_model = available_models[0]
        
    return genai.GenerativeModel(
        model_name=chosen_model,
        system_instruction=SYSTEM_PROMPT
    )

try:
    model = get_model()
except Exception as e:
    st.error(f"Error conectando con los modelos de Google: {e}")
    st.stop()

# --- MEMORIA DEL CHAT ---
if "messages" not in st.session_state:
    st.session_state.messages = []

# Mostrar historial
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# --- ENTRADA DE MENSAJES ---
if prompt := st.chat_input("Escribe tu pregunta aquí..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        message_placeholder = st.empty()
        try:
            # Construir historial para responder en contexto
            chat_history = [
                {"role": "user" if m["role"] == "user" else "model", "parts": [m["content"]]}
                for m in st.session_state.messages[:-1]
            ]
            chat = model.start_chat(history=chat_history)
            response = chat.send_message(prompt)
            message_placeholder.markdown(response.text)
            st.session_state.messages.append({"role": "assistant", "content": response.text})
        except Exception as e:
            message_placeholder.error(f"Error al responder: {str(e)}")
