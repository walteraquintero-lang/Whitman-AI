import streamlit as st
import google.generativeai as genai

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(
    page_title="Whitman AI - Asistente Escolar",
    page_icon="🎓",
    layout="centered"
)

# --- ESTILO Y BIENVENIDA ---
st.title("🎓 Whitman AI")
st.caption("Tu asistente escolar de cabecera. ¡Pregúntame sobre tus clases, proyectos o tareas!")

# --- CONEXIÓN CON LA IA ---
# Toma la clave API de forma segura desde los Secrets de Streamlit
api_key = st.secrets.get("GEMINI_API_KEY")

if not api_key:
    st.error("Falta configurar la clave GEMINI_API_KEY en los secretos.")
    st.stop()

genai.configure(api_key=api_key)

# Personalidad y directivas de Whitman AI (El reemplazo de la Gema)
SYSTEM_PROMPT = """
Eres Whitman AI, el asistente virtual oficial del colegio.
- Tu tono es entusiasta, pedagógico, respetuoso y muy claro.
- Tu objetivo es guiar a los estudiantes para que aprendan y piensen por sí mismos, no solo darles las respuestas directas.
- Ayuda a profesores y padres con dudas sobre actividades y proyectos escolares.
- Mantén siempre un entorno seguro para menores.
"""

# Configuración del modelo (Gemini 1.5 Flash es súper rápido y económico)
model = genai.GenerativeModel(
    model_name="gemini-1.5-flash",
    system_instruction=SYSTEM_PROMPT
)

# --- MEMORIA DE LA CONVERSACIÓN ---
if "messages" not in st.session_state:
    st.session_state.messages = []

# Mostrar el historial en pantalla
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# --- CAJA DE ENTRADA DEL USUARIO ---
if prompt := st.chat_input("Escribe tu pregunta aquí..."):
    # Guardar y mostrar mensaje del usuario
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # Respuesta de Whitman AI
    with st.chat_message("assistant"):
        message_placeholder = st.empty()
        
        # Formatear el historial para Gemini
        chat_history = [
            {"role": "user" if m["role"] == "user" else "model", "parts": [m["content"]]}
            for m in st.session_state.messages[:-1]
        ]
        
        chat = model.start_chat(history=chat_history)
        response = chat.send_message(prompt)
        
        message_placeholder.markdown(response.text)
        st.session_state.messages.append({"role": "assistant", "content": response.text})
