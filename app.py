import streamlit as st
import mammoth
import pandas as pd
from pptx import Presentation
from markdownify import markdownify as md
import base64
from io import BytesIO
from PIL import Image
import os

st.set_page_config(page_title="Zeins Converter Pro", page_icon="🚀")

CONTRASEÑA_MAESTRA = "Chris_PAss2026MKD@"

# --- LOGIN ---
if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

if not st.session_state['autenticado']:
    st.title("🔐 Acceso Restringido")
    pass_input = st.text_input("Contraseña:", type="password")
    if st.button("Entrar"):
        if pass_input == CONTRASEÑA_MAESTRA:
            st.session_state['autenticado'] = True
            st.rerun()
        else:
            st.error("Incorrecto")
    st.stop()

# --- INTERFAZ ---
st.title("📄 Convertidor Multi-Formato a MD")
st.info("Creado por Christopher Ccoicca para la UARM")

# --- NUEVA FUNCIÓN: CHECKBOX DE IMÁGENES ---
incluir_imagenes = st.checkbox("🖼️ Incluir imágenes (Aumenta el peso del archivo)", value=False)

uploaded_files = st.file_uploader("Sube tus archivos", type=["docx", "xlsx", "pptx"], accept_multiple_files=True)

def image_to_base64(image_bytes):
    """Optimiza y convierte imagen a Base64 para MD."""
    img = Image.open(BytesIO(image_bytes))
    # Redimensionar para que no pese tanto (Max 800px)
    img.thumbnail((800, 800))
    buffered = BytesIO()
    img.save(buffered, format="JPEG", quality=70) # Comprime a JPG 70%
    img_str = base64.b64encode(buffered.getvalue()).decode()
    return f"![imagen](data:image/jpeg;base64,{img_str})"

if uploaded_files:
    for file in uploaded_files:
        ext = file.name.split(".")[-1].lower()
        content = ""
        
        try:
            if ext == "docx":
                # Mammoth maneja imágenes si se configura, por ahora texto:
                content = md(mammoth.convert_to_html(file).value)
            
            elif ext == "xlsx":
                df = pd.read_excel(file)
                content = md(df.to_html(index=False))
            
            elif ext == "pptx":
                prs = Presentation(file)
                slides_text = []
                for i, slide in enumerate(prs.slides):
                    slides_text.append(f"## Slide {i+1}")
                    # Extraer Texto
                    for shape in slide.shapes:
                        if hasattr(shape, "text"):
                            slides_text.append(shape.text)
                        
                        # --- LÓGICA DE IMÁGENES ---
                        if incluir_imagenes and shape.shape_type == 13: # 13 es tipo Imagen
                            img_bytes = shape.image.blob
                            slides_text.append(image_to_base64(img_bytes))
                    
                    # Extraer Notas del Orador (Opcional pero recomendado)
                    if slide.has_notes_slide:
                        notas = slide.notes_slide.notes_text_frame.text
                        if notas:
                            slides_text.append(f"\n> **Notas:** {notas}")
                
                content = "\n\n".join(slides_text)

            st.success(f"Procesado: {file.name}")
            st.download_button(f"📥 Descargar {file.name}.md", content, file_name=f"{file.name}.md")
            
        except Exception as e:
            st.error(f"Error en {file.name}: {e}")
