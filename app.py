import streamlit as st
import mammoth
import pandas as pd
from pptx import Presentation
from markdownify import markdownify as md
import base64
from io import BytesIO
from PIL import Image
import os

st.set_page_config(page_title="Zeins Converter Elite", page_icon="🚀")

CONTRASEÑA_MAESTRA = "Chris_PAss2026MKD@"

if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

if not st.session_state['autenticado']:
    st.title("🔐 Acceso Restringido")
    pass_input = st.text_input("Contraseña:", type="password")
    if st.button("Entrar"):
        if pass_input == CONTRASEÑA_MAESTRA:
            st.session_state['autenticado'] = True
            st.rerun()
    st.stop()

def image_to_base64(image_bytes):
    try:
        img = Image.open(BytesIO(image_bytes))
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        img.thumbnail((800, 800))
        buffered = BytesIO()
        img.save(buffered, format="JPEG", quality=75)
        img_str = base64.b64encode(buffered.getvalue()).decode()
        return f"\n\n![imagen](data:image/jpeg;base64,{img_str})\n\n"
    except:
        return "\n\n> [Error en imagen]\n\n"

st.title("📄 Convertidor de Alta Fidelidad")
st.info("Modo de extracción secuencial (Mantiene el orden original)")

incluir_imagenes = st.checkbox("🖼️ Incluir imágenes (Ordenadas)", value=True)

uploaded_files = st.file_uploader("Selecciona archivos", type=["docx", "xlsx", "pptx"], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        ext = file.name.split(".")[-1].lower()
        content = ""
        
        try:
            if ext == "docx":
                content = md(mammoth.convert_to_html(file).value)
            elif ext == "xlsx":
                df = pd.read_excel(file)
                content = md(df.to_html(index=False))
            elif ext == "pptx":
                prs = Presentation(file)
                slides_output = []
                
                for i, slide in enumerate(prs.slides):
                    slides_output.append(f"--- \n## Slide {i+1}")
                    
                    # ORDENAR ELEMENTOS POR POSICIÓN (Arriba hacia abajo)
                    # Esto asegura que si una imagen está arriba del texto, salga primero
                    shapes = sorted(slide.shapes, key=lambda s: (s.top, s.left))
                    
                    for shape in shapes:
                        # 1. Si es Texto (incluye cuadros, formas con texto, etc.)
                        if hasattr(shape, "text") and shape.text.strip():
                            slides_output.append(shape.text)
                        
                        # 2. Si es Imagen
                        elif incluir_imagenes and shape.shape_type == 13:
                            img_str = image_to_base64(shape.image.blob)
                            slides_output.append(img_str)
                        
                        # 3. Si es un Grupo de formas (procesar sub-formas)
                        elif shape.shape_type == 6: # Group
                            for s in sorted(shape.shapes, key=lambda x: (x.top, x.left)):
                                if hasattr(s, "text") and s.text.strip():
                                    slides_output.append(s.text)
                    
                    # Notas al final del slide
                    if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
                        slides_output.append(f"\n> **Notas:** {slide.notes_slide.notes_text_frame.text}")

                content = "\n\n".join(slides_output)

            with st.expander(f"✅ {file.name}"):
                st.download_button(f"Descargar {file.name}.md", content, file_name=f"{file.name}.md")
        except Exception as e:
            st.error(f"Error: {e}")

st.sidebar.write("By: Christopher Ccoicca")
