import streamlit as st
import mammoth
import pandas as pd
from pptx import Presentation
from markdownify import markdownify as md
import base64
from io import BytesIO
from PIL import Image
import os

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Zeins Converter Pro", page_icon="🚀")

# --- CONFIGURACIÓN DE ACCESO ---
CONTRASEÑA_MAESTRA = "Chris_PAss2026MKD@"

if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

# --- SISTEMA DE LOGIN ---
if not st.session_state['autenticado']:
    st.title("🔐 Acceso Restringido")
    st.subheader("UARM Edition - Christopher Ccoicca")
    
    pass_input = st.text_input("Introduce la contraseña de Chris:", type="password")
    
    if st.button("Entrar"):
        if pass_input == CONTRASEÑA_MAESTRA:
            st.session_state['autenticado'] = True
            st.success("Acceso concedido...")
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")
    st.stop()

# --- FUNCIONES DE PROCESAMIENTO ---
def image_to_base64(image_bytes):
    """
    Optimiza la imagen: resuelve el error 'Mode P', 
    redimensiona a 800px y comprime a JPG.
    """
    try:
        img = Image.open(BytesIO(image_bytes))
        
        # Corregir error de modo P o RGBA (transparencias) para poder guardar como JPEG
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        
        # Redimensionar para ahorrar tokens y peso (máximo 800px)
        img.thumbnail((800, 800))
        
        buffered = BytesIO()
        img.save(buffered, format="JPEG", quality=70) # Calidad 70% para equilibrio peso/detalle
        img_str = base64.b64encode(buffered.getvalue()).decode()
        
        return f"\n\n![imagen](data:image/jpeg;base64,{img_str})\n\n"
    except Exception as e:
        return f"\n\n> [!ERROR] No se pudo procesar una imagen: {e}\n\n"

# --- INTERFAZ PRINCIPAL ---
st.title("📄 Convertidor de Documentos a Markdown")
st.info("Herramienta optimizada para IA (Claude/Gemini/GPT)")

# Opciones de conversión
col1, col2 = st.columns(2)
with col1:
    incluir_imagenes = st.checkbox("🖼️ Incluir imágenes (Base64)", value=False)
with col2:
    incluir_notas = st.checkbox("💡 Incluir notas del orador", value=True)

uploaded_files = st.file_uploader("Selecciona archivos (DOCX, XLSX, PPTX)", 
                                  type=["docx", "xlsx", "pptx"], 
                                  accept_multiple_files=True)

if uploaded_files:
    st.markdown("---")
    for file in uploaded_files:
        ext = file.name.split(".")[-1].lower()
        content = ""
        
        try:
            # --- PROCESAR WORD ---
            if ext == "docx":
                html = mammoth.convert_to_html(file).value
                content = md(html)
            
            # --- PROCESAR EXCEL ---
            elif ext == "xlsx":
                df = pd.read_excel(file)
                html_table = df.to_html(index=False)
                content = md(html_table)
            
            # --- PROCESAR POWERPOINT ---
            elif ext == "pptx":
                prs = Presentation(file)
                slides_text = [f"# Presentación: {file.name}\n"]
                
                for i, slide in enumerate(prs.slides):
                    slides_text.append(f"--- \n## Slide {i+1}")
                    
                    # Extraer texto y/o imágenes por cada forma en el slide
                    for shape in slide.shapes:
                        # Texto
                        if hasattr(shape, "text") and shape.text.strip():
                            slides_text.append(shape.text)
                        
                        # Imágenes (Solo si el checkbox está activo)
                        if incluir_imagenes and shape.shape_type == 13: # 13 es tipo Imagen
                            img_base64 = image_to_base64(shape.image.blob)
                            slides_text.append(img_base64)
                    
                    # Notas del orador (Si el checkbox está activo)
                    if incluir_notas and slide.has_notes_slide:
                        notas = slide.notes_slide.notes_text_frame.text
                        if notas.strip():
                            slides_text.append(f"\n> **Notas del slide:** {notas}")
                
                content = "\n\n".join(slides_text)

            # --- RESULTADOS ---
            with st.expander(f"✅ {file.name} listo"):
                st.text_area("Vista previa (Markdown):", content[:1000] + "...", height=150)
                base_name = os.path.splitext(file.name)[0]
                st.download_button(
                    label=f"📥 Descargar {base_name}.md",
                    data=content,
                    file_name=f"{base_name}.md",
                    mime="text/markdown",
                    key=file.name
                )
                
        except Exception as e:
            st.error(f"Error procesando {file.name}: {str(e)}")

# PIE DE PÁGINA
st.sidebar.markdown("---")
st.sidebar.write("### Creado por:")
st.sidebar.success("Christopher Ccoicca")
st.sidebar.write("Universidad Antonio Ruiz de Montoya")
