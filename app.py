import streamlit as st
import mammoth
import pandas as pd
from pptx import Presentation
from markdownify import markdownify as md
import os

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Convertidor MD - Zeins", page_icon="🚀")

# --- CONFIGURACIÓN DE ACCESO (TU CONTRASEÑA) ---
CONTRASEÑA_MAESTRA = "Chris_PAss2026MKD@"

# Manejo de estado de autenticación en la web
if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

# --- PANTALLA DE LOGIN ---
if not st.session_state['autenticado']:
    st.title("🔐 Acceso Restringido")
    st.subheader("UARM Edition - Christopher Ccoicca")
    
    pass_input = st.text_input("Introduce la contraseña de Chris para usar la herramienta:", type="password")
    
    if st.button("Entrar"):
        if pass_input == CONTRASEÑA_MAESTRA:
            st.session_state['autenticado'] = True
            st.success("Acceso concedido. Cargando herramienta...")
            st.rerun()
        else:
            st.error("Contraseña incorrecta. Contacta con Christopher Ccoicca.")
    st.stop()

# --- APP PRINCIPAL (Solo se ejecuta si se autentica) ---
st.title("📄 Compresor de Documentos a Markdown")
st.markdown("---")

# Selector de archivos (Reemplaza a filedialog)
uploaded_files = st.file_uploader(
    "Selecciona archivos (Word, Excel, PPTX)", 
    type=["docx", "xlsx", "pptx"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.info(f"📁 {len(uploaded_files)} archivos listos para procesar")
    
    # Procesar cada archivo
    for file in uploaded_files:
        ext = file.name.split(".")[-1].lower()
        content = ""
        
        try:
            if ext == "docx":
                # Mammoth lee directamente el objeto de archivo de Streamlit
                result = mammoth.convert_to_html(file)
                content = md(result.value)
                
            elif ext == "xlsx":
                df = pd.read_excel(file)
                content = md(df.to_html(index=False))
                
            elif ext == "pptx":
                prs = Presentation(file)
                content = "\n\n".join([shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text")])
            
            # Mostrar resultado y botón de descarga por cada archivo
            with st.expander(f"✅ Vista previa: {file.name}"):
                st.text_area(f"Contenido de {file.name}:", content, height=200)
                
                # Nombre del archivo de salida
                base_name = os.path.splitext(file.name)[0]
                
                st.download_button(
                    label=f"📥 Descargar {base_name}.md",
                    data=content,
                    file_name=f"{base_name}.md",
                    mime="text/markdown",
                    key=file.name # Key única para evitar errores en bucle
                )
                
        except Exception as e:
            st.error(f"Error procesando {file.name}: {str(e)}")

# Pie de página
st.sidebar.markdown("---")
st.sidebar.write("### Creado por:")
st.sidebar.info("Christopher Ccoicca")
