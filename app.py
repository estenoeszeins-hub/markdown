import streamlit as st
import aspose.slides as slides
import aspose.slides.export as export
from markdownify import markdownify as md
import mammoth
import pandas as pd
import os

# --- CONFIGURACIÓN ELITE ---
st.set_page_config(page_title="Zeins Elite Compressor", page_icon="💎", layout="wide")

CONTRASEÑA_MAESTRA = "Chris_PAss2026MKD@"

if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False

if not st.session_state['autenticado']:
    st.title("🔐 Acceso de Administrador")
    pass_input = st.text_input("Introduce el código de acceso:", type="password")
    if st.button("Activar"):
        if pass_input == CONTRASEÑA_MAESTRA:
            st.session_state['autenticado'] = True
            st.rerun()
    st.stop()

# --- LÓGICA DE CONVERSIÓN DE ALTA FIDELIDAD ---
def convertir_pptx_elite(file):
    # Cargamos el archivo en el motor de Aspose
    with slides.Presentation(file) as presentation:
        # Usamos opciones de exportación para mantener la estructura
        save_options = export.MarkdownSaveOptions()
        save_options.show_hidden_slides = True
        
        # Guardamos temporalmente para capturar el contenido
        temp_path = "temp_output.md"
        presentation.save(temp_path, export.SaveFormat.MD, save_options)
        
        with open(temp_path, "r", encoding="utf-8") as f:
            content = f.read()
        
        os.remove(temp_path)
        return content

# --- INTERFAZ ---
st.title("🚀 UARM - Compresor Pro (Aspose Engine)")
st.warning("Nota: Esta versión utiliza el motor de reconstrucción 1:1. Aparecerá una marca de agua técnica en el archivo, pero la IA la ignorará.")

files = st.file_uploader("Sube tus archivos para deconstrucción total", accept_multiple_files=True)

if files:
    for f in files:
        ext = f.name.split(".")[-1].lower()
        
        try:
            with st.spinner(f"Analizando capas de {f.name}..."):
                if ext == "pptx":
                    # Este método reconstruye el PPT tal cual es
                    final_content = convertir_pptx_elite(f)
                
                elif ext == "docx":
                    final_content = md(mammoth.convert_to_html(f).value)
                
                elif ext == "xlsx":
                    df = pd.read_excel(f)
                    final_content = md(df.to_html(index=False))

            st.success(f"¡{f.name} codificado con éxito!")
            st.download_button(
                label=f"📥 Descargar {f.name}.md (Versión Pro)",
                data=final_content,
                file_name=f"{f.name}.md",
                mime="text/markdown"
            )
        except Exception as e:
            st.error(f"Error en el motor: {e}")

st.sidebar.markdown("---")
st.sidebar.write("### Nivel: Elite")
st.sidebar.info("Desarrollado por: Christopher Ccoicca")
