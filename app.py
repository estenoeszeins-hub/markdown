import streamlit as st
import aspose.slides as slides
import aspose.slides.export as export
from markdownify import markdownify as md
import markdown
import streamlit.components.v1 as components
import mammoth
import pandas as pd
import tempfile
import os

# --- CONFIGURACIÓN ---
st.set_page_config(
    page_title="Zeins Elite Compressor",
    page_icon="💎",
    layout="wide"
)

CONTRASEÑA_MAESTRA = "Chris_PAss2026MKD@"

# --- LOGIN SIMPLE ---
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    st.title("🔐 Acceso de Administrador")
    pass_input = st.text_input("Introduce el código de acceso:", type="password")

    if st.button("Activar"):
        if pass_input == CONTRASEÑA_MAESTRA:
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Código incorrecto")

    st.stop()

# --- FUNCIÓN PPTX → MARKDOWN VISUAL ---
def convertir_pptx_elite(file):
    with slides.Presentation(file) as presentation:
        save_options = export.MarkdownSaveOptions()
        save_options.show_hidden_slides = True

        # 🔥 modo visual casi 1:1
        save_options.export_type = export.MarkdownExportType.VISUAL

        with tempfile.NamedTemporaryFile(delete=False, suffix=".md") as tmp:
            temp_path = tmp.name

        presentation.save(temp_path, export.SaveFormat.MD, save_options)

        with open(temp_path, "r", encoding="utf-8") as f:
            content = f.read()

        os.remove(temp_path)
        return content

# --- INTERFAZ ---
st.title("🚀 UARM - Compresor Pro (Vista PPTX Real)")
st.warning("Motor Aspose Visual 1:1 activado")

files = st.file_uploader(
    "📂 Sube tus archivos",
    accept_multiple_files=True
)

if files:
    for f in files:
        ext = f.name.split(".")[-1].lower()

        try:
            with st.spinner(f"⚙️ Procesando {f.name}..."):
                if ext == "pptx":
                    final_content = convertir_pptx_elite(f)

                elif ext == "docx":
                    final_content = md(mammoth.convert_to_html(f).value)

                elif ext == "xlsx":
                    df = pd.read_excel(f)
                    final_content = md(df.to_html(index=False))

                else:
                    st.warning(f"Formato no soportado: {ext}")
                    continue

            st.success(f"✅ {f.name} procesado con éxito")

            # --- VISTA PREVIA VISUAL ---
            st.subheader("👀 Vista previa")
            html_preview = markdown.markdown(
                final_content,
                extensions=["tables"]
            )

            components.html(
                f"""
                <div style="
                    background:white;
                    padding:30px;
                    border-radius:10px;
                    box-shadow:0 0 10px rgba(0,0,0,0.1);
                ">
                    {html_preview}
                </div>
                """,
                height=900,
                scrolling=True
            )

            # --- DESCARGA ---
            st.download_button(
                label=f"📥 Descargar {f.name}.md",
                data=final_content,
                file_name=f"{f.name}.md",
                mime="text/markdown"
            )

        except Exception as e:
            st.error(f"❌ Error en {f.name}: {e}")

st.sidebar.markdown("---")
st.sidebar.write("### 💎 Nivel: Elite")
st.sidebar.info("Desarrollado por: Christopher Ccoicca")
