import streamlit as st
import mammoth
import pandas as pd
from pptx import Presentation
from markdownify import markdownify as md
import base64
from io import BytesIO
from PIL import Image
import os

st.set_page_config(page_title="Zeins Ultra Converter", page_icon="🎨", layout="wide")

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
        img.thumbnail((1000, 1000)) # Un poco más de calidad para que no sea "feo"
        buffered = BytesIO()
        img.save(buffered, format="JPEG", quality=85)
        img_str = base64.b64encode(buffered.getvalue()).decode()
        return f"\n\n![imagen](data:image/jpeg;base64,{img_str})\n"
    except:
        return ""

st.title("🚀 UARM PPT Reconstructor")
st.markdown("### Convierte con jerarquía visual para Claude/Gemini")

incluir_imagenes = st.checkbox("🖼️ Mantener elementos visuales", value=True)

uploaded_files = st.file_uploader("Arrastra tu PPTX aquí", type=["pptx", "docx", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        ext = file.name.split(".")[-1].lower()
        
        try:
            if ext == "pptx":
                prs = Presentation(file)
                full_md = []
                
                for i, slide in enumerate(prs.slides):
                    # CABECERA DE SLIDE ESTILO "TARJETA"
                    full_md.append(f"\n\n{'='*40}\n## 🎞️ DIAPOSITIVA {i+1}\n{'='*40}\n")
                    
                    # 1. TÍTULO DEL SLIDE (Si existe, siempre va primero)
                    if slide.shapes.title:
                        full_md.append(f"# {slide.shapes.title.text}\n")
                    
                    # 2. CUERPO (Ordenado por lectura natural: Arriba -> Abajo)
                    # Filtramos el título para no repetirlo
                    other_shapes = [s for s in slide.shapes if s != slide.shapes.title]
                    sorted_shapes = sorted(other_shapes, key=lambda s: (s.top, s.left))
                    
                    for shape in sorted_shapes:
                        # Texto en cuadros o formas
                        if hasattr(shape, "text") and shape.text.strip():
                            # Si es un cuadro de texto grande, le damos formato de bloque
                            text = shape.text.strip()
                            if len(text) > 100:
                                full_md.append(f"\n{text}\n")
                            else:
                                full_md.append(f"### {text}")
                        
                        # Imágenes integradas en el flujo
                        elif incluir_imagenes and shape.shape_type == 13:
                            full_md.append(image_to_base64(shape.image.blob))
                        
                        # Tablas dentro del PPT
                        elif shape.has_table:
                            rows = []
                            for row in shape.table.rows:
                                rows.append([cell.text_frame.text.strip() for cell in row.cells])
                            df_tmp = pd.DataFrame(rows)
                            full_md.append("\n" + md(df_tmp.to_html(index=False, header=False)))

                    # 3. NOTAS (Separadas visualmente)
                    if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
                        full_md.append(f"\n\n> 💡 **CONTEXTO ADICIONAL:**\n> {slide.notes_slide.notes_text_frame.text}")

                content = "\n".join(full_md)
            
            # (Mantenemos la lógica de DOCX y XLSX igual que antes...)
            elif ext == "docx":
                content = md(mammoth.convert_to_html(file).value)
            elif ext == "xlsx":
                content = md(pd.read_excel(file).to_html(index=False))

            st.success(f"¡{file.name} reconstruido!")
            st.download_button(f"📥 Descargar {file.name}.md", content, file_name=f"{file.name}.md")
            
        except Exception as e:
            st.error(f"Error crítico: {e}")

st.sidebar.info("Zeins Edition v3.0")
