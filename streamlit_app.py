import streamlit as st
import pandas as pd
import re
import io
import os
from datetime import datetime

import fitz  # PyMuPDF
import openpyxl
from openpyxl.utils import get_column_letter

import pytesseract
from PIL import Image

# Para leer .xls (inasistidas) si te llegan en ese formato
# En requirements.txt va xlrd==2.x
# pd.read_excel(..., engine="xlrd")

# ---------------------------
# Utilidades generales
# ---------------------------

def safe_filename(name: str) -> str:
    name = name.strip().replace("\\", "_").replace("/", "_")
    return re.sub(r"[^a-zA-Z0-9._-]+", "_", name)

def now_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def bytes_to_image(file_bytes: bytes) -> Image.Image:
    return Image.open(io.BytesIO(file_bytes)).convert("RGB")

def ocr_image_to_text(file_bytes: bytes) -> str:
    img = bytes_to_image(file_bytes)
    # Puedes mejorar OCR con preprocesamiento si lo necesitas
    return pytesseract.image_to_string(img)

def autosize_columns(ws):
    # Ajuste simple de ancho de columnas
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                max_len = max(max_len, len(str(cell.value)) if cell.value is not None else 0)
            except:
                pass
        ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 60)

# ---------------------------
# Módulo 1: Transformar PDFs/Imágenes -> Excel
# ---------------------------

def transformar_archivos_a_excel(uploaded_files):
    """
    Procesa PDFs e imágenes subidas y genera un Excel en memoria (xlsx).
    - PDFs: extrae texto por página y aplica regex de documento + parsing básico
    - Imágenes: OCR con Tesseract
    """
    regex_documento = re.compile(r"^(CC|TI|CE|RC|NIT)\s+(\d{5,})\s+(.+)$")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Datos de PDF"

    filas_agregadas = 0
    archivos_procesados = 0

    # Encabezado sugerido (puedes ajustarlo)
    ws.append(["TipoDoc", "NumDoc", "Nombre", "Fecha", "Codigo", "Descripcion", "Valor", "Total", "ArchivoOrigen"])

    for uf in uploaded_files:
        nombre_arch = uf.name
        ext = os.path.splitext(nombre_arch)[1].lower()
        file_bytes = uf.getvalue()

        tipo_doc = num_doc = nombre = ""

        if ext == ".pdf":
            try:
                doc = fitz.open(stream=file_bytes, filetype="pdf")
            except Exception as e:
                ws.append(["", "", "", "", "", f"ERROR abriendo PDF: {e}", "", "", nombre_arch])
                continue

            for page in doc:
                try:
                    texto = page.get_text("text") or ""
                except Exception:
                    texto = ""

                for linea in texto.splitlines():
                    s = linea.strip()
                    m = regex_documento.match(s)
                    if m:
                        tipo_doc, num_doc, nombre = m.groups()
                        continue

                    # Parsing “genérico” como en tu script: ajusta según el formato real
                    partes = s.split()
                    if len(partes) >= 6:
                        fecha = partes[0]
                        codigo = partes[1]
                        descripcion = " ".join(partes[2:-2])
                        valor = partes[-2].replace(",", "")
                        total_val = partes[-1].replace(",", "")
                        ws.append([tipo_doc, num_doc, nombre, fecha, codigo, descripcion, valor, total_val, nombre_arch])
                        filas_agregadas += 1

            doc.close()

        elif ext in [".png", ".jpg", ".jpeg", ".tif", ".tiff"]:
            texto = ocr_image_to_text(file_bytes)
            for linea in texto.splitlines():
                s = linea.strip()
                if s:
                    ws.append(["", "", "", "", "", s, "", "", nombre_arch])
                    filas_agregadas += 1
        else:
            ws.append(["", "", "", "", "", f"Formato no soportado: {ext}", "", "", nombre_arch])

        archivos_procesados += 1

    autosize_columns(ws)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    resumen = {
        "archivos_procesados": archivos_procesados,
        "filas_agregadas": filas_agregadas,
    }
    return out, resumen

# ---------------------------
# Módulo 2: Firmar PDFs con imagen de firma
# ---------------------------

def firmar_pdfs_en_zip(pdf_files, firma_file, buscar_texto="Firma Prestador"):
    """
    Firma cada PDF subido insertando la imagen en la última página.
    Si encuentra el texto buscar_texto, firma cerca de ahí; si no, usa coordenadas por defecto.
    Devuelve un ZIP en memoria con los PDFs firmados.
    """
    import zipfile

    firma_bytes = firma_file.getvalue()
    firma_name = firma_file.name

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for uf in pdf_files:
            nombre_pdf = uf.name
            pdf_bytes = uf.getvalue()

            try:
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                page = doc[-1]

                # Buscar texto
                instances = page.search_for(buscar_texto) if buscar_texto else []
                if instances:
                    rect_text = instances[0]
                    x0, y0, x1, y1 = rect_text
                    firma_width = 120
                    firma_height = 50
                    x = x0 + 10
                    y = y0 - (firma_height / 2)
                    rect = fitz.Rect(x, y, x + firma_width, y + firma_height)
                else:
                    rect = fitz.Rect(70, 100, 270, 200)

                # Insertar firma desde bytes (stream)
                page.insert_image(rect, stream=firma_bytes)

                out_pdf = io.BytesIO()
                doc.save(out_pdf)
                doc.close()
                out_pdf.seek(0)

                zf.writestr(safe_filename(nombre_pdf), out_pdf.read())

            except Exception as e:
                # Si falla, dejamos un .txt con el error dentro del zip
                zf.writestr(safe_filename(nombre_pdf) + ".error.txt", str(e))

    zip_buffer.seek(0)
    return zip_buffer

# ---------------------------
# Módulo 3: Reprogramar Citas Canceladas
# ---------------------------

def reprogramar_canceladas_excel(file_bytes: bytes, filename: str):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    hoja = xls.sheet_names[0]
    df_raw = pd.read_excel(xls, sheet_name=hoja, header=None)

    registros = []
    doctor_actual = ""

    for _, fila in df_raw.iterrows():
        if len(fila) > 1 and isinstance(fila[1], str):
            texto = fila[1].strip()
            if texto.isupper() and len(texto) > 5 and "CITAS" not in texto:
                doctor_actual = texto

        if len(fila) > 8 and isinstance(fila[2], str) and re.match(r'\d{2}/\d{2}/\d{2}', fila[2]):
            fecha_inicial = fila[2].strip()
            nombre = str(fila[5]).strip() if len(fila) > 5 else ""
            telefono = str(fila[6]).strip() if len(fila) > 6 else ""
            nueva_cita = str(fila[8]).strip() if pd.notna(fila[8]) else ""

            f1 = pd.to_datetime(fecha_inicial, dayfirst=True, errors="coerce")
            f2 = pd.to_datetime(nueva_cita, dayfirst=True, errors="coerce")

            # Mantiene tu lógica: si la nueva cita es posterior, se omite
            if pd.notna(f2) and pd.notna(f1) and f2 > f1:
                continue

            anotaciones = str(fila[12]).strip() if len(fila) > 12 and pd.notna(fila[12]) else ""
            registros.append([fecha_inicial, nombre, telefono, nueva_cita, doctor_actual, anotaciones])

    df_final = pd.DataFrame(registros, columns=["Cita_inici", "Nombre", "Telefono", "Nueva_cita", "Doctor", "Anotaciones"])
    df_final.insert(0, "Conse", range(1, len(df_final) + 1))

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="REPORTE")
    out.seek(0)

    return out, df_final

# ---------------------------
# Módulo 4: Reprogramar Citas Inasistidas
# ---------------------------

def reprogramar_inasistidas_xls(file_bytes: bytes):
    # Tu archivo original era .xls; usamos xlrd
    df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="xlrd")

    df["Doctor"] = None
    doctor_actual = None

    for i, fila in df.iterrows():
        txt = str(fila[0]).strip()
        if txt.isupper() and len(txt) > 4:
            doctor_actual = txt
        df.at[i, "Doctor"] = doctor_actual

    df = df[df[3].notnull() & df[0].notnull() & df[6].notnull()]

    df = df.rename(columns={
        0: "Cita_inici",
        2: "Identifica",
        3: "Nombre",
        4: "Telefono",
        6: "Nueva_cita"
    })

    df["Cita_inici"] = pd.to_datetime(df["Cita_inici"], errors="coerce")
    df["Nueva_cita"] = pd.to_datetime(df["Nueva_cita"], errors="coerce")

    df_filtrado = df[df["Nueva_cita"] <= df["Cita_inici"]].copy()
    df_filtrado = df_filtrado.dropna(subset=["Cita_inici", "Nueva_cita"]).reset_index(drop=True)

    df_filtrado["Conse"] = df_filtrado.index + 1
    df_filtrado["Anotaciones"] = ""

    # Salida en XLSX (más moderno y fácil en web)
    out = io.BytesIO()
    cols = ["Conse", "Nombre", "Identifica", "Telefono", "Cita_inici", "Nueva_cita", "Doctor", "Anotaciones"]
    df_out = df_filtrado[cols].copy()

    # Formato de fechas similar a tu xls
    df_out["Cita_inici"] = df_out["Cita_inici"].dt.strftime("%d/%m/%Y")
    df_out["Nueva_cita"] = df_out["Nueva_cita"].dt.strftime("%d/%m/%Y")

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="REPORTE")
    out.seek(0)

    return out, df_out

# ---------------------------
# UI Streamlit
# ---------------------------

st.set_page_config(page_title="Denti Manager Web", layout="centered")
st.title("Denti Manager (Web)")
st.caption("Versión web del sistema: procesa archivos desde el navegador y descarga resultados. Sin instalaciones.")

tab1, tab2, tab3, tab4 = st.tabs([
    "📄 Transformar PDFs/Imágenes → Excel",
    "✍️ Firmar PDFs",
    "📅 Reprogramar Canceladas",
    "🚫 Reprogramar Inasistidas"
])

with tab1:
    st.subheader("Transformar PDFs/Imágenes a Excel")
    st.write("Sube uno o varios archivos (PDF/PNG/JPG/TIFF) y descarga el Excel consolidado.")
    files = st.file_uploader(
        "Archivos",
        type=["pdf", "png", "jpg", "jpeg", "tif", "tiff"],
        accept_multiple_files=True
    )
    if st.button("🚀 Procesar", key="proc1", disabled=not files):
        with st.spinner("Procesando..."):
            out, resumen = transformar_archivos_a_excel(files)
        st.success(f"Listo. Archivos: {resumen['archivos_procesados']} | Filas: {resumen['filas_agregadas']}")
        st.download_button(
            "⬇️ Descargar Excel",
            data=out,
            file_name=f"DATOS_PDF_{now_stamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

with tab2:
    st.subheader("Firmar PDFs con imagen")
    st.write("Sube PDFs y una imagen de firma (PNG/JPG). Descarga un ZIP con los PDFs firmados.")
    firma = st.file_uploader("Firma (PNG/JPG)", type=["png", "jpg", "jpeg"], key="firma")
    pdfs = st.file_uploader("PDFs a firmar", type=["pdf"], accept_multiple_files=True, key="pdfs_firma")
    buscar_txt = st.text_input("Texto de referencia (opcional)", value="Firma Prestador")
    if st.button("✍️ Firmar", key="proc2", disabled=not (firma and pdfs)):
        with st.spinner("Firmando PDFs..."):
            zip_out = firmar_pdfs_en_zip(pdfs, firma, buscar_texto=buscar_txt.strip())
        st.success("Listo. Descarga el ZIP.")
        st.download_button(
            "⬇️ Descargar ZIP (PDFs firmados)",
            data=zip_out,
            file_name=f"PDFS_FIRMADOS_{now_stamp()}.zip",
            mime="application/zip"
        )

with tab3:
    st.subheader("Reprogramar Citas Canceladas")
    st.write("Sube el Excel de canceladas (.xls/.xlsx). Devuelve un Excel con el reporte.")
    canc = st.file_uploader("Archivo de canceladas", type=["xls", "xlsx"], key="canceladas")
    if st.button("📅 Generar reporte", key="proc3", disabled=not canc):
        with st.spinner("Procesando..."):
            out, df_prev = reprogramar_canceladas_excel(canc.getvalue(), canc.name)
        if df_prev.empty:
            st.warning("No hubo registros válidos.")
        else:
            st.success(f"Listo. Registros: {len(df_prev)}")
            st.dataframe(df_prev.head(50), use_container_width=True)
            st.download_button(
                "⬇️ Descargar reporte",
                data=out,
                file_name=f"CITAS_CANCELADAS_{now_stamp()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

with tab4:
    st.subheader("Reprogramar Citas Inasistidas")
    st.write("Sube el archivo .xls de inasistidas. Devuelve un XLSX con el reporte filtrado.")
    ina = st.file_uploader("Archivo INASISTIDAS (.xls)", type=["xls"], key="inasistidas")
    if st.button("🚫 Generar reporte", key="proc4", disabled=not ina):
        with st.spinner("Procesando..."):
            out, df_prev = reprogramar_inasistidas_xls(ina.getvalue())
        if df_prev.empty:
            st.warning("No hay registros para guardar.")
        else:
            st.success(f"Listo. Registros: {len(df_prev)}")
            st.dataframe(df_prev.head(50), use_container_width=True)
            st.download_button(
                "⬇️ Descargar reporte",
                data=out,
                file_name=f"INASISTIDAS_{now_stamp()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
