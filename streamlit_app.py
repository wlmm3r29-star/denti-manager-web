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

# ---------------------------
# Utilidades generales
# ---------------------------

def safe_filename(name: str) -> str:
    name = name.strip().replace("\\", "_").replace("/", "_")
    return re.sub(r"[^a-zA-Z0-9._-]+", "_", name)

def now_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def bytes_to_image(file_bytes: bytes) -> Image.Image:
    img = Image.open(io.BytesIO(file_bytes))
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    return img

def ocr_image_to_text(file_bytes: bytes) -> str:
    try:
        img = bytes_to_image(file_bytes)
        return pytesseract.image_to_string(img)
    except Exception:
        return ""

def autosize_columns(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 60)

# ---------------------------
# MÓDULO 1 – PDF / IMÁGENES → EXCEL (CORREGIDO)
# ---------------------------

def transformar_archivos_a_excel(uploaded_files):
    regex_documento = re.compile(r"^(CC|TI|CE|RC|NIT)\s+(\d{5,})\s+(.+)$")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Datos de PDF"

    ws.append([
        "TipoDoc", "NumDoc", "Nombre",
        "Col1", "Col2", "Col3", "Col4", "Col5",
        "ArchivoOrigen"
    ])

    filas_agregadas = 0
    archivos_procesados = 0

    for uf in uploaded_files:
        nombre_arch = uf.name
        ext = os.path.splitext(nombre_arch)[1].lower()
        file_bytes = uf.getvalue()

        tipo_doc = ""
        num_doc = ""
        nombre = ""

        if ext == ".pdf":
            try:
                doc = fitz.open(stream=file_bytes, filetype="pdf")
            except Exception as e:
                ws.append(["", "", "", f"ERROR PDF: {e}", "", "", "", "", nombre_arch])
                continue

            for page in doc:
                tablas = page.find_tables()
                if not tablas:
                    continue

                for tabla in tablas:
                    matriz = tabla.extract()

                    for fila in matriz:
                        if not any(str(c).strip() for c in fila if c):
                            continue

                        fila_texto = " ".join(str(c) for c in fila if c)
                        match = regex_documento.match(fila_texto.strip())
                        if match:
                            tipo_doc = match.group(1)
                            num_doc = match.group(2)
                            nombre = match.group(3)
                            continue

                        fila_limpia = []
                        for celda in fila:
                            if isinstance(celda, str):
                                celda = celda.replace("$", "").replace(",", "").strip()
                                try:
                                    celda = float(celda)
                                except:
                                    pass
                            fila_limpia.append(celda)

                        ws.append(
                            [tipo_doc, num_doc, nombre] +
                            fila_limpia +
                            [nombre_arch]
                        )
                        filas_agregadas += 1

                    ws.append([])

            doc.close()

        elif ext in [".png", ".jpg", ".jpeg", ".tif", ".tiff"]:
            texto = ocr_image_to_text(file_bytes)
            for linea in texto.splitlines():
                if linea.strip():
                    ws.append(["", "", "", linea.strip(), "", "", "", "", nombre_arch])
                    filas_agregadas += 1

        else:
            ws.append(["", "", "", f"Formato no soportado: {ext}", "", "", "", "", nombre_arch])

        archivos_procesados += 1

    autosize_columns(ws)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    return out, {
        "archivos_procesados": archivos_procesados,
        "filas_agregadas": filas_agregadas
    }

# ---------------------------
# MÓDULO 2 – Firmar PDFs
# ---------------------------

def firmar_pdfs_en_zip(pdf_files, firma_file, buscar_texto="Firma Prestador"):
    import zipfile

    firma_bytes = firma_file.getvalue()
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for uf in pdf_files:
            try:
                doc = fitz.open(stream=uf.getvalue(), filetype="pdf")
                page = doc[-1]

                inst = page.search_for(buscar_texto) if buscar_texto else []
                rect = inst[0] if inst else fitz.Rect(70, 100, 270, 200)

                page.insert_image(rect, stream=firma_bytes)

                out_pdf = io.BytesIO()
                doc.save(out_pdf)
                doc.close()
                out_pdf.seek(0)

                zf.writestr(safe_filename(uf.name), out_pdf.read())
            except Exception as e:
                zf.writestr(safe_filename(uf.name) + ".error.txt", str(e))

    zip_buffer.seek(0)
    return zip_buffer

# ---------------------------
# UI STREAMLIT
# ---------------------------

st.set_page_config(page_title="Denti Manager Web", layout="centered")
st.title("Administrador de dentistas (web)")
st.caption("Versión web del sistema: procesa archivos desde el navegador y descarga resultados. Sin instalaciones.")

tab1, tab2 = st.tabs([
    "📄 Transformar PDFs/Imágenes → Excel",
    "✍️ Firmar archivos PDF"
])

with tab1:
    files = st.file_uploader(
        "Archivos PDF / Imagen",
        type=["pdf", "png", "jpg", "jpeg", "tif", "tiff"],
        accept_multiple_files=True
    )
    if st.button("🚀 Procesar", disabled=not files):
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
    firma = st.file_uploader("Firma (PNG/JPG)", type=["png", "jpg", "jpeg"])
    pdfs = st.file_uploader("PDFs a firmar", type=["pdf"], accept_multiple_files=True)
    if st.button("✍️ Firmar PDFs", disabled=not (firma and pdfs)):
        zip_out = firmar_pdfs_en_zip(pdfs, firma)
        st.download_button(
            "⬇️ Descargar ZIP",
            data=zip_out,
            file_name=f"PDFS_FIRMADOS_{now_stamp()}.zip",
            mime="application/zip"
        )
