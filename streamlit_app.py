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

# ===========================
# UTILIDADES GENERALES
# ===========================

def safe_filename(name: str) -> str:
    name = name.strip().replace("\\", "_").replace("/", "_")
    return re.sub(r"[^a-zA-Z0-9._-]+", "_", name)

def now_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def autosize_columns(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 60)

def ocr_image_to_text(file_bytes: bytes) -> str:
    try:
        img = Image.open(io.BytesIO(file_bytes))
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        return pytesseract.image_to_string(img)
    except Exception:
        return ""

# ===========================
# MÓDULO 1 – PDF / IMÁGENES → EXCEL
# ===========================

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

        tipo_doc = num_doc = nombre = ""

        if ext == ".pdf":
            doc = fitz.open(stream=file_bytes, filetype="pdf")

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
                        m = regex_documento.match(fila_texto.strip())
                        if m:
                            tipo_doc, num_doc, nombre = m.groups()
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

        archivos_procesados += 1

    autosize_columns(ws)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    return out, {"archivos_procesados": archivos_procesados, "filas_agregadas": filas_agregadas}

# ===========================
# MÓDULO 2 – FIRMAR PDFs
# ===========================

def firmar_pdfs_en_zip(pdf_files, firma_file, buscar_texto="Firma Prestador"):
    import zipfile

    firma_bytes = firma_file.getvalue()
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for uf in pdf_files:
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

    zip_buffer.seek(0)
    return zip_buffer

# ===========================
# MÓDULO 3 – CITAS CANCELADAS
# ===========================

def reprogramar_canceladas_excel(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes), header=None)

    registros = []
    doctor_actual = ""

    for _, fila in df.iterrows():
        if isinstance(fila[1], str) and fila[1].isupper():
            doctor_actual = fila[1]

        if isinstance(fila[2], str) and re.match(r"\d{2}/\d{2}/\d{2}", fila[2]):
            fecha = fila[2]
            nombre = fila[5]
            telefono = fila[6]
            nueva = fila[8]

            f1 = pd.to_datetime(fecha, dayfirst=True, errors="coerce")
            f2 = pd.to_datetime(nueva, dayfirst=True, errors="coerce")

            if pd.notna(f2) and f2 > f1:
                continue

            registros.append([fecha, nombre, telefono, nueva, doctor_actual])

    df_out = pd.DataFrame(registros, columns=[
        "Cita_inici", "Nombre", "Telefono", "Nueva_cita", "Doctor"
    ])
    df_out.insert(0, "Conse", range(1, len(df_out) + 1))

    out = io.BytesIO()
    df_out.to_excel(out, index=False)
    out.seek(0)

    return out, df_out

# ===========================
# MÓDULO 4 – CITAS INASISTIDAS
# ===========================

def reprogramar_inasistidas_xls(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="xlrd")

    df["Doctor"] = None
    doctor = None
    for i, fila in df.iterrows():
        if isinstance(fila[0], str) and fila[0].isupper():
            doctor = fila[0]
        df.at[i, "Doctor"] = doctor

    df = df.rename(columns={
        0: "Cita_inici",
        2: "Identifica",
        3: "Nombre",
        4: "Telefono",
        6: "Nueva_cita"
    })

    df["Cita_inici"] = pd.to_datetime(df["Cita_inici"], errors="coerce")
    df["Nueva_cita"] = pd.to_datetime(df["Nueva_cita"], errors="coerce")

    df = df[df["Nueva_cita"] <= df["Cita_inici"]].dropna()

    df.insert(0, "Conse", range(1, len(df) + 1))
    df["Anotaciones"] = ""

    out = io.BytesIO()
    df.to_excel(out, index=False)
    out.seek(0)

    return out, df

# ===========================
# UI STREAMLIT
# ===========================

st.set_page_config(page_title="Denti Manager Web", layout="centered")
st.title("Administrador de dentistas (web)")
st.caption("Procesa archivos desde el navegador y descarga resultados")

tab1, tab2, tab3, tab4 = st.tabs([
    "📄 PDF / Imágenes → Excel",
    "✍️ Firmar PDFs",
    "📅 Citas Canceladas",
    "🚫 Citas Inasistidas"
])

with tab1:
    files = st.file_uploader("Archivos", type=["pdf", "png", "jpg", "jpeg", "tif", "tiff"], accept_multiple_files=True)
    if st.button("Procesar", disabled=not files):
        out, resumen = transformar_archivos_a_excel(files)
        st.success(f"Archivos: {resumen['archivos_procesados']} | Filas: {resumen['filas_agregadas']}")
        st.download_button("Descargar Excel", out, f"PDF_{now_stamp()}.xlsx")

with tab2:
    firma = st.file_uploader("Firma", type=["png", "jpg", "jpeg"])
    pdfs = st.file_uploader("PDFs", type=["pdf"], accept_multiple_files=True)
    if st.button("Firmar", disabled=not (firma and pdfs)):
        zip_out = firmar_pdfs_en_zip(pdfs, firma)
        st.download_button("Descargar ZIP", zip_out, f"FIRMADOS_{now_stamp()}.zip")

with tab3:
    file = st.file_uploader("Excel Canceladas", type=["xls", "xlsx"])
    if st.button("Generar reporte", disabled=not file):
        out, df = reprogramar_canceladas_excel(file.getvalue())
        st.dataframe(df.head())
        st.download_button("Descargar", out, f"CANCELADAS_{now_stamp()}.xlsx")

with tab4:
    file = st.file_uploader("XLS Inasistidas", type=["xls"])
    if st.button("Generar reporte", disabled=not file):
        out, df = reprogramar_inasistidas_xls(file.getvalue())
        st.dataframe(df.head())
        st.download_button("Descargar", out, f"INASISTIDAS_{now_stamp()}.xlsx")
