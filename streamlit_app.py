import streamlit as st
import pandas as pd
import re
import io
import os
from datetime import datetime
import fitz
import openpyxl
from openpyxl.utils import get_column_letter
import pytesseract
from PIL import Image

# ===========================
# UTILIDADES GENERALES
# ===========================

def safe_filename(name):
    return re.sub(r"[^a-zA-Z0-9._-]+", "_", name)

def now_stamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def autosize_columns(ws):
    for col in ws.columns:
        max_len = max(len(str(c.value)) if c.value else 0 for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(10, max_len + 2), 60)

def ocr_image_to_text(file_bytes):
    try:
        img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
        return pytesseract.image_to_string(img)
    except:
        return ""

# ===========================
# MÓDULO 1 – PDF → EXCEL (TABLAS)
# ===========================

def transformar_archivos_a_excel(uploaded_files):
    regex_documento = re.compile(r"^(CC|TI|CE|RC|NIT)\s+(\d{5,})\s+(.+)$")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Datos PDF"
    ws.append(["TipoDoc","NumDoc","Nombre","Col1","Col2","Col3","Col4","Col5","Archivo"])

    filas = archivos = 0

    for uf in uploaded_files:
        archivos += 1
        doc = fitz.open(stream=uf.getvalue(), filetype="pdf")
        tipo = num = nombre = ""

        for page in doc:
            for tabla in page.find_tables():
                for fila in tabla.extract():
                    if not any(fila):
                        continue
                    texto = " ".join(str(c) for c in fila if c)
                    m = regex_documento.match(texto)
                    if m:
                        tipo, num, nombre = m.groups()
                        continue

                    fila_limpia = []
                    for c in fila:
                        if isinstance(c, str):
                            c = c.replace("$","").replace(",","").strip()
                            try: c = float(c)
                            except: pass
                        fila_limpia.append(c)

                    ws.append([tipo,num,nombre] + fila_limpia + [uf.name])
                    filas += 1

        doc.close()

    autosize_columns(ws)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out, archivos, filas

# ===========================
# MÓDULO 2 – FIRMAR PDFs
# ===========================

def firmar_pdfs_en_zip(pdfs, firma):
    import zipfile

    firma_bytes = firma.getvalue()
    z = io.BytesIO()

    with zipfile.ZipFile(z, "w", zipfile.ZIP_DEFLATED) as zipf:
        for pdf in pdfs:
            doc = fitz.open(stream=pdf.getvalue(), filetype="pdf")
            page = doc[-1]

            instancias = page.search_for("Firma Prestador")

            # Tamaño de la firma (ajustable)
            firma_width = 140
            firma_height = 55

            if instancias:
                rect_texto = instancias[0]

                # 👉 Colocar la firma SOBRE la línea
                x = rect_texto.x0
                y = rect_texto.y1 - firma_height + 8

                rect_firma = fitz.Rect(
                    x,
                    y,
                    x + firma_width,
                    y + firma_height
                )
            else:
                rect_firma = fitz.Rect(70, 130, 210, 185)

            page.insert_image(rect_firma, stream=firma_bytes)

            buf = io.BytesIO()
            doc.save(buf)
            doc.close()

            zipf.writestr(pdf.name, buf.getvalue())

    z.seek(0)
    return z



# ===========================
# MÓDULO 3 – CANCELADAS
# ===========================

def reprogramar_canceladas_excel(file_bytes):
    import io
    import re
    import pandas as pd
    import openpyxl
    from openpyxl import load_workbook

    # Leer archivo origen
    df = pd.read_excel(io.BytesIO(file_bytes), header=None)

    # -------------------------------------------------
    # TOMAR FECHA DE IMPRESIÓN DESDE B1 (fila 0, col 1)
    # -------------------------------------------------
    impresion_origen = ""
    try:
        if isinstance(df.iloc[0, 1], str) and "Impres" in df.iloc[0, 1]:
            impresion_origen = df.iloc[0, 1].strip()
    except:
        impresion_origen = ""

    registros = []
    doctor_actual = ""

    for _, fila in df.iterrows():

        if isinstance(fila[1], str):
            texto = fila[1].strip()
            if texto.isupper() and "CITAS" not in texto and len(texto) > 5:
                doctor_actual = texto

        if isinstance(fila[2], str) and re.match(r"\*?\d{2}/\d{2}/\d{2}", fila[2]):
            fecha_cita = fila[2].replace("*", "").strip()
            nombre = str(fila[5]).strip()
            telefono = str(fila[6]).strip()
            nueva_cita = str(fila[8]).strip() if pd.notna(fila[8]) else ""

            fecha_dt = pd.to_datetime(fecha_cita, dayfirst=True, errors="coerce")
            nueva_dt = pd.to_datetime(nueva_cita, dayfirst=True, errors="coerce")

            if pd.notna(nueva_dt) and nueva_dt > fecha_dt:
                continue

            anotaciones = (
                str(fila[12]).strip()
                if len(fila) > 12 and pd.notna(fila[12])
                else ""
            )

            if nombre.lower() != "nan":
                registros.append([
                    fecha_cita,
                    nombre,
                    telefono,
                    nueva_cita,
                    doctor_actual,
                    anotaciones
                ])

    df_out = pd.DataFrame(
        registros,
        columns=["Cita", "Nombre", "Telefono", "Nueva", "Doctor", "Anotaciones"]
    )

    df_out.insert(0, "Conse", range(1, len(df_out) + 1))

    # -------------------------------------------------
    # EXPORTAR EXCEL CON LA MISMA FECHA DE IMPRESIÓN
    # -------------------------------------------------
    temp_output = io.BytesIO()
    df_out.to_excel(temp_output, index=False, startrow=1)
    temp_output.seek(0)

    wb = load_workbook(temp_output)
    ws = wb.active

    if impresion_origen:
        ws["A1"] = impresion_origen
        ws["A1"].font = openpyxl.styles.Font(bold=True)

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    return final_output, df_out


# ===========================
# MÓDULO 4 – INASISTIDAS
# ===========================

def reprogramar_inasistidas_xls(file_bytes):
    import io
    import pandas as pd
    import openpyxl
    from openpyxl import load_workbook

    # -----------------------------------------
    # Leer archivo origen (sin encabezados)
    # -----------------------------------------
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="xlrd")

    # -----------------------------------------
    # Tomar fecha de impresión desde A1
    # -----------------------------------------
    impresion_origen = ""
    try:
        if isinstance(df_raw.iloc[0, 0], str) and "Impres" in df_raw.iloc[0, 0]:
            impresion_origen = df_raw.iloc[0, 0].strip()
    except:
        impresion_origen = ""

    # -----------------------------------------
    # Lógica original
    # -----------------------------------------
    df = df_raw.copy()
    doctor = None
    df["Doctor"] = None

    for i, r in df.iterrows():
        if isinstance(r[0], str) and r[0].isupper():
            doctor = r[0]
        df.at[i, "Doctor"] = doctor

    df = df.rename(columns={
        0: "Cita",
        2: "ID",
        3: "Nombre",
        4: "Telefono",
        6: "Nueva"
    })

    df["Cita"] = pd.to_datetime(df["Cita"], errors="coerce")
    df["Nueva"] = pd.to_datetime(df["Nueva"], errors="coerce")

    df = df[df["Nueva"] <= df["Cita"]].dropna()
    df.insert(0, "Conse", range(1, len(df) + 1))

    # -----------------------------------------
    # Exportar Excel con fila 1 = fecha impresión
    # -----------------------------------------
    temp_output = io.BytesIO()
    df.to_excel(temp_output, index=False, startrow=1)
    temp_output.seek(0)

    wb = load_workbook(temp_output)
    ws = wb.active

    if impresion_origen:
        ws["A1"] = impresion_origen
        ws["A1"].font = openpyxl.styles.Font(bold=True)

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    return final_output, df

# ===========================
# UI STREAMLIT
# ===========================

st.set_page_config("Denti Manager Web", layout="centered")
st.title("Administrador de Dentistas (Web)")

tab1, tab2, tab3, tab4 = st.tabs([
    "📄 PDF → Excel",
    "✍️ Firmar PDFs",
    "📅 Canceladas",
    "🚫 Inasistidas"
])

with tab1:
    files = st.file_uploader("PDFs", type=["pdf"], accept_multiple_files=True)
    if st.button("Procesar PDFs", key="btn_pdf", disabled=not files):
        out, a, f = transformar_archivos_a_excel(files)
        st.success(f"Archivos: {a} | Filas: {f}")
        st.download_button("Descargar Excel", out, f"PDF_{now_stamp()}.xlsx", key="dl_pdf")

with tab2:
    firma = st.file_uploader("Firma", type=["png","jpg"], key="firma")
    pdfs = st.file_uploader("PDFs", type=["pdf"], accept_multiple_files=True, key="pdfs")
    if st.button("Firmar PDFs", key="btn_firmar", disabled=not (firma and pdfs)):
        z = firmar_pdfs_en_zip(pdfs, firma)
        st.download_button("Descargar ZIP", z, f"FIRMADOS_{now_stamp()}.zip", key="dl_zip")

with tab3:
    f = st.file_uploader("Canceladas", type=["xls","xlsx"], key="cancel")
    if st.button("Generar Canceladas", key="btn_cancel", disabled=not f):
        out, df = reprogramar_canceladas_excel(f.getvalue())
        st.dataframe(df.head())
        st.download_button("Descargar", out, f"CANCELADAS_{now_stamp()}.xlsx", key="dl_cancel")

with tab4:
    f = st.file_uploader("Inasistidas", type=["xls"], key="inasis")
    if st.button("Generar Inasistidas", key="btn_inas", disabled=not f):
        out, df = reprogramar_inasistidas_xls(f.getvalue())
        st.dataframe(df.head())
        st.download_button("Descargar", out, f"INASISTIDAS_{now_stamp()}.xlsx", key="dl_inas")















