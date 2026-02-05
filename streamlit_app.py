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
    df = pd.read_excel(io.BytesIO(file_bytes), header=None)
    out_df = []
    doctor = ""
    for _, r in df.iterrows():
        if isinstance(r[1], str) and r[1].isupper():
            doctor = r[1]
        if isinstance(r[2], str) and re.match(r"\d{2}/\d{2}/\d{2}", r[2]):
            f1 = pd.to_datetime(r[2], dayfirst=True, errors="coerce")
            f2 = pd.to_datetime(r[8], dayfirst=True, errors="coerce")
            if pd.notna(f2) and f2 > f1:
                continue
            out_df.append([r[2], r[5], r[6], r[8], doctor])

    df_out = pd.DataFrame(out_df, columns=["Cita","Nombre","Telefono","Nueva","Doctor"])
    df_out.insert(0,"Conse",range(1,len(df_out)+1))
    out = io.BytesIO()
    df_out.to_excel(out,index=False)
    out.seek(0)
    return out, df_out

# ===========================
# MÓDULO 4 – INASISTIDAS
# ===========================

def reprogramar_inasistidas_xls(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="xlrd")
    doctor = None
    df["Doctor"] = None
    for i,r in df.iterrows():
        if isinstance(r[0], str) and r[0].isupper():
            doctor = r[0]
        df.at[i,"Doctor"] = doctor

    df = df.rename(columns={0:"Cita",2:"ID",3:"Nombre",4:"Telefono",6:"Nueva"})
    df["Cita"] = pd.to_datetime(df["Cita"], errors="coerce")
    df["Nueva"] = pd.to_datetime(df["Nueva"], errors="coerce")
    df = df[df["Nueva"] <= df["Cita"]].dropna()
    df.insert(0,"Conse",range(1,len(df)+1))
    out = io.BytesIO()
    df.to_excel(out,index=False)
    out.seek(0)
    return out, df

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




