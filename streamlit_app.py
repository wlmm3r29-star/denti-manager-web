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
    from datetime import datetime

    # Leer archivo origen (soporta xls/xlsx)
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), header=None)
    except Exception:
        df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="xlrd")

    # Tomar fecha de impresión desde B1 (fila 0, col 1)
    impresion_origen = ""
    try:
        if isinstance(df.iloc[0, 1], str):
            impresion_origen = df.iloc[0, 1].strip()
    except Exception:
        impresion_origen = ""

    def parse_fecha(x):
        """Convierte fecha que puede venir como str, datetime, Timestamp o vacío."""
        if pd.isna(x):
            return pd.NaT

        # Si ya es fecha
        if isinstance(x, (pd.Timestamp, datetime)):
            return pd.to_datetime(x, errors="coerce")

        # Si es texto
        s = str(x).strip()
        if s.lower() in ("", "nan", "none"):
            return pd.NaT
        s = s.replace("*", "").strip()

        # Intenta dd/mm/yy o dd/mm/yyyy
        return pd.to_datetime(s, dayfirst=True, errors="coerce")

    registros = []
    doctor_actual = ""

    for _, fila in df.iterrows():

        # Detectar doctor (col B = índice 1)
        if isinstance(fila[1], str):
            texto = fila[1].strip()
            if texto.isupper() and "CITAS" not in texto and len(texto) > 5:
                doctor_actual = texto

        # Detectar cita: columna C = índice 2 (puede ser str o fecha real)
        f_cita_dt = parse_fecha(fila[2])

        # Solo procesar filas donde Cita sea válida
        if pd.notna(f_cita_dt):
            fecha_cita_txt = (
                fila[2].replace("*", "").strip()
                if isinstance(fila[2], str)
                else f_cita_dt.strftime("%d/%m/%y")
            )

            nombre = str(fila[5]).strip()       # col F
            telefono = str(fila[6]).strip()     # col G

            # Nueva: columna I = índice 8 (puede ser vacía / str / fecha real)
            nueva_raw = fila[8]
            f_nueva_dt = parse_fecha(nueva_raw)

            nueva_cita_txt = ""
            if pd.notna(f_nueva_dt):
                # conserva el texto original si venía como string; si no, lo formatea
                nueva_cita_txt = (
                    str(nueva_raw).strip()
                    if isinstance(nueva_raw, str)
                    else f_nueva_dt.strftime("%d/%m/%y")
                )

            # ✅ REGLA CORRECTA:
            # Incluir si Nueva está en blanco (NaT) o Nueva <= Cita
            # Excluir solo si Nueva > Cita
            if pd.notna(f_nueva_dt) and f_nueva_dt > f_cita_dt:
                continue

            quien_cancela = (
                str(fila[10]).strip()
                if len(fila) > 10 and pd.notna(fila[10])
                else ""
            )

            motivo = (
                str(fila[11]).strip()
                if len(fila) > 11 and pd.notna(fila[11])
                else ""
            )

            anotaciones = (
                str(fila[12]).strip()
                if len(fila) > 12 and pd.notna(fila[12])
                else ""
            )


            if nombre.lower() != "nan":
                registros.append([
                    fecha_cita_txt,
                    nombre,
                    telefono,
                    nueva_cita_txt,
                    doctor_actual,
                    quien_cancela,
                    motivo,
                    anotaciones
                ])

    df_out = pd.DataFrame(
        registros,
        columns=["Cita", "Nombre", "Telefono", "Nueva", "Doctor", "Quien Cancela","Motivo","Observaciones"]
    )
    df_out.insert(0, "Conse", range(1, len(df_out) + 1))

    # Exportar Excel con la misma fecha de impresión en A1
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

    # -------------------------------------------------
    # 1. Leer archivo origen (para lógica)
    # -------------------------------------------------
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="xlrd")

    # -------------------------------------------------
    # 2. Tomar la fecha/leyenda desde A1 del origen
    # -------------------------------------------------
    encabezado_origen = ""
    try:
        if isinstance(df_raw.iloc[0, 0], str):
            encabezado_origen = df_raw.iloc[0, 0].strip()
    except:
        encabezado_origen = ""

    # -------------------------------------------------
    # 3. LÓGICA QUE YA FUNCIONA (NO TOCADA)
    # -------------------------------------------------
    df = df_raw.copy()

    df["Doctor"] = None
    doctor_actual = None

    for i, row in df.iterrows():
        texto = str(row[0]).strip()
        if texto.isupper() and len(texto.split()) > 1:
            doctor_actual = texto
        df.at[i, "Doctor"] = doctor_actual

    df = df[df[3].notnull() & df[0].notnull()]

    df = df.rename(columns={
        0: "Cita_inici",
        2: "Identifica",
        3: "Nombre_paciente",
        4: "Telefono",
        6: "Nueva_cit"
    })

    df["Cita_inici"] = pd.to_datetime(df["Cita_inici"], errors="coerce")
    df["Nueva_cit"] = pd.to_datetime(df["Nueva_cit"], errors="coerce")

    # ✅ incluir si Nueva_cit está en blanco o <= Cita_inici
    df_filtrado = df[df["Nueva_cit"].isna() | (df["Nueva_cit"] <= df["Cita_inici"])].copy()
    df_filtrado = df_filtrado[df_filtrado["Cita_inici"].notnull()]

    df_filtrado = df_filtrado.reset_index(drop=True)
    df_filtrado.insert(0, "Conse", df_filtrado.index + 1)
    df_filtrado["Anotaciones"] = ""

    # -------------------------------------------------
    # 4. Exportar Excel con encabezado en la fila 1
    # -------------------------------------------------
    temp_out = io.BytesIO()
    df_filtrado.to_excel(temp_out, index=False, startrow=1)
    temp_out.seek(0)

    wb = load_workbook(temp_out)
    ws = wb.active

    if encabezado_origen:
        ws["A1"] = encabezado_origen
        ws["A1"].font = openpyxl.styles.Font(bold=True)

    final_out = io.BytesIO()
    wb.save(final_out)
    final_out.seek(0)

    return final_out, df_filtrado





# ===========================
# UI STREAMLIT
# ===========================

st.set_page_config("Denti Manager Web", layout="centered")
st.title("Denti Manager")

tab1, tab2, tab3, tab4 = st.tabs([
    "📄 PDF → Excel",
    "✍️ Firmar PDFs",
    "🚷 Citas Canceladas",
    "🔄 Citas Inasistidas"
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
        st.download_button(
            label="Descargar",
            data=out,
            file_name=f"CANCELADAS_{now_stamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_cancel"
        )


with tab4:
    f = st.file_uploader("Inasistidas", type=["xls"], key="inasis")
    if st.button("Generar Inasistidas", key="btn_inas", disabled=not f):
        out, df = reprogramar_inasistidas_xls(f.getvalue())
        st.dataframe(df.head())
        st.download_button("Descargar", out, f"INASISTIDAS_{now_stamp()}.xlsx", key="dl_inas")

































