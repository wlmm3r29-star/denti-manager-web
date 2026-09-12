import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
import fitz
import openpyxl
from openpyxl.utils import get_column_letter
import pytesseract
from PIL import Image

# ===========================
# ESTADO STREAMLIT
# ===========================
if "coomeva_uploader_key" not in st.session_state:
    st.session_state.coomeva_uploader_key = 0

if "coomeva_resultado" not in st.session_state:
    st.session_state.coomeva_resultado = None

if "coomeva_preview" not in st.session_state:
    st.session_state.coomeva_preview = None

if "coomeva_errores" not in st.session_state:
    st.session_state.coomeva_errores = None

if "coomeva_resumen" not in st.session_state:
    st.session_state.coomeva_resumen = None


def limpiar_coomeva():
    st.session_state.coomeva_uploader_key += 1
    st.session_state.coomeva_resultado = None
    st.session_state.coomeva_preview = None
    st.session_state.coomeva_errores = None
    st.session_state.coomeva_resumen = None


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
    except Exception:
        return ""


# ===========================
# MÓDULO 1 – PDF → EXCEL (TABLAS)
# ===========================
def transformar_archivos_a_excel(uploaded_files):
    regex_documento = re.compile(r"^(CC|TI|CE|RC|NIT)\s+(\d{5,})\s+(.+)$")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Datos PDF"
    ws.append(["TipoDoc", "NumDoc", "Nombre", "Col1", "Col2", "Col3", "Col4", "Col5", "Archivo"])

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
                            c = c.replace("$", "").replace(",", "").strip()
                            try:
                                c = float(c)
                            except Exception:
                                pass
                        fila_limpia.append(c)

                    ws.append([tipo, num, nombre] + fila_limpia + [uf.name])
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
def firmar_pdfs_en_zip(pdfs):
    import zipfile

    with open("firma.png", "rb") as f:
        firma_bytes = f.read()

    img = Image.open(io.BytesIO(firma_bytes)).convert("RGBA")
    buffer_img = io.BytesIO()
    img.save(buffer_img, format="PNG")
    firma_bytes = buffer_img.getvalue()

    z = io.BytesIO()

    with zipfile.ZipFile(z, "w", zipfile.ZIP_DEFLATED) as zipf:
        for pdf in pdfs:
            doc = fitz.open(stream=pdf.getvalue(), filetype="pdf")
            page = doc[-1]

            instancias = page.search_for("Firma Prestador")

            firma_width = 140
            firma_height = 55

            if instancias:
                rect_texto = instancias[0]
                x = rect_texto.x0
                y = rect_texto.y1 - firma_height + 8
                rect_firma = fitz.Rect(x, y, x + firma_width, y + firma_height)
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
    from openpyxl import load_workbook

    try:
        df = pd.read_excel(io.BytesIO(file_bytes), header=None)
    except Exception:
        df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="xlrd")

    impresion_origen = ""
    try:
        if isinstance(df.iloc[0, 1], str):
            impresion_origen = df.iloc[0, 1].strip()
    except Exception:
        impresion_origen = ""

    def parse_fecha(x):
        if pd.isna(x):
            return pd.NaT

        if isinstance(x, (pd.Timestamp, datetime)):
            return pd.to_datetime(x, errors="coerce")

        s = str(x).strip()
        if s.lower() in ("", "nan", "none"):
            return pd.NaT
        s = s.replace("*", "").strip()

        return pd.to_datetime(s, dayfirst=True, errors="coerce")

    registros = []
    doctor_actual = ""

    for _, fila in df.iterrows():
        if isinstance(fila[1], str):
            texto = fila[1].strip()
            if texto.isupper() and "CITAS" not in texto and len(texto) > 5:
                doctor_actual = texto

        f_cita_dt = parse_fecha(fila[2])

        if pd.notna(f_cita_dt):
            fecha_cita_txt = (
                fila[2].replace("*", "").strip()
                if isinstance(fila[2], str)
                else f_cita_dt.strftime("%d/%m/%y")
            )

            nombre = str(fila[5]).strip()
            telefono = str(fila[6]).strip()

            nueva_raw = fila[8]
            f_nueva_dt = parse_fecha(nueva_raw)

            nueva_cita_txt = ""
            if pd.notna(f_nueva_dt):
                nueva_cita_txt = (
                    str(nueva_raw).strip()
                    if isinstance(nueva_raw, str)
                    else f_nueva_dt.strftime("%d/%m/%y")
                )

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
        columns=["Cita", "Nombre", "Telefono", "Nueva", "Doctor", "Quien Cancela", "Motivo", "Observaciones"]
    )
    df_out.insert(0, "Conse", range(1, len(df_out) + 1))

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
    from openpyxl import load_workbook

    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine="xlrd")

    encabezado_origen = ""
    try:
        if isinstance(df_raw.iloc[0, 0], str):
            encabezado_origen = df_raw.iloc[0, 0].strip()
    except Exception:
        encabezado_origen = ""

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

    df_filtrado = df[df["Nueva_cit"].isna() | (df["Nueva_cit"] <= df["Cita_inici"])].copy()
    df_filtrado = df_filtrado[df_filtrado["Cita_inici"].notnull()]

    df_filtrado = df_filtrado.reset_index(drop=True)
    df_filtrado.insert(0, "Conse", df_filtrado.index + 1)
    df_filtrado["Anotaciones"] = ""

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
# MÓDULO 5 – CERTIFICADOS COOMEVA
# ===========================
def limpiar_linea_coomeva(texto: str) -> str:
    if texto is None:
        return ""
    texto = texto.replace("\xa0", " ")
    texto = re.sub(r"[ \t]+", " ", texto)
    return texto.strip()


def extraer_texto_pdf_bytes(file_bytes) -> str:
    texto = []
    with fitz.open(stream=file_bytes, filetype="pdf") as doc:
        for page in doc:
            texto.append(page.get_text("text"))
    return "\n".join(texto)


def obtener_lineas_coomeva(texto: str):
    return [limpiar_linea_coomeva(x) for x in texto.splitlines() if limpiar_linea_coomeva(x)]


def buscar_indice_linea_coomeva(lineas, patron_exacto):
    for i, linea in enumerate(lineas):
        if linea == patron_exacto:
            return i
    return -1


def es_numero_puro_coomeva(linea: str) -> bool:
    return bool(re.fullmatch(r"\d{5,}", linea))


def es_fecha_coomeva(linea: str) -> bool:
    return bool(re.fullmatch(r"\d{2}/\d{2}/\d{4}", linea))


def es_cups_coomeva(linea: str) -> bool:
    return bool(re.fullmatch(r"\d{4}[A-Z]\d{2}", linea))


def es_valor_monetario_coomeva(linea: str) -> bool:
    return bool(re.fullmatch(r"[\d\.,]+", linea))


def extraer_carnet_coomeva(texto: str) -> str:
    m = re.search(r"([A-Z0-9]+)\s+Fecha\s+Generaci[oó]n\s*:", texto, flags=re.IGNORECASE)
    return m.group(1).strip() if m else ""


def a_numero_coomeva(valor: str):
    valor = limpiar_linea_coomeva(valor)
    if not valor:
        return None
    valor = valor.replace(",", "").replace("$", "").strip()
    try:
        return float(valor)
    except ValueError:
        return None


def extraer_datos_usuario_coomeva(texto: str, archivo: str) -> dict:
    lineas = obtener_lineas_coomeva(texto)

    idx_procedimiento = buscar_indice_linea_coomeva(lineas, "Procedimiento")
    if idx_procedimiento == -1:
        idx_procedimiento = next(
            (i for i, x in enumerate(lineas) if x.startswith("Procedimiento")),
            len(lineas)
        )

    head = lineas[:idx_procedimiento]
    texto_head = "\n".join(head)

    carnet = extraer_carnet_coomeva(texto_head)

    edad = ""
    plan_tarifario = ""
    fecha_generacion = ""

    idx_plan_tarifario = buscar_indice_linea_coomeva(head, "Plan Tarifario:")
    if idx_plan_tarifario != -1:
        siguientes = head[idx_plan_tarifario + 1: idx_plan_tarifario + 6]

        for val in siguientes:
            if not edad and re.fullmatch(r"\d{1,3}", val):
                edad = val
            elif not plan_tarifario and not es_numero_puro_coomeva(val) and not re.fullmatch(r"\d{1,3}", val) and not es_fecha_coomeva(val):
                plan_tarifario = val
            elif not fecha_generacion and es_fecha_coomeva(val):
                fecha_generacion = val

    plan = ""
    idx_carnet_label = buscar_indice_linea_coomeva(head, "Carnet:")
    if idx_carnet_label != -1 and idx_carnet_label + 1 < len(head):
        plan = head[idx_carnet_label + 1]

    nombre = ""
    apellidos = ""
    documento = ""
    programa = ""

    idx_documento = buscar_indice_linea_coomeva(head, "Documento")

    if idx_documento != -1:
        if idx_documento + 1 < len(head):
            apellidos = head[idx_documento + 1]

        if idx_documento + 4 < len(head):
            nombre = head[idx_documento + 4]

        if idx_documento + 5 < len(head):
            documento = head[idx_documento + 5]

        if idx_documento + 6 < len(head):
            programa = head[idx_documento + 6]

    return {
        "Archivo": archivo,
        "Fecha_carga": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Nombre": nombre,
        "Apellidos": apellidos,
        "Documento": documento,
        "Carnet": carnet,
        "Programa": programa,
        "Plan": plan,
        "Plan_Tarifario": plan_tarifario,
        "Fecha_Generacion": fecha_generacion,
        "Edad": edad,
    }


def extraer_tabla_procedimientos_coomeva(texto: str):
    lineas = obtener_lineas_coomeva(texto)

    idx_inicio = next(
        (i for i, x in enumerate(lineas) if x == "Pagar por Coomeva"),
        -1
    )

    idx_fin = next(
        (i for i, x in enumerate(lineas)
         if x == "DATOS USUARIO" or x.startswith("Confirmo que los tratamientos relacionados")),
        len(lineas)
    )

    if idx_inicio == -1 or idx_inicio + 1 >= idx_fin:
        return []

    body = lineas[idx_inicio + 1:idx_fin]

    detalles = []
    i = 0

    while i < len(body):
        if not body[i]:
            i += 1
            continue

        procedimiento_lines = []

        while i < len(body) and not es_cups_coomeva(body[i]):
            if es_valor_monetario_coomeva(body[i]):
                i += 1
                continue
            procedimiento_lines.append(body[i])
            i += 1

        if i >= len(body):
            break

        if not es_cups_coomeva(body[i]):
            i += 1
            continue

        cups = body[i]
        i += 1

        diagnostico_lines = []
        while i < len(body) and not es_numero_puro_coomeva(body[i]):
            if body[i] == "DATOS USUARIO" or body[i].startswith("Confirmo que los tratamientos relacionados"):
                break
            diagnostico_lines.append(body[i])
            i += 1

        if i >= len(body):
            break

        if not es_numero_puro_coomeva(body[i]):
            i += 1
            continue

        no_autorizacion = body[i]
        i += 1

        tarifa = body[i] if i < len(body) else ""
        i += 1

        copago = body[i] if i < len(body) else ""
        i += 1

        valor = body[i] if i < len(body) else ""
        i += 1

        detalles.append({
            "Procedimiento": " ".join(procedimiento_lines).strip(),
            "CUPS": cups.strip(),
            "Diagnostico": " ".join(diagnostico_lines).strip(),
            "No_de_Autorizacion": no_autorizacion.strip(),
            "Tarifa_Antes_de_Copago": a_numero_coomeva(tarifa),
            "Copago_con_IVA": a_numero_coomeva(copago),
            "Valor_Autorizado_a_Pagar_por_Coomeva": a_numero_coomeva(valor),
        })

        while i < len(body) and es_valor_monetario_coomeva(body[i]):
            i += 1

    return detalles


def transformar_certificados_coomeva(uploaded_files):
    registros = []
    errores = []
    archivos = 0

    for uf in uploaded_files:
        archivos += 1
        try:
            texto = extraer_texto_pdf_bytes(uf.getvalue())

            datos_usuario = extraer_datos_usuario_coomeva(texto, uf.name)
            detalles = extraer_tabla_procedimientos_coomeva(texto)

            if detalles:
                for detalle in detalles:
                    registros.append({**datos_usuario, **detalle})
            else:
                registros.append({
                    **datos_usuario,
                    "Procedimiento": "",
                    "CUPS": "",
                    "Diagnostico": "",
                    "No_de_Autorizacion": "",
                    "Tarifa_Antes_de_Copago": None,
                    "Copago_con_IVA": None,
                    "Valor_Autorizado_a_Pagar_por_Coomeva": None,
                })

        except Exception as e:
            errores.append({
                "Archivo": uf.name,
                "Error": str(e)
            })

    columnas_salida = [
        "Archivo",
        "Fecha_carga",
        "Nombre",
        "Apellidos",
        "Documento",
        "Carnet",
        "Programa",
        "Plan",
        "Plan_Tarifario",
        "Fecha_Generacion",
        "Edad",
        "Procedimiento",
        "CUPS",
        "Diagnostico",
        "No_de_Autorizacion",
        "Tarifa_Antes_de_Copago",
        "Copago_con_IVA",
        "Valor_Autorizado_a_Pagar_por_Coomeva",
    ]

    df = pd.DataFrame(registros)

    if df.empty:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Datos_Coomeva"
        ws.append(columnas_salida)
        if errores:
            ws_err = wb.create_sheet("Errores")
            ws_err.append(["Archivo", "Error"])
            for e in errores:
                ws_err.append([e["Archivo"], e["Error"]])
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out, 0, 0, pd.DataFrame(), pd.DataFrame(errores)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df[columnas_salida].to_excel(writer, sheet_name="Datos_Coomeva", index=False)

        ws = writer.sheets["Datos_Coomeva"]
        encabezados = {cell.value: cell.column for cell in ws[1]}

        columnas_moneda = [
            "Tarifa_Antes_de_Copago",
            "Copago_con_IVA",
            "Valor_Autorizado_a_Pagar_por_Coomeva",
        ]

        for nombre_col in columnas_moneda:
            col_idx = encabezados.get(nombre_col)
            if col_idx:
                for fila in range(2, ws.max_row + 1):
                    ws.cell(row=fila, column=col_idx).number_format = '$#,##0'

        autosize_columns(ws)

        if errores:
            df_err = pd.DataFrame(errores)
            df_err.to_excel(writer, sheet_name="Errores", index=False)
            autosize_columns(writer.sheets["Errores"])

    out.seek(0)
    return out, archivos, len(df), df, pd.DataFrame(errores)


# ===========================
# UI STREAMLIT
# ===========================
st.set_page_config("Denti Manager Web", layout="centered")
st.title("Denti Manager")

tab1, tab5, citas, firmas = st.tabs([
    "📄 PDF → Excel",
    "🦷 Coomeva",
    "📅 Citas",
    "✍️ Firmas",
])

with firmas:
    st.caption("Seleccione el tipo de firma que necesita.")
    tab6 = st.expander("🖊️ Firmas de pacientes · Wacom", expanded=True)
    tab2 = st.expander("📄 Firmar PDFs · firma del prestador")

with citas:
    st.caption("Abra el grupo de citas que desea procesar.")
    tab3 = st.expander("🚷 Citas canceladas")
    tab4 = st.expander("🔄 Citas inasistidas")

with tab1:
    files = st.file_uploader("PDFs", type=["pdf"], accept_multiple_files=True)
    if st.button("Procesar PDFs", key="btn_pdf", disabled=not files):
        out, a, f = transformar_archivos_a_excel(files)
        st.success(f"Archivos: {a} | Filas: {f}")
        st.download_button("Descargar Excel", out, f"PDF_{now_stamp()}.xlsx", key="dl_pdf")

with tab2:
    pdfs = st.file_uploader(
        "Subir PDFs para firmar",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdfs"
    )

    if st.button("✍️ Firmar PDFs", key="btn_firmar", disabled=not pdfs):
        z = firmar_pdfs_en_zip(pdfs)
        st.download_button(
            "Descargar ZIP Firmado",
            z,
            f"FIRMADOS_{now_stamp()}.zip",
            mime="application/zip",
            key="dl_zip"
        )

with tab3:
    f = st.file_uploader("Canceladas", type=["xls", "xlsx"], key="cancel")

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

with tab5:
    files = st.file_uploader(
        "Sube los certificados Coomeva en PDF",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"coomeva_pdfs_{st.session_state.coomeva_uploader_key}"
    )

    col1, col2 = st.columns([3, 1])

    with col1:
        if st.button("Procesar Certificados Coomeva", key="btn_coomeva", disabled=not files):
            out, archivos, filas, df_preview, df_err = transformar_certificados_coomeva(files)

            st.session_state.coomeva_resultado = out.getvalue()
            st.session_state.coomeva_preview = df_preview
            st.session_state.coomeva_errores = df_err
            st.session_state.coomeva_resumen = {
                "archivos": archivos,
                "filas": filas
            }

    with col2:
        if st.button(
            "🗑️ Limpiar archivos",
            key="btn_limpiar_coomeva",
            disabled=not files and st.session_state.coomeva_resultado is None
        ):
            limpiar_coomeva()
            st.rerun()

    if st.session_state.coomeva_resumen:
        st.success(
            f"Archivos procesados: {st.session_state.coomeva_resumen['archivos']} | "
            f"Filas generadas: {st.session_state.coomeva_resumen['filas']}"
        )

    if st.session_state.coomeva_preview is not None and not st.session_state.coomeva_preview.empty:
        st.dataframe(st.session_state.coomeva_preview.head())

    if st.session_state.coomeva_errores is not None and not st.session_state.coomeva_errores.empty:
        st.warning("Algunos archivos presentaron errores.")
        st.dataframe(st.session_state.coomeva_errores)

    if st.session_state.coomeva_resultado is not None:
        st.download_button(
            "Descargar Excel Coomeva",
            st.session_state.coomeva_resultado,
            f"COOMEVA_{now_stamp()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_coomeva"
        )


with tab6:
    # Firmas Wacom: mensajes de la tablet dirigidos al paciente.
    from firmas_paciente import render
    render()
