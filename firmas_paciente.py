"""Firma manuscrita de pacientes: datos solo en la sesión, nunca en disco/cache."""
import base64
import hashlib
import io
from pathlib import Path

import fitz
from PIL import Image
import streamlit as st
import streamlit.components.v2 as components

ROOT = Path(__file__).parent
capture = components.component(
    "wacom_stu540",
    html='''<div><button id="connect">Conectar Wacom STU-540</button>
    <button id="clear">Borrar / repetir</button><button id="accept">Aceptar firma</button>
    <button id="disconnect">Desconectar</button>
    <p id="status" role="status" aria-live="polite">Conecte la tablet para comenzar.</p>
    <canvas width="800" height="480" aria-label="Vista de la firma capturada"></canvas></div>''',
    css='''button {padding: .6rem; margin: .2rem; border-radius: 8px; cursor: pointer;}
    canvas {width:100%; background:white; border:1px solid #aaa; border-radius:8px;}
    p {font-family:var(--st-font); color:var(--st-text-color);}''',
    js=(ROOT / "wacom_capture.js").read_text(encoding="utf-8"),
)


def signature_png(value):
    if not isinstance(value, str) or not value.startswith("data:image/png;base64,") or len(value) > 3_000_000:
        raise ValueError("Firma no válida. Capture de nuevo.")
    raw = base64.b64decode(value.split(",", 1)[1], validate=True)
    with Image.open(io.BytesIO(raw)) as image:
        if image.format != "PNG" or image.size != (800, 480):
            raise ValueError("Tamaño de firma no válido.")
        image = image.convert("RGBA")
        bounds = image.getchannel("A").getbbox()
        if not bounds or bounds[2] - bounds[0] < 8 or bounds[3] - bounds[1] < 3:
            raise ValueError("La firma está vacía o incompleta.")
        image = image.crop(bounds)
        out = io.BytesIO()
        image.save(out, format="PNG")
        return out.getvalue()


def signed_pdf(original, signature, page_number, rect):
    with fitz.open(stream=original, filetype="pdf") as doc:
        if doc.needs_pass or not 0 <= page_number < len(doc):
            raise ValueError("PDF protegido o página no válida.")
        page = doc[page_number]
        box = fitz.Rect(rect)
        if box.is_empty or box.is_infinite or not page.rect.contains(box):
            raise ValueError("La firma debe quedar dentro de la página.")
        # UI coordinates follow the displayed, rotated page. Normalize this page
        # without altering its appearance before placing the image.
        if page.rotation:
            page.remove_rotation()
        page.insert_image(box, stream=signature, keep_proportion=True, overlay=True)
        return doc.tobytes(garbage=3, deflate=True)


def render():
    st.subheader("Firmas paciente")
    st.write("Cargue el documento del paciente, capture su firma y revise el PDF antes de descargarlo.")
    st.info("Use Chrome o Edge en el computador conectado por USB a la STU-540. Cierre otros programas de firma. El navegador pedirá permiso para conectar la tablet.")
    st.caption("Captura sin suscripción ni SDK de pago. El PDF y la firma se procesan en la sesión de Streamlit; descargue el resultado antes de salir. Firma manuscrita como imagen, sin certificado digital.")
    epoch = st.session_state.get("patient_epoch", 0)
    if st.button("Nuevo paciente / limpiar", key="patient_reset"):
        for key in list(st.session_state):
            if key.startswith("patient_"):
                del st.session_state[key]
        st.session_state.patient_epoch = epoch + 1
        st.rerun()
    uploaded = st.file_uploader("PDF del paciente", type=["pdf"], key=f"patient_upload_{epoch}")
    if not uploaded:
        return
    original = uploaded.getvalue()
    if len(original) > 25 * 1024 * 1024:
        st.error("El PDF debe pesar como máximo 25 MB.")
        return
    identity = hashlib.sha256(original).hexdigest()
    if st.session_state.get("patient_document") != identity:
        for key in list(st.session_state):
            if key.startswith("patient_") and not key.startswith("patient_upload") and key != "patient_epoch":
                del st.session_state[key]
        st.session_state.patient_document = identity
    try:
        with fitz.open(stream=original, filetype="pdf") as doc:
            if doc.needs_pass or len(doc) == 0:
                st.error("Suba un PDF sin contraseña y con al menos una página.")
                return
            page_index = st.number_input("Página que va a firmar", 1, len(doc), len(doc), key="patient_page") - 1
            page = doc[page_index]
            width, height = page.rect.width, page.rect.height
            candidates = []
            for widget in page.widgets() or []:
                if widget.field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE and not widget.field_value:
                    candidates.append((f"Campo: {widget.field_label or widget.field_name}", widget.rect * page.rotation_matrix))
            for label in ("Firma Paciente", "Firma del Paciente", "Firma Usuario", "Firma del Usuario"):
                for found in page.search_for(label):
                    found = found * page.rotation_matrix
                    candidate = fitz.Rect(found.x0, max(0,found.y0-55), min(width,found.x0+150),found.y0)
                    if not candidate.is_empty:
                        candidates.append((f"{label} ({len(candidates)+1})",candidate))
            choice = st.selectbox("Campo de firma", ["Ubicación manual"]+[name for name,_ in candidates], key=f"patient_field_{page_index}")
            st.caption("Ubicación en porcentajes de la página, medidos desde la esquina superior izquierda.")
            cols = st.columns(2)
            x = cols[0].slider("Izquierda (%)", 0, 90, 10, key="patient_x")
            y = cols[1].slider("Arriba (%)", 0, 90, 75, key="patient_y")
            w = cols[0].slider("Ancho (%)", 5, 100 - x, min(30, 100 - x), key="patient_w")
            h = cols[1].slider("Alto (%)", 3, 100 - y, min(12, 100 - y), key="patient_h")
            box = fitz.Rect(x * width / 100, y * height / 100, (x+w)*width/100, (y+h)*height/100)
            if choice != "Ubicación manual":
                box = next(rect for name,rect in candidates if name == choice)
                st.caption("Se utiliza el campo seleccionado. Para ajustar los porcentajes, elija Ubicación manual.")
            if page.rotation:
                page.remove_rotation()
            page.draw_rect(box, color=(0, .5, .8), width=1.5)
            pix = page.get_pixmap(matrix=fitz.Matrix(min(1.5, 1200/width), min(1.5, 1200/width)))
            st.image(pix.tobytes("png"), caption="El recuadro azul indica dónde quedará la firma.")
        accepted = st.session_state.get("patient_signature")
        if accepted:
            if st.button("Repetir firma", key="patient_repeat"):
                del st.session_state["patient_signature"]
                st.session_state.patient_capture_generation = st.session_state.get("patient_capture_generation", 0) + 1
                st.rerun()
        else:
            generation = st.session_state.get("patient_capture_generation", 0)
            result = capture(key=f"patient_capture_{epoch}_{identity}_{generation}", data={"document": identity},
                             default={"signature": None}, on_signature_change=lambda: None)
            if not result.signature:
                return
            accepted = signature_png(result.signature)
            st.session_state.patient_signature = accepted
            st.rerun()
        signature = accepted
        output = signed_pdf(original, signature, page_index, box)
        with fitz.open(stream=output, filetype="pdf") as preview:
            p = preview[page_index]
            st.image(p.get_pixmap(matrix=fitz.Matrix(min(1.5,1200/p.rect.width), min(1.5,1200/p.rect.width))).tobytes("png"), caption="Vista previa del PDF firmado")
        # Confirmation is tied to source, signature, page and coordinates (not PDF serialization IDs).
        confirmation = hashlib.sha256(original + signature + str((page_index, tuple(box))).encode()).hexdigest()
        confirmed = st.checkbox("Revisé el documento del paciente y la ubicación de su firma", key=f"patient_confirm_{confirmation}")
        st.download_button("Descargar PDF firmado", output, Path(uploaded.name).stem + "_firmado.pdf",
                           mime="application/pdf", disabled=not confirmed, key="patient_download")
    except Exception as exc:
        st.error(f"No se pudo preparar la firma: {exc}")
