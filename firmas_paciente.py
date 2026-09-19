"""Firma manuscrita de pacientes: datos solo en la sesión, nunca en disco/cache."""
import base64
import hashlib
import io
import re
import unicodedata
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path

import fitz
from PIL import Image
import streamlit as st
import streamlit.components.v2 as components
from signing_flow import new_flow, consume_event, ready_to_save

ROOT = Path(__file__).parent
UI_REVISION = "mouse-test-v10"


def patient_details(original, page_index):
    """Read labeled values on the selected page without crossing into other columns."""
    with fitz.open(stream=original, filetype="pdf") as doc:
        page = doc[page_index]
        words = page.get_text("words")
        def value(label):
            matches = [fitz.Rect(w[:4]) for w in words if w[4].rstrip(':').casefold() == label.casefold()]
            if not matches:
                matches = page.search_for(label + ":") or (page.search_for(label) if ' ' in label else [])
            if not matches:
                return ""
            # Prefer the label in the patient column on the left.
            box = min(matches, key=lambda r: (r.x0, r.y0))
            middle = (box.y0+box.y1)/2
            row = sorted((w for w in words if abs((w[1]+w[3])/2-middle)<max(3, min(box.height,w[3]-w[1])*.55) and w[0]>=box.x1-1),key=lambda w:w[0])
            output=[]
            for w in row:
                if ':' in w[4] or w[4].lower() in ('plan','tarifario','razón','razon','nit','teléfono','telefono','edad','fecha'):
                    break
                output.append(w[4])
            return " ".join(output).strip()
        name = value('Nombre del paciente') or value('Nombre paciente') or value('Nombre')
        surnames = value('Apellidos')
        document = value('No. Documento') or value('Documento')
        # Preserve leading zeros. Reject extra columns or non-document text.
        if not re.fullmatch(r'(?:(?:CC|TI|CE|RC|PA|PEP|PPT)\s*)?[A-Z0-9.-]{4,25}',document,re.I):
            document = ""
        return " ".join(v for v in (name,surnames) if v),document


def signed_filename(name, document, signed_at):
    def clean(value):
        value = unicodedata.normalize('NFKD',value).encode('ascii','ignore').decode().upper()
        return re.sub(r'[^A-Z0-9]+','_',value).strip('_')[:100]
    return f"{clean(name)}_{clean(document)}_{signed_at.strftime('%Y-%m-%d_%H%M%S')}.pdf"
select_area = components.component(
    "pdf_signature_area",
    html='<canvas aria-label="Documento PDF: arrastre para marcar el espacio de firma"></canvas><p role="status"></p>',
    css='canvas {display:block;width:100%;cursor:crosshair;touch-action:none;background:white;border-radius:6px;} p {font-family:var(--st-font);color:var(--st-text-color);}',
    js=(ROOT / "pdf_signature_area.js").read_text(encoding="utf-8"),
)
capture = components.component(
    "wacom_dual_role_v8",
    html='''<div><button id="connect">Conectar pad de firma</button>
    <button id="clear">Repetir firma</button><button id="accept">Aceptar firma</button>
    <button id="disconnect">Desconectar pad de firma</button>
    <canvas id="mousepad" hidden aria-label="Firma de prueba con mouse" style="touch-action:none;max-height:240px"></canvas><p id="status" role="status" aria-live="polite">Conecte la tablet para comenzar.</p></div>''',
    css='''div {display:flex;flex-wrap:wrap;gap:.35rem;align-items:center;}
    button {font:500 14px var(--st-font, sans-serif);padding:.55rem .65rem;border-radius:8px;cursor:pointer;}
    canvas {width:100%; background:white; border:1px solid #aaa; border-radius:8px;}
    p {flex-basis:100%;margin:.35rem 0;font:14px var(--st-font,sans-serif);color:var(--st-text-color);}''',
    js=(ROOT / "wacom_capture.js").read_text(encoding="utf-8").replace("export default function", "function mountUsb") + "\n" + (ROOT / "mouse_capture.js").read_text(encoding="utf-8"),
)
save_dialog = components.component(
    "signature_save_as_v8",
    html='<button id="save">Descargar PDF firmado</button><p id="save_status" role="status"></p>',
    css='button {font:500 15px var(--st-font,sans-serif);padding:.65rem 1rem;border-radius:8px;cursor:pointer;} p {font:14px var(--st-font,sans-serif);color:var(--st-text-color);}',
    js=(ROOT / "save_signed_pdf.js").read_text(encoding="utf-8"),
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
        assert_unsigned_digital(doc)
        # Keep page rotation, existing content streams, images, forms and annotations.
        page.insert_image(box * page.derotation_matrix, stream=signature,
                          rotate=page.rotation, keep_proportion=True, overlay=True)
        return doc.tobytes(deflate=True)


def assert_unsigned_digital(doc):
    for xref in range(1, doc.xref_length()):
        if doc.xref_get_key(xref, 'ByteRange')[0] != 'null':
            raise ValueError('Este PDF contiene una firma digital certificada. No se modifica para evitar invalidarla. Use una copia sin certificación para añadir firmas manuscritas.')


def coomeva_box(original, page_index):
    with fitz.open(stream=original, filetype='pdf') as doc:
        page = doc[page_index]
        labels = page.search_for('Firma Prestador')
        if len(labels) != 1:
            raise ValueError("No se encontró un único campo 'Firma Prestador'. Seleccione la página correcta para Firma Coomeva.")
        label = labels[0]
        box = fitz.Rect(label.x0, label.y0 - 55, label.x0 + 140, label.y0 - 2)
        box = box * page.rotation_matrix
        if not page.rect.contains(box):
            raise ValueError('El campo de Firma Coomeva queda fuera de la página.')
        return tuple(box)


def add_provider_signature(original, page_index):
    """Add the saved provider signature only at an unambiguous labeled position."""
    return signed_pdf(original, (ROOT / 'firma.png').read_bytes(), page_index,
                      coomeva_box(original, page_index))


def validate_target(original, page_index, box, accepted, role):
    box = fitz.Rect(box)
    with fitz.open(stream=original, filetype='pdf') as doc:
        page = doc[page_index]
        if not page.rect.contains(box) or box.width < 8 or box.height < 3:
            raise ValueError('Marque un recuadro válido dentro de la página.')
        for other_role, other in accepted.items():
            if other_role != role and other['page'] == page_index and box.intersects(fitz.Rect(other['box'])):
                raise ValueError('Este espacio corresponde a otra firma aceptada. Marque un campo diferente.')
        # Check the actual rendered area, including scanned and vector signatures.
        # A occupied rectangle is rejected rather than covering existing ink.
        inner = box + (1, 1, -1, -1)
        pix = page.get_pixmap(clip=inner, matrix=fitz.Matrix(1, 1), colorspace=fitz.csGRAY, alpha=False)
        dark = sum(v < 190 for v in pix.samples)
        if dark > max(8, pix.width * pix.height * .003):
            raise ValueError('El recuadro contiene texto, trazos o una firma previa. Seleccione un espacio vacío para conservar el documento.')
    return tuple(box)


def compose_pdf(original, accepted):
    output = original
    for role in ('paciente', 'prestador', 'coomeva'):
        entry = accepted.get(role)
        if entry:
            output = signed_pdf(output, entry['png'], entry['page'], entry['box'])
    return output


def reset_document():
    epoch = st.session_state.get('patient_epoch', 0)
    for key in list(st.session_state):
        if key.startswith('patient_'):
            del st.session_state[key]
    st.session_state.patient_epoch = epoch + 1


def set_role(role):
    flow = st.session_state.patient_flow
    flow['role'] = role
    flow['target'] = None
    flow['phase'] = ('FIRMA_' + role.upper() + '_ACEPTADA' if role in flow['accepted']
                     else 'ESPERANDO_FIRMA_' + role.upper())


def prepare_document():
    epoch = st.session_state.get('patient_epoch', 0)
    uploaded = st.file_uploader('PDF del paciente', type=['pdf'], key=f'patient_upload_{epoch}')
    if not uploaded:
        return None
    original = uploaded.getvalue()
    if len(original) > 25 * 1024 * 1024:
        st.error('El PDF debe pesar como máximo 25 MB.')
        return None
    identity = hashlib.sha256(original).hexdigest()
    if st.session_state.get('patient_document') != identity or 'patient_flow' not in st.session_state:
        for key in list(st.session_state):
            if key.startswith('patient_') and not key.startswith('patient_upload') and key != 'patient_epoch':
                del st.session_state[key]
        st.session_state.patient_document = identity
        st.session_state.patient_flow = new_flow(identity)
    try:
        with fitz.open(stream=original, filetype='pdf') as doc:
            if doc.needs_pass or not len(doc):
                raise ValueError('Suba un PDF sin contraseña y con al menos una página.')
            assert_unsigned_digital(doc)
            if 'patient_detected' not in st.session_state:
                details = [patient_details(original, i) for i in range(len(doc))]
                # Keep both values from the same page to avoid mixing identities.
                st.session_state.patient_detected = max(details, key=lambda v: bool(v[0])+bool(v[1]))
            return original, uploaded.name, len(doc)
    except Exception as exc:
        st.error(f'No se puede preparar este documento: {exc}')
        return None


def render():
    title, reset = st.columns([3, 2], vertical_alignment='center')
    title.markdown('**Clínica DentiCenter | Firmas**')
    reset.button('Nuevo paciente / limpiar', key='patient_reset', on_click=reset_document)
    controls = st.container()
    with controls:
        st.markdown('**Control de Tablet**')
        pad_area = st.container()
        simulate = st.checkbox('Prueba con mouse', key='signature_test_mode', disabled=bool(st.session_state.get('patient_document')), help='Para cambiar de modo, use Nuevo paciente / limpiar. En prueba se pausa la captura USB.')
    if 'patient_flow' not in st.session_state and st.session_state.get('patient_signed_pdf'):
        st.info('Hay un PDF firmado en la sesión anterior. Guárdelo antes de comenzar con el nuevo flujo de firmas.')
        capture(key='wacom_connection_v8', data={'context':None,'role':'paciente','simulate':simulate},
                default={'event':None}, on_event_change=lambda: None)
        save_dialog(key='patient_previous_save', data={'pdf':base64.b64encode(st.session_state.patient_signed_pdf).decode(),
                                                      'filename':'documento_firmado_sesion_anterior.pdf'})
        st.button('Continuar con un nuevo documento', on_click=reset_document)
        return
    document = prepare_document()
    if not document:
        with pad_area:
            capture(key='wacom_connection_v8', data={'context': None, 'role': 'paciente','simulate':simulate},
                    default={'event': None}, on_event_change=lambda: None)
        return
    original, source_name, page_count = document
    flow = st.session_state.patient_flow
    event = st.session_state.get('wacom_connection_v8', {}).get('event')
    try:
        consume_event(flow, event, signature_png)
    except Exception as exc:
        st.error(f'No se aceptó la captura: {exc}. Seleccione Repetir firma.')
    role = flow['role']
    accepted = flow['accepted']
    current = accepted.get(role)
    pending = flow['phase'].endswith('_CAPTURADA')
    detected_name, detected_doc = st.session_state.patient_detected
    patient_name, patient_doc = detected_name, detected_doc
    if current:
        st.caption(f'Firma del {role} aceptada. Para cambiarla, pulse Repetir firma.')
    else:
        st.caption(f'Firma del {role}: marque un espacio vacío en el PDF y firme en el pad.')

    page_key = f'patient_page_{role}'
    page_index = st.number_input('Página', 1, page_count,
        current['page']+1 if current else page_count, key=page_key, disabled=bool(current) or pending)-1
    preview = compose_pdf(original, accepted)
    with fitz.open(stream=preview, filetype='pdf') as doc:
        page = doc[page_index]
        width, height = page.rect.width, page.rect.height
        pix = page.get_pixmap(matrix=fitz.Matrix(min(1.5,1200/width), min(1.5,1200/width)))
    area_key = f"patient_area_v8_{flow['identity']}_{role}_{page_index}"
    area = st.session_state.get(area_key, {}).get('selection')
    selection = select_area(key=area_key, data={'image':base64.b64encode(pix.tobytes('png')).decode(),
        'selection':area, 'locked':bool(current) or pending, 'role':role},
        default={'selection':None}, on_selection_change=lambda: None).selection
    target = None
    if current:
        target = {k:current[k] for k in ('page','box','context')}
    elif selection:
        try:
            if not isinstance(selection,list) or len(selection)!=4 or not all(isinstance(v,(int,float)) and 0<=v<=1 for v in selection):
                raise ValueError('Marque de nuevo el campo de firma.')
            box = validate_target(original, page_index, (selection[0]*width,selection[1]*height,
                selection[2]*width,selection[3]*height), accepted, role)
            context = hashlib.sha256(str((flow['identity'],role,page_index,box,flow['generation'],simulate)).encode()).hexdigest()
            target = dict(page=page_index, box=box, context=context)
        except Exception as exc:
            st.warning(str(exc))
    flow['target'] = target
    if target and not current and not pending:
        flow['phase'] = 'ESPERANDO_FIRMA_' + role.upper()
    with pad_area:
        capture(key='wacom_connection_v8', data={'context':target['context'] if target else None,
                'role':role, 'completed':bool(current),'simulate':simulate}, default={'event':event}, on_event_change=lambda: None)

    if not detected_name or not detected_doc:
        st.caption('Complete los datos para nombrar el PDF.')
        name_col, doc_col = st.columns([3, 2])
        patient_name = name_col.text_input('Nombre y apellidos', value=detected_name, key='patient_filename_name')
        patient_doc = doc_col.text_input('No. Documento', value=detected_doc, key='patient_filename_document')
    coomeva, provider, patient = st.columns(3)
    if coomeva.button('Firma Coomeva', key='patient_coomeva', use_container_width=True,
                     disabled='paciente' not in accepted or pending or role not in accepted or 'prestador' in accepted or 'coomeva' in accepted,
                     help='Añade el sello guardado después de aceptar la firma del paciente.'):
        try:
            box = coomeva_box(original, page_index)
            validate_target(original, page_index, box, accepted, 'coomeva')
            accepted['coomeva'] = dict(page=page_index,box=box,png=(ROOT/'firma.png').read_bytes())
            st.rerun()
        except Exception as exc:
            st.warning(str(exc))
    provider.button('Firma prestador', on_click=set_role, args=('prestador',), use_container_width=True,
             disabled='paciente' not in accepted or pending or role == 'prestador' or 'coomeva' in accepted,
             key='patient_choose_provider', help='Captura la firma del prestador en otro campo.')
    patient.button('Firma paciente', on_click=set_role, args=('paciente',), use_container_width=True,
             disabled=role == 'paciente' or pending, key='patient_choose_patient',
             help='Vuelve al modo paciente. Conserva las firmas aceptadas.')
    if 'coomeva' in accepted:
        st.caption('Sello Coomeva añadido.')
    ready = ready_to_save(flow)
    if ready and patient_name.strip() and patient_doc.strip():
        output = compose_pdf(original, accepted)
        timestamp = max(v['signed_at'] for v in accepted.values() if 'signed_at' in v)
        filename = ('PRUEBA_' if simulate else '') + signed_filename(patient_name, patient_doc, timestamp)
        st.session_state.patient_signed_pdf = output
        flow['document_state'] = 'DOCUMENTO_LISTO_PARA_GUARDAR'
        save_dialog(key='patient_save_as', data={'pdf':base64.b64encode(output).decode(), 'filename':filename})
        st.caption(filename)
    else:
        flow['document_state'] = flow['phase']
        st.button('Descargar PDF firmado', disabled=True, key='patient_save_disabled')
        st.caption('Acepte la firma y complete los datos para guardar.')

