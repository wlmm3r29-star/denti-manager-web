# Firmas de paciente y prestador

1. Cargar PDF. Se leen el nombre y el documento; se pueden corregir para nombrar el archivo.
2. Marcar un espacio vacío para el paciente y firmar en el pad.
3. Aceptar en el pad o en Control de Tablet. La pantalla queda en blanco y la firma permanece en la sesión.
4. Opcional: Firmar prestador, marcar su espacio (en la misma página o en otra), firmar y aceptar.
5. Para el formato Coomeva, usar Firma Coomeva: añade el sello existente, sin capturar otra firma. No se combina con una captura de prestador para evitar dos firmas en el mismo campo.
6. Descargar PDF firmado abre Guardar como en Chrome/Edge de escritorio: elegir carpeta y nombre. Cancelar no borra el documento ni inicia una descarga automática.

Control de Tablet contiene Conectar pad de firma, Repetir firma, Aceptar firma y Desconectar pad de firma. Repetir afecta solamente al rol seleccionado; para repetir paciente después de prestador, seleccionar Firma del paciente y después Repetir firma. Ambas firmas aceptadas tienen estado independiente.

El pad permanece conectado entre las ejecuciones de Streamlit, con reconexión automática si el navegador conserva el permiso. Desconectar es una acción explícita del operador. Cerrar la página termina la sesión; guardar antes de cerrarla.

El PDF original no se rasteriza ni se reemplazan sus imágenes. Se conserva la rotación y se añaden las firmas sobre el contenido existente. Se rechazan campos ocupados para no cubrir textos o firmas; en un documento escaneado puede ser necesario marcar un área más precisa. Los PDF con firmas digitales certificadas se rechazan para evitar invalidar su certificación. Las firmas manuscritas existentes sí se conservan.

Las capturas pendientes bloquean el guardado tanto en el servidor como inmediatamente en el navegador. Los eventos llevan documento, rol y contexto del recuadro; eventos antiguos no se aplican a otra firma.

Archivos: firmas_paciente.py (interfaz y PDF), signing_flow.py (estados), wacom_capture.js (pad), pdf_signature_area.js (recuadro) y save_signed_pdf.js (Guardar como). No se añaden dependencias. Esta revisión no activa almacenamiento externo de datos.
