# Firmas paciente con Wacom STU-540

1. Abra la web directamente en Chrome o Edge de escritorio (HTTPS).
2. Conecte una STU-540 en modo USB HID y cierre sign pro PDF u otras aplicaciones de captura.
3. En **Firmas paciente**, cargue cualquier PDF y elija la página. Arrastre directamente sobre el documento para dibujar el espacio de la firma. Dibuje otro recuadro si quiere cambiarlo.
4. Pulse **Conectar Wacom STU-540** y elija la tablet en el permiso del navegador.
5. Espere a que termine de cargar la pantalla. Firme dentro del recuadro.
6. Use **REVISAR**, **BORRAR** o **CANCELAR** en la tablet. Después de revisar pulse **ACEPTAR**. Los controles también están en el computador.
7. Revise el PDF final en la web, marque la confirmación y descargue el PDF firmado.
8. Pulse **Nuevo paciente / limpiar** antes de pasar al siguiente paciente.

## Configuración y límites

- No requiere suscripción de captura, sign pro PDF ni Signature SDK. El protocolo de referencia es MIT; se incluye su licencia.
- No modifica firmware, modo USB/serial ni restaura valores de fábrica. Guarda los ajustes de escritura, tinta, fondo y brillo; los restaura al cerrar la conexión. Desenchufar o cerrar abruptamente el navegador puede impedir la restauración; vuelva a conectar para verificar.
- Pantalla personalizada temporal de 800 × 480. La transferencia USB de la pantalla puede tardar; espere el mensaje de conexión antes de escribir.
- STU-540 únicamente. Firefox y Safari no ofrecen la conexión WebHID necesaria. Si una política del navegador/alojamiento bloquea HID, se muestra el error; no se desactivan protecciones.
- Una firma y una página por operación. La selección visual funciona sobre cualquier PDF, incluidos escaneados, sin depender de rótulos ni de campos de formulario.
- Se añade una imagen manuscrita al PDF. No se emite un certificado digital ni se rellena criptográficamente un campo de firma digital.
- PDF hasta 25 MB, sin contraseña. Los documentos y la imagen viajan al servidor de Streamlit de la sesión para generar el resultado. Este módulo no escribe archivos de pacientes en disco, bases de datos ni cachés compartidas. Descargar antes de cerrar.
- Requiere Streamlit 1.52.0 (componente v2 sin iframe) y PyMuPDF 1.26.7 (coordenadas de páginas giradas).

## Validación

Pruebas automáticas: inicio de los seis módulos; rechazo de firmas vacías; inserción y límites de página; rotaciones 0/90/180/270; captura, revisión, borrado y restauración con HID simulado.
La verificación de hardware real debe hacerse con una firma de prueba en la STU-540 conectada al navegador.

Referencias:
- https://github.com/pabloko/Wacom-STU-WebHID
- https://docs.streamlit.io/develop/api-reference/custom-components
- https://github.com/Wacom-Developer/stu-sdk-samples/blob/master/GETTING-STARTED.md
