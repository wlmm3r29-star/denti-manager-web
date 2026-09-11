// STU-540 protocol adapted from the local wacom_stu540.py implementation.
// Reference: pabloko/Wacom-STU-WebHID (MIT; see WACOM_WEBHID_LICENSE.txt).
export default function({parentElement, setStateValue}) {
  const q = s => parentElement.querySelector(s);
  const canvas = q('canvas'), ctx = canvas.getContext('2d');
  const status = text => { q('#status').textContent = text; };
  let device = null, saved = [], maxX = 0, maxY = 0, previous = null;
  let points = 0, busy = false, active = true, capturing = false, reviewing = false, buttonDown = false;
  const screen = document.createElement('canvas'); screen.width=800; screen.height=480;
  const sc = screen.getContext('2d');
  async function showScreen() {
    sc.fillStyle='#ffffff'; sc.fillRect(0,0,800,480);
    sc.fillStyle='#142d40'; sc.fillRect(0,0,800,75);
    sc.fillStyle='#ffffff'; sc.font='bold 28px Arial';
    sc.fillText(reviewing ? 'Revise su firma antes de aceptar' : 'Denti Manager | Firma del paciente',25,48);
    sc.strokeStyle='#668899'; sc.lineWidth=2; sc.strokeRect(25,95,750,285);
    if (reviewing) sc.drawImage(canvas,0,0);
    else {sc.fillStyle='#555555';sc.font='18px Arial';sc.fillText('Firme dentro del recuadro con el lapiz',35,365);}
    for (const [x,color,label] of [[25,'#555555','CANCELAR'],[285,'#526777','BORRAR'],[545,'#08785c',reviewing?'ACEPTAR':'REVISAR']]) {
      sc.fillStyle=color;sc.fillRect(x,405,230,55);sc.fillStyle='white';sc.font='bold 22px Arial';sc.fillText(label,x+42,440);
    }
    const rgba=sc.getImageData(0,0,800,480).data;
    const bgr=new Uint8Array(800*480*3);
    for(let p=0,o=0;p<rgba.length;p+=4){bgr[o++]=rgba[p+2];bgr[o++]=rgba[p+1];bgr[o++]=rgba[p];}
    status('Preparando pantalla de la Wacom… espere antes de firmar.');
    await send(0x25,[4]);
    // Feature reports have a fixed payload: length (little endian) + 253 bytes.
    // Await every packet so writeImageEnd can never overtake the image data.
    for(let offset=0;offset<bgr.length;offset+=253){
      if(!active) throw new Error('Captura cerrada');
      const chunk=bgr.subarray(offset,offset+253), packet=new Uint8Array(255);
      packet[0]=chunk.length;packet.set(chunk,2);await send(0x26,packet);
    }
    await send(0x27,[0]);
  }
  const send = (id, bytes) => device.sendFeatureReport(id, new Uint8Array(bytes));
  function buttons() {
    q('#connect').disabled = busy || !!device;
    q('#clear').disabled = busy || !device;
    q('#accept').disabled = busy || !device || points < 2 || !capturing;
    q('#accept').textContent = reviewing ? 'Aceptar firma' : 'Revisar firma';
    q('#disconnect').disabled = busy || !device;
  }
  function resetInk() { ctx.clearRect(0,0,800,480); previous = null; points = 0; buttons(); }
  async function close() {
    capturing = false;
    if (device) {
      device.removeEventListener('inputreport', pen);
      for (const [id, bytes] of [[0x21,[0]], [0x20,[0]], ...saved]) {
        try { await send(id, bytes); } catch (_) { /* Unplugged device */ }
      }
      try { await device.close(); } catch (_) { /* Already closed */ }
    }
    device = null; saved = []; previous = null;
    if (active) buttons();
  }
  function pen(event) {
    if (!capturing || ![1,0x34].includes(event.reportId)) return;
    const d = event.data;
    if (d.byteLength < (event.reportId === 1 ? 6 : 10)) return;
    const flags = d.getUint16(0), xRaw = d.getUint16(2), yRaw = d.getUint16(4);
    if (xRaw > maxX || yRaw > maxY) { previous = null; return; }
    if ((flags & 0x9000) !== 0x9000) { previous = null; buttonDown = false; return; }
    const x = xRaw/maxX*800, y = yRaw/maxY*480;
    if (y>=405 && y<=460) {
      previous=null;
      if (!buttonDown && !busy) {
        buttonDown=true;
        if(x>=25 && x<=255) q('#disconnect').onclick();
        else if(x>=285 && x<=515) q('#clear').onclick();
        else if(x>=545 && x<=775 && points>=2) q('#accept').onclick();
      }
      return;
    }
    if(reviewing || x<30 || x>770 || y<100 || y>350){previous=null;return;}
    ctx.strokeStyle = '#141414'; ctx.fillStyle = '#141414'; ctx.lineWidth = 2.5;
    ctx.lineCap = 'round'; ctx.lineJoin = 'round';
    if (previous) { ctx.beginPath(); ctx.moveTo(...previous); ctx.lineTo(x,y); ctx.stroke(); }
    else { ctx.beginPath(); ctx.arc(x,y,1.25,0,Math.PI*2); ctx.fill(); }
    previous = [x,y]; points++; buttons();
  }
  async function run(action) {
    if (busy) return;
    busy = true; buttons();
    try { await action(); }
    catch (error) { await close(); status(`No se pudo completar: ${error.message}. Cierre otros programas de firma y vuelva a conectar.`); }
    finally { busy = false; if (active) buttons(); }
  }
  q('#connect').onclick = () => run(async () => {
    if (!navigator.hid || !window.isSecureContext) throw new Error('Abra la página HTTPS en Chrome o Edge de escritorio');
    if (document.permissionsPolicy?.allowsFeature && !document.permissionsPolicy.allowsFeature('hid'))
      throw new Error('El alojamiento bloqueó el permiso USB HID');
    const selected = await navigator.hid.requestDevice({filters:[{vendorId:0x056a,productId:0x00a8}]});
    if (!selected.length || !active) return;
    device = selected[0]; await device.open();
    const cap = await device.receiveFeatureReport(9);
    if (cap.byteLength < 12 || cap.getUint8(0) !== 9) throw new Error('Características STU no válidas');
    maxX = cap.getUint16(1); maxY = cap.getUint16(3);
    if (!maxX || !maxY || cap.getUint16(5) !== 1023 || cap.getUint16(7)!==800 || cap.getUint16(9)!==480) throw new Error('La tablet no es compatible');
    for (const [id, length] of [[0x0e,2],[0x21,2],[0x2d,5],[0x2e,4],[0x2b,3]]) {
      const value = await device.receiveFeatureReport(id);
      if (value.byteLength !== length || value.getUint8(0) !== id) throw new Error('No se pudo leer la configuración');
      saved.push([id, Array.from(new Uint8Array(value.buffer,value.byteOffset+1,value.byteLength-1))]);
    }
    await send(0x21,[0]); await send(0x0e,[0]); await send(0x2e,[255,255,255]);
    await send(0x2d,[20,20,20,2]); await send(0x2b,[3,0]); await send(0x20,[0]);
    reviewing=false; resetInk(); await showScreen(); device.addEventListener('inputreport', pen);
    capturing = true; await send(0x21,[1]);
    status('Wacom conectada. El paciente puede firmar en la tablet.');
  });
  q('#clear').onclick = () => run(async () => {
    capturing = false; reviewing=false; await send(0x21,[0]); await send(0x20,[0]); resetInk(); await showScreen();
    capturing = true; await send(0x21,[1]);
    status('Pantalla limpia. Capture una nueva firma.');
  });
  q('#accept').onclick = () => run(async () => {
    if (points < 2) throw new Error('Firme antes de aceptar');
    if (!reviewing) {
      capturing=false; reviewing=true; await send(0x21,[0]); await showScreen(); capturing=true;
      status('Revise la firma en la tablet. Pulse ACEPTAR o BORRAR para repetir.');return;
    }
    capturing = false;
    const png = canvas.toDataURL('image/png');
    await close(); status('Firma aceptada. Revise el PDF debajo.');
    setStateValue('signature', png);
  });
  q('#disconnect').onclick = () => run(async () => {await close(); resetInk(); status('Wacom desconectada.');});
  const unplug = e => { if (e.device === device) void run(async () => {await close(); resetInk(); status('Se desconectó el USB. Conecte de nuevo y repita la firma.');}); };
  navigator.hid?.addEventListener('disconnect', unplug);
  buttons();
  return () => { active = false; navigator.hid?.removeEventListener('disconnect', unplug); void close(); };
}
