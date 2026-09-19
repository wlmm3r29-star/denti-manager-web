// Protocol reference: pabloko/Wacom-STU-WebHID (MIT).
function createConnection(){
  const m={device:null,saved:[],context:null,desired:null,accepted:false,points:0,previous:null,down:false,busy:false,view:null,paused:false,role:'paciente',desiredRole:'paciente',repeatRequested:false,message:'Conecte el pad de firma una vez para autorizar el navegador.',queue:Promise.resolve()};
  const ink=document.createElement('canvas');ink.width=800;ink.height=480;const ctx=ink.getContext('2d');
  const screen=document.createElement('canvas');screen.width=800;screen.height=480;const sc=screen.getContext('2d');
  const send=(id,b)=>m.device.sendFeatureReport(id,new Uint8Array(b));
  const update=()=>m.view?.update(m),message=s=>{m.message=s;update();};
  function clearInk(){ctx.clearRect(0,0,800,480);m.points=0;m.previous=null;m.down=false;}
  async function close(){
    if(m.device){m.device.removeEventListener('inputreport',pen);
      for(const [id,b] of [[33,[0]],[32,[0]],...m.saved]){try{await send(id,b);}catch{}}
      try{await m.device.close();}catch{}
    }m.device=null;m.saved=[];m.context=null;clearInk();
  }
  function enqueue(action){m.queue=m.queue.then(async()=>{m.busy=true;update();
    try{await action();}catch(e){await close();message(`No se pudo conectar o capturar: ${e.message}. Cierre otros programas de firma y pulse Conectar.`);}
    finally{m.busy=false;update();}});return m.queue;}
  async function showScreen(){
    if(!m.context || m.accepted){
      await send(33,[0]);
      await send(46,[255,255,255]);
      await send(32,[0]);
      return;
    }
    sc.fillStyle='white';sc.fillRect(0,0,800,480);
    if(m.context && !m.accepted){
      sc.fillStyle='#142d40';sc.fillRect(0,0,800,75);
      sc.fillStyle='white';sc.font='bold 26px Arial';sc.fillText(`Clínica DentiCenter | Firma del ${m.role}`,25,48);
      sc.strokeStyle='#668899';sc.lineWidth=2;sc.strokeRect(25,95,750,285);
      sc.fillStyle='#555';sc.font='18px Arial';sc.fillText(m.accepted?'Gracias. Su firma ha sido recibida.':m.context?`Firma ${m.role}: firme y pulse ACEPTAR`:'Por favor, espere la indicación de recepción para firmar.',35,365);
      for(const [x,color,label] of [[25,'#526777','REPETIR'],[415,'#08785c','ACEPTAR']]){
        sc.fillStyle=m.context && !m.accepted?color:'#b0b6bb';sc.fillRect(x,405,360,55);sc.fillStyle='white';sc.font='bold 24px Arial';sc.fillText(label,x+120,440);
      }
    }
    const rgba=sc.getImageData(0,0,800,480).data,bgr=new Uint8Array(800*480*3);
    for(let p=0,o=0;p<rgba.length;p+=4){bgr[o++]=rgba[p+2];bgr[o++]=rgba[p+1];bgr[o++]=rgba[p];}
    message('Pad conectado. Preparando pantalla…');await send(37,[4]);
    for(let offset=0;offset<bgr.length;offset+=253){const chunk=bgr.subarray(offset,offset+253),packet=new Uint8Array(255);packet[0]=chunk.length;packet.set(chunk,2);await send(38,packet);}
    await send(39,[0]);
  }
  async function sync(){
    if(!m.device || m.context===m.desired)return;
    await send(33,[0]);clearInk();m.context=m.desired;m.role=m.desiredRole;m.accepted=!!m.completed;m.repeatRequested=false;
    await showScreen();await send(33,[m.context && !m.accepted?1:0]);
    message(m.accepted?`Firma del ${m.role} aceptada. Pantalla en blanco. Puede continuar o seleccionar Repetir firma.`:m.context?`Firma ${m.role}: firme y pulse ACEPTAR.`:'Pad conectado. Pantalla en blanco hasta seleccionar el campo de firma.');
  }
  async function open(device){
    if(m.device || !device || m.paused)return;
    m.device=device;await device.open();const cap=await device.receiveFeatureReport(9);
    if(cap.byteLength<12 || cap.getUint8(0)!==9 || cap.getUint16(7)!==800 || cap.getUint16(9)!==480)throw new Error('Tablet no compatible');
    m.maxX=cap.getUint16(1);m.maxY=cap.getUint16(3);
    if(!m.maxX || !m.maxY || cap.getUint16(5)!==1023)throw new Error('Características STU no válidas');
    for(const [id,size] of [[14,2],[33,2],[45,5],[46,4],[43,3]]){
      const v=await device.receiveFeatureReport(id);
      if(v.byteLength!==size || v.getUint8(0)!==id)throw new Error('No se pudo leer la configuración');
      m.saved.push([id,Array.from(new Uint8Array(v.buffer,v.byteOffset+1,v.byteLength-1))]);
    }
    await send(33,[0]);await send(14,[0]);await send(46,[255,255,255]);await send(45,[20,20,20,2]);await send(43,[3,0]);
    device.addEventListener('inputreport',pen);m.context=undefined;await sync();
  }
  m.autoConnect=()=>enqueue(async()=>{
    if(!navigator.hid || m.device || m.paused)return;
    const devices=(await navigator.hid.getDevices()).filter(d=>d.vendorId===0x056a && d.productId===0x00a8);
    if(devices.length===1)await open(devices[0]);else if(devices.length>1)message('Hay varios pads. Pulse Conectar pad de firma para elegir uno.');
  });
  m.connect=async()=>{
    if(!navigator.hid || !window.isSecureContext){message('Use Chrome o Edge de escritorio en HTTPS.');return;}
    m.paused=false;
    try{const selected=await navigator.hid.requestDevice({filters:[{vendorId:0x056a,productId:0x00a8}]});if(selected[0])await enqueue(()=>open(selected[0]));}
    catch(e){message(`No se autorizó la conexión: ${e.message}`);}
  };
  const emit=(kind,extra={})=>m.view?.emit({kind,context:m.context,role:m.role,id:crypto.randomUUID(),...extra});
  m.repeat=()=>enqueue(async()=>{
    if(!m.context || m.context!==m.desired || m.repeatRequested)return;
    if(m.accepted){
      m.repeatRequested=true;
      if(m.device){await send(33,[0]);await showScreen();}
      emit('repeat');message(`Repitiendo firma del ${m.role}. Preparando pad…`);return;
    }
    if(!m.device){message('Conecte el pad para repetir la firma.');return;}
    await send(33,[0]);clearInk();await showScreen();await send(33,[1]);
    emit('ready');message(`Repitiendo firma del ${m.role}. Firme y pulse ACEPTAR.`);
  });
  m.accept=()=>enqueue(async()=>{
    if(!m.device || !m.context || m.context!==m.desired || m.points<2 || m.accepted)return;
    const pixels=ctx.getImageData(0,0,800,480).data;let minX=800,minY=480,maxX=0,maxY=0;
    for(let y=0;y<480;y++)for(let x=0;x<800;x++)if(pixels[(y*800+x)*4+3]){minX=Math.min(minX,x);maxX=Math.max(maxX,x);minY=Math.min(minY,y);maxY=Math.max(maxY,y);}
    if(maxX-minX<8 || maxY-minY<3){message('Firma incompleta. Haga un trazo más amplio o pulse REPETIR.');return;}
    const payload={kind:'accepted',id:crypto.randomUUID(),role:m.role,context:m.context,png:ink.toDataURL('image/png'),acceptedAt:new Date().toISOString()};m.accepted=true;m.view?.emit(payload);
    await send(33,[0]);clearInk();await showScreen();message(`Firma del ${m.role} aceptada. Pantalla en blanco. Puede continuar o seleccionar Repetir firma.`);
  });
  function pen(e){
    if(!m.view || m.busy || !m.context || m.context!==m.desired || m.accepted || ![1,52].includes(e.reportId))return;
    const d=e.data;if(d.byteLength<(e.reportId===1?6:10))return;
    if((d.getUint16(0)&0x9000)!==0x9000){m.previous=null;m.down=false;return;}
    const x=d.getUint16(2)/m.maxX*800,y=d.getUint16(4)/m.maxY*480;
    if(y>=405 && y<=460){m.previous=null;if(!m.down){m.down=true;if(x>=25 && x<=385)void m.repeat();else if(x>=415 && x<=775)void m.accept();}return;}
    if(x<30 || x>770 || y<100 || y>350){m.previous=null;return;}
    ctx.strokeStyle='#141414';ctx.fillStyle='#141414';ctx.lineWidth=2.5;ctx.lineCap='round';ctx.lineJoin='round';
    if(m.previous){ctx.beginPath();ctx.moveTo(...m.previous);ctx.lineTo(x,y);ctx.stroke();}else{ctx.beginPath();ctx.arc(x,y,1.25,0,Math.PI*2);ctx.fill();}
    m.previous=[x,y];m.points++;if(m.points===1)emit('captured');update();
  }
  m.disconnect=()=>{m.paused=true;return enqueue(async()=>{await close();message('Pad desconectado por el operador.');});};
  m.setContext=context=>{m.desired=context;return enqueue(sync);};
  navigator.hid?.addEventListener('connect',()=>{void m.autoConnect();});
  navigator.hid?.addEventListener('disconnect',e=>{if(e.device===m.device)void enqueue(async()=>{await close();message('Pad desconectado. Se reconectará al volver a enchufarlo; la firma aceptada se conserva.');});});
  window.addEventListener('pagehide',()=>{void close();});return m;
}
export default function({parentElement,data,setStateValue}){
  const slot=Symbol.for('denti.wacom.connection.v8');
  if(!window[slot]){
    const previous=window[Symbol.for('denti.wacom.connection.v7')] || window[Symbol.for('denti.wacom.connection.v6')] || window[Symbol.for('denti.wacom.connection.v5')] || window[Symbol.for('denti.wacom.connection.v4')] || window[Symbol.for('denti.wacom.connection.v3')] || window[Symbol.for('denti.wacom.connection.v2')];
    const next=createConnection();
    if(previous)next.queue=previous.disconnect();
    window[slot]=next;
  }
  const m=window[slot];m.completed=!!data.completed;m.desiredRole=data.role==='prestador'?'prestador':'paciente';
  const q=s=>parentElement.querySelector(s);
  const view={emit:p=>setStateValue('event',p),update:s=>{
    q('#status').textContent=s.message;q('#connect').disabled=s.busy || !!s.device;
    q('#clear').disabled=s.busy || !s.context || s.repeatRequested;
    q('#clear').title=`Repetir únicamente la firma del ${s.role}`;
    q('#accept').disabled=s.busy || !s.device || !s.context || s.points<2 || s.accepted || s.repeatRequested;
    q('#disconnect').disabled=s.busy || !s.device;
  }};
  clearTimeout(m.detachTimer);m.view=view;view.update(m);
  q('#connect').onclick=m.connect;q('#clear').onclick=m.repeat;q('#accept').onclick=m.accept;q('#disconnect').onclick=m.disconnect;
  void m.setContext(data.context??null);void m.autoConnect();
  return ()=>{if(m.view===view){m.view=null;m.detachTimer=setTimeout(()=>{if(!m.view)void m.setContext(null);},1000);}};
}
