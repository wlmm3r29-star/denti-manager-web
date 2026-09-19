// Optional simulator. The USB implementation above remains unchanged.
function mountMouse({parentElement,data,setStateValue}) {
  const slot=Symbol.for('denti.mouse.connection.v1');
  const m=window[slot] ||= {canvas:document.createElement('canvas'),context:null,accepted:false,points:0};
  if(!m.canvas.width || m.canvas.width!==800){m.canvas.width=800;m.canvas.height=480;}
  m.active=true;m.desired=data.context??null;m.role=data.role||'paciente';
  const q=s=>parentElement.querySelector(s), canvas=q('#mousepad'), ctx=canvas.getContext('2d');
  const ink=m.canvas.getContext('2d');canvas.hidden=false;canvas.width=800;canvas.height=480;
  let previous=null,down=false;
  function clear(){ink.clearRect(0,0,800,480);m.points=0;previous=null;}
  if(m.context!==m.desired){clear();m.context=m.desired;m.accepted=!!data.completed;m.repeatRequested=false;}
  function emit(kind,extra={}){setStateValue('event',{kind,id:crypto.randomUUID(),context:m.context,role:m.role,...extra});}
  function paint(){
    ctx.fillStyle='white';ctx.fillRect(0,0,800,480);ctx.drawImage(m.canvas,0,0);
    q('#connect').disabled=true;q('#disconnect').disabled=true;
    q('#clear').disabled=!m.context||m.repeatRequested;
    q('#accept').disabled=!m.context||m.accepted||m.points<2;
    q('#status').textContent=!m.context?'Prueba con mouse: seleccione un campo en el PDF.':m.accepted?
      `Firma de prueba del ${m.role} aceptada.`:`Prueba: dibuje la firma del ${m.role} aquí y pulse Aceptar firma.`;
  }
  canvas.onpointerdown=e=>{
    if(e.button!==0||!m.context||m.accepted||m.repeatRequested)return;
    down=true;previous=null;canvas.setPointerCapture(e.pointerId);draw(e);
  };
  function draw(e){
    if(!down)return;e.preventDefault();const b=canvas.getBoundingClientRect();
    const x=Math.max(1,Math.min(799,(e.clientX-b.left)*800/b.width));
    const y=Math.max(1,Math.min(479,(e.clientY-b.top)*480/b.height));
    ink.strokeStyle='#141414';ink.fillStyle='#141414';ink.lineWidth=3;ink.lineCap='round';
    if(previous){ink.beginPath();ink.moveTo(...previous);ink.lineTo(x,y);ink.stroke();}
    else{ink.beginPath();ink.arc(x,y,1.5,0,Math.PI*2);ink.fill();}
    previous=[x,y];m.points++;paint();
  }
  canvas.onpointermove=draw;
  canvas.onpointerup=e=>{if(!down)return;draw(e);down=false;previous=null;canvas.releasePointerCapture(e.pointerId);emit('captured');};
  canvas.onpointercancel=()=>{down=false;previous=null;if(m.points)emit('captured');};
  q('#clear').onclick=()=>{
    if(m.accepted){m.repeatRequested=true;emit('repeat');paint();return;}
    clear();emit('ready');paint();
  };
  q('#accept').onclick=()=>{
    if(!m.context||m.accepted||m.points<2)return;
    const a=ink.getImageData(0,0,800,480).data;let minX=800,minY=480,maxX=0,maxY=0;
    for(let y=0;y<480;y++)for(let x=0;x<800;x++)if(a[(y*800+x)*4+3]){minX=Math.min(minX,x);maxX=Math.max(maxX,x);minY=Math.min(minY,y);maxY=Math.max(maxY,y);}
    if(maxX-minX<8||maxY-minY<3){q('#status').textContent='Haga un trazo más amplio antes de aceptar.';return;}
    m.accepted=true;emit('accepted',{png:m.canvas.toDataURL('image/png')});paint();
  };
  paint();return ()=>{canvas.onpointerdown=canvas.onpointermove=canvas.onpointerup=canvas.onpointercancel=null;};
}
export default function(args){
  const usb=window[Symbol.for('denti.wacom.connection.v8')];
  const mouse=window[Symbol.for('denti.mouse.connection.v1')];
  if(args.data.simulate){
    if(usb&&!usb.simulationPaused){usb.simulationPaused=true;usb.pausedBeforeSimulation=usb.paused;void usb.disconnect();}
    return mountMouse(args);
  }
  if(mouse){mouse.active=false;mouse.desired=null;}
  if(usb?.simulationPaused){usb.simulationPaused=false;usb.paused=usb.pausedBeforeSimulation;}
  args.parentElement.querySelector('#mousepad').hidden=true;
  return mountUsb(args);
}
