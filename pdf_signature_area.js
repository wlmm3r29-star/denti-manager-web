export default function({parentElement,data,setStateValue}) {
  const canvas=parentElement.querySelector('canvas'), ctx=canvas.getContext('2d');
  const label=parentElement.querySelector('p');
  let start=null, rect=data.selection, ready=false, active=true;
  const image=new Image();
  function draw(){
    if(!ready)return;
    ctx.clearRect(0,0,canvas.width,canvas.height);ctx.drawImage(image,0,0);
    if(rect){
      const [x,y,x2,y2]=rect;
      ctx.fillStyle='rgba(0,120,190,.12)';ctx.strokeStyle='#0078be';ctx.lineWidth=3;
      ctx.fillRect(x*canvas.width,y*canvas.height,(x2-x)*canvas.width,(y2-y)*canvas.height);
      ctx.strokeRect(x*canvas.width,y*canvas.height,(x2-x)*canvas.width,(y2-y)*canvas.height);
    }
    label.textContent=data.locked?'Campo bloqueado durante la captura o tras aceptar. Use Control de Tablet para aceptar o repetir.':rect?'Espacio de firma seleccionado. Puede marcar otro recuadro para cambiarlo.':'Marque el espacio de la firma arrastrando sobre el documento.';
  }
  image.onload=()=>{if(!active)return;canvas.width=image.naturalWidth;canvas.height=image.naturalHeight;ready=true;draw();};
  image.src='data:image/png;base64,'+data.image;
  function position(e){const b=canvas.getBoundingClientRect();return [Math.max(0,Math.min(1,(e.clientX-b.left)/b.width)),Math.max(0,Math.min(1,(e.clientY-b.top)/b.height))];}
  canvas.onpointerdown=e=>{if(data.locked || !ready || e.button!==0)return;e.preventDefault();start=position(e);canvas.setPointerCapture(e.pointerId);};
  canvas.onpointermove=e=>{if(!start)return;const end=position(e);rect=[Math.min(start[0],end[0]),Math.min(start[1],end[1]),Math.max(start[0],end[0]),Math.max(start[1],end[1])];draw();};
  canvas.onpointerup=e=>{
    if(!start)return;
    const end=position(e);rect=[Math.min(start[0],end[0]),Math.min(start[1],end[1]),Math.max(start[0],end[0]),Math.max(start[1],end[1])];start=null;
    canvas.releasePointerCapture(e.pointerId);
    if((rect[2]-rect[0])*canvas.width<12 || (rect[3]-rect[1])*canvas.height<6){rect=data.selection;draw();label.textContent='Arrastre para dibujar un recuadro más grande.';return;}
    draw();setStateValue('selection',rect);
  };
  canvas.onpointercancel=()=>{start=null;rect=data.selection;draw();};
  return ()=>{active=false;image.onload=null;canvas.onpointerdown=canvas.onpointermove=canvas.onpointerup=canvas.onpointercancel=null;};
}
