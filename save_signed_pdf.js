export default function({parentElement,data}) {
  const button=parentElement.querySelector('#save'), status=parentElement.querySelector('#save_status');
  let active=true, saving=false;
  function pending(){
    const pad=window[Symbol.for('denti.wacom.connection.v8')];
    return pad && (pad.repeatRequested || (pad.desired && !pad.accepted));
  }
  function update(){button.disabled=saving || !data.pdf || pending();}
  update();const timer=setInterval(update,100);
  button.onclick=async()=>{
    if(saving || !data.pdf || pending())return;
    if(!window.showSaveFilePicker){status.textContent='Para elegir la carpeta y el nombre al guardar, abra esta página directamente en Chrome o Edge de escritorio.';return;}
    saving=true;update();
    let writable;
    try{
      // Invoke directly during the user gesture; never fall back to an automatic download.
      const handle=await window.showSaveFilePicker({suggestedName:data.filename,
        types:[{description:'Documento PDF firmado',accept:{'application/pdf':['.pdf']}}]});
      if(!active || pending())throw new Error('La firma cambió. Cierre esta ventana y vuelva a descargar el documento actualizado.');
      const bytes=Uint8Array.from(atob(data.pdf),c=>c.charCodeAt(0));
      writable=await handle.createWritable();await writable.write(bytes);await writable.close();writable=null;
      status.textContent='PDF guardado en la ubicación seleccionada.';
    }catch(error){
      if(writable){try{await writable.abort();}catch{}}
      status.textContent=error.name==='AbortError'?'Guardado cancelado. El PDF sigue disponible en esta sesión.':
        error.name==='SecurityError'?'Abra Denti Manager directamente en una pestaña de Chrome o Edge para elegir dónde guardar.':`No se pudo guardar: ${error.message}`;
    }finally{saving=false;if(active)update();}
  };
  return ()=>{active=false;clearInterval(timer);button.onclick=null;};
}
