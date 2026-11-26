// Shared helpers: save/open files in the Android app and in a normal browser
window.RB = {
  saveBlob(name, blob, mime){
    const done = ()=>{};
    if (window.Android && Android.saveFile) {
      const fr = new FileReader();
      fr.onload = ()=>{ const b64 = fr.result.split(',')[1]; Android.saveFile(name, b64, mime||blob.type||'application/octet-stream'); };
      fr.readAsDataURL(blob);
    } else {
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
      document.body.appendChild(a); a.click(); setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove();},500);
    }
  },
  pickFile(accept, cb, asText){
    const i = document.createElement('input'); i.type='file'; if(accept) i.accept = accept;
    i.onchange = ()=>{ const f = i.files[0]; if(!f) return; const fr = new FileReader();
      fr.onload = ()=>cb(f.name, fr.result); asText ? fr.readAsText(f) : fr.readAsArrayBuffer(f); };
    i.click();
  },
  ask(title, def, cb){
    const m = document.getElementById('rbModal') || (()=>{ const d=document.createElement('div'); d.id='rbModal'; d.className='modal';
      d.innerHTML='<div class="box"><h3></h3><input type="text" id="rbIn"><div class="row"><button class="sec" id="rbNo">Cancel</button><button id="rbOk">OK</button></div></div>'; document.body.appendChild(d); return d; })();
    m.querySelector('h3').textContent = title; const inp = m.querySelector('#rbIn'); inp.value = def||'';
    m.style.display='flex'; inp.focus();
    m.querySelector('#rbNo').onclick = ()=>{m.style.display='none'};
    m.querySelector('#rbOk').onclick = ()=>{m.style.display='none'; cb(inp.value)};
  },
  home(){ location.href='index.html'; }
};
