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
  home(){
    const X=window.__exit; if(!X||window.__discard){ location.href='index.html'; return; }
    let empty=false; try{ empty=X.empty&&X.empty(); }catch(e){}
    if(empty){ X.keys.forEach(k=>{try{localStorage.removeItem(k)}catch(e){}}); location.href='index.html'; return; }
    let m=document.getElementById('rbExit'); if(m) m.remove();
    m=document.createElement('div'); m.id='rbExit'; m.className='modal'; m.style.cssText='display:flex;z-index:2000';
    m.innerHTML='<div class="box"><h3>Save this file?</h3><div style="margin:6px 0 12px;font-size:13px;opacity:.85">Do you want to save before closing? If you don\'t save, the next time you open it you will start with a blank file.</div><div class="row" style="flex-wrap:wrap;gap:8px"><button class="sec" id="rbxC">Cancel</button><button class="sec" id="rbxN">Don\'t save</button><button id="rbxS">Save</button></div></div>';
    document.body.appendChild(m);
    m.querySelector('#rbxC').onclick=()=>m.remove();
    m.querySelector('#rbxN').onclick=()=>{ window.__discard=true; X.keys.forEach(k=>{try{localStorage.removeItem(k)}catch(e){}}); location.href='index.html'; };
    m.querySelector('#rbxS').onclick=async()=>{ m.remove(); try{ await X.save(); }catch(e){} setTimeout(()=>{location.href='index.html'},400); };
  }
};

(function(){const o=Storage.prototype.setItem;Storage.prototype.setItem=function(k,v){if(window.__discard&&window.__exit&&window.__exit.keys.indexOf(k)>=0)return;return o.apply(this,arguments)}})();
