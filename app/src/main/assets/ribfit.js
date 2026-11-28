/* One-row ribbon: groups that do not fit collapse into "Label ▾" drop-downs (Excel + Word). */
(function(){
const PRI=['clipboard','styles','cells','editing','rules','font','paragraph','proofing','voice','sensitivity','show','window','macros'];
const css=document.createElement('style');css.textContent=`
.rib .tb{flex-wrap:nowrap!important;max-height:none!important;overflow-x:auto!important;overflow-y:hidden!important;align-items:flex-start!important}
.rib .tb .grp{flex:0 0 auto;border-bottom:0!important;min-width:0}
.rib .tb .gb{min-height:60px!important;max-height:64px}
.rib .tb .grp.cc.open>.gb{max-height:60vh}
.rib .tb .col{gap:0!important}
.rib .tb .col>.r button{height:21px}
.rib .tb button.big{height:60px}
.rib .tb .gl{padding:1px 0 2px}
.rib .tb .grp.cc>.gb{display:none}
.rib .tb .grp.cc{justify-content:center;padding:0 2px;border-right:0}
.rib .tb .grp.cc>.gl{display:flex;align-items:center;justify-content:center;gap:6px;height:34px;min-height:0;margin:auto 0;padding:0 11px 0 13px;font-size:12.5px;font-weight:500;letter-spacing:.1px;color:#2d2d2d;border:1px solid transparent;border-radius:7px;background:transparent;cursor:pointer;transition:background .12s,border-color .12s,box-shadow .12s}
.rib .tb .grp.cc>.gl *{display:none}
.rib .tb .grp.cc>.gl::after{content:"";width:6px;height:6px;border-right:1.6px solid #6b6b6b;border-bottom:1.6px solid #6b6b6b;transform:rotate(45deg) translateY(-2px);transition:transform .15s}
.rib .tb .grp.cc>.gl:active{background:#e7e7e7}
.rib .tb .grp.cc.open>.gl{background:#fff;border-color:#d4d4d4;box-shadow:0 1px 2px rgba(0,0,0,.08);color:var(--hlt)}
.rib .tb .grp.cc.open>.gl::after{transform:rotate(225deg) translateY(-2px);border-color:var(--hlt)}
.rib .tb .grp.cc.open>.gb{display:flex;align-items:flex-start;gap:8px 12px;position:fixed;z-index:950;background:#fff;border:1px solid #d9d9d9;box-shadow:0 10px 28px rgba(0,0,0,.22),0 2px 6px rgba(0,0,0,.08);padding:14px 16px 12px;border-radius:10px;flex-wrap:wrap;max-width:calc(100vw - 12px);max-height:62vh;overflow:auto;animation:rfpop .12s ease-out}
@keyframes rfpop{from{opacity:0;transform:translateY(-4px)}to{opacity:1;transform:none}}
.rib .tb .grp.cc.open>.gb::before{content:attr(data-gl);flex:0 0 100%;font-size:10.5px;font-weight:600;letter-spacing:.7px;text-transform:uppercase;color:#8a8a8a;margin:-4px 0 2px;padding-bottom:6px;border-bottom:1px solid #ececec}
.rib .tb .grp.cc.open>.gb .col{align-self:flex-start;align-items:stretch;gap:8px;justify-content:flex-start;min-width:140px}
.rib .tb .grp.cc.open>.gb .r{gap:4px;justify-content:flex-start;align-items:center;flex-wrap:wrap}
.rib .tb .grp.cc.open>.gb .col>.r:only-child,.rib .tb .grp.cc.open>.gb .col>button{width:100%}
.rib .tb .grp.cc.open>.gb button{height:32px;padding:0 12px;border-radius:6px;font-size:13px;justify-content:flex-start;gap:8px;min-width:32px}
.rib .tb .grp.cc.open>.gb button:hover{background:#f0f0f0!important}
.rib .tb .grp.cc.open>.gb button:active{background:var(--hl)!important}
.rib .tb .grp.cc.open>.gb button.big{height:auto;min-height:64px;min-width:68px;justify-content:center;padding:6px 8px;border-radius:8px}
.rib .tb .grp.cc.open>.gb select,.rib .tb .grp.cc.open>.gb input[type=text],.rib .tb .grp.cc.open>.gb input[type=number],.rib .tb .grp.cc.open>.gb input:not([type]){height:32px;box-sizing:border-box;border:1px solid #cfcfcf;border-radius:6px;padding:0 10px;font-size:13px;background:#fff;color:#222;min-width:130px;max-width:none;width:100%;margin:0}
.rib .tb .grp.cc.open>.gb select{appearance:auto}
.rib .tb .grp.cc.open>.gb select:focus,.rib .tb .grp.cc.open>.gb input:focus{outline:2px solid var(--hl);border-color:var(--hlt)}
.rib .tb .grp.cc.open>.gb label{font-size:12px;color:#555;display:block;margin:0 0 3px}
.rib .tb .grp.cc.open>.gb .gal{max-height:none;max-width:min(380px,86vw)}
.rib .tb .grp.cc.open>.gb button.btn-ic{width:32px;padding:0;justify-content:center}
`;document.head.appendChild(css);
let busy=false;
function lab(g){const l=g.querySelector(':scope>.gl');return l?l.textContent.trim().toLowerCase():''}
function closeAll(){document.querySelectorAll('.grp.cc.open').forEach(g=>g.classList.remove('open'))}
function fit(tb){
  const gs=[...tb.querySelectorAll(':scope>.grp')];if(!gs.length)return;
  gs.forEach(g=>g.classList.remove('cc','open'));
  const order=[];PRI.forEach(p=>gs.forEach(g=>{if(lab(g)===p&&!order.includes(g))order.push(g)}));
  gs.slice().reverse().forEach(g=>{if(!order.includes(g))order.push(g)});
  for(const g of order){if(tb.scrollWidth<=tb.clientWidth+1)break;g.classList.add('cc')}
}
function fitAll(){if(busy)return;busy=true;try{closeAll();document.querySelectorAll('.rib .tb').forEach(tb=>{if(tb.offsetParent!==null||tb.getClientRects().length)fit(tb)})}finally{busy=false}}
let raf=0;function sched(){cancelAnimationFrame(raf);raf=requestAnimationFrame(fitAll)}
function wire(){
  const rib=document.querySelector('.rib');if(!rib)return false;
  new MutationObserver(ms=>{if(busy)return;for(const m of ms){const t=m.target;if(m.type==='childList'&&(t.classList.contains('tb')||t.classList.contains('rib')||t.id==='rib')){sched();return}if(m.type==='attributes'&&t.classList.contains('tb')){sched();return}}}).observe(rib,{childList:true,subtree:true,attributes:true,attributeFilter:['class','style']});
  rib.addEventListener('click',e=>{const l=e.target.closest('.grp.cc>.gl');if(l){const g=l.parentElement,was=g.classList.contains('open');closeAll();if(!was){g.classList.add('open');const gb=g.querySelector(':scope>.gb');gb.setAttribute('data-gl',lab(g).replace(/[\u2198\u25BE]/g,'').trim()),r=l.getBoundingClientRect();gb.style.top=r.bottom+2+'px';gb.style.left='0px';const w=gb.offsetWidth;gb.style.left=Math.max(4,Math.min(r.left,innerWidth-w-6))+'px'}e.stopPropagation();return}
    const b=e.target.closest('.grp.cc.open .gb button');if(b&&!b.closest('.gal'))setTimeout(closeAll,60)},true);
  document.addEventListener('pointerdown',e=>{if(!e.target.closest('.grp.cc.open'))closeAll()},true);
  addEventListener('resize',sched);addEventListener('orientationchange',()=>setTimeout(sched,200));
  window.ribFit=fitAll;sched();setTimeout(sched,400);setTimeout(sched,1500);return true}
function boot(){if(!wire())setTimeout(boot,200)}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',boot);else boot();
})();
