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
.rib .tb .grp.cc{justify-content:center;padding:0 3px}
.rib .tb .grp.cc>.gl{display:flex;align-items:center;justify-content:center;min-height:50px;margin:4px 0;padding:0 9px;font-size:12px;color:#222;border:1px solid #d6d6d6;border-radius:4px;background:#fafafa;cursor:pointer}
.rib .tb .grp.cc>.gl::after{content:" \\25BE";margin-left:4px}
.rib .tb .grp.cc.open>.gl{background:var(--hl);border-color:#a9b9d6}
.rib .tb .grp.cc.open>.gb{display:flex;position:fixed;z-index:950;background:#f3f3f3;border:1px solid #b9b9b9;box-shadow:0 4px 14px rgba(0,0,0,.3);padding:6px;border-radius:4px;flex-wrap:wrap;max-width:calc(100vw - 12px);max-height:60vh;overflow:auto}
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
  rib.addEventListener('click',e=>{const l=e.target.closest('.grp.cc>.gl');if(l){const g=l.parentElement,was=g.classList.contains('open');closeAll();if(!was){g.classList.add('open');const gb=g.querySelector(':scope>.gb'),r=l.getBoundingClientRect();gb.style.top=r.bottom+2+'px';gb.style.left='0px';const w=gb.offsetWidth;gb.style.left=Math.max(4,Math.min(r.left,innerWidth-w-6))+'px'}e.stopPropagation();return}
    const b=e.target.closest('.grp.cc.open .gb button');if(b&&!b.closest('.gal'))setTimeout(closeAll,60)},true);
  document.addEventListener('pointerdown',e=>{if(!e.target.closest('.grp.cc.open'))closeAll()},true);
  addEventListener('resize',sched);addEventListener('orientationchange',()=>setTimeout(sched,200));
  window.ribFit=fitAll;sched();setTimeout(sched,400);setTimeout(sched,1500);return true}
function boot(){if(!wire())setTimeout(boot,200)}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',boot);else boot();
})();
