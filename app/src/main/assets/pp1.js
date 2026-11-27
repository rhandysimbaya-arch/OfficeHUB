/* PowerPoint-style editor core: model, themes, layouts, shape catalogue, rendering */
const $=id=>document.getElementById(id);
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
let deck,cur=0,sel=[],name='Presentation1',scale=1,uid=1,uz=1,view='normal',editing=null,multiMode=false,tcell=null;
const PV={grid:false,snap:true,guides:true,notes:false,fpane:null,ruler:false};
const S=()=>deck.slides[cur];
const nid=()=>'i'+(uid++)+Math.floor(Math.random()*1e4);
const selId=()=>sel[sel.length-1]||null;
const itemById=(id,s)=>(s||S()).items.find(i=>i.id===id);
const selItems=()=>sel.map(id=>itemById(id)).filter(Boolean);
const one=()=>sel.length===1?itemById(sel[0]):null;
/* ---------------- themes ---------------- */
const THEMES={
 'Office':{bg:'#ffffff',tc:'#1b1b1b',bc:'#333333',major:'Calibri Light',minor:'Calibri',acc:['#c43e1c','#2b579a','#70ad47','#ffc000','#5b9bd5','#7f7f7f'],deco:'none'},
 'Facet':{bg:'#ffffff',tc:'#2c4a2c',bc:'#333333',major:'Trebuchet MS',minor:'Trebuchet MS',acc:['#90c226','#54a021','#e6b91e','#e76618','#c42f1a','#918655'],deco:'leftbar'},
 'Ion':{bg:'#1e2a38',tc:'#ffffff',bc:'#dfe7f0',major:'Century Gothic',minor:'Century Gothic',acc:['#b01513','#ea6312','#e6b729','#6aac90','#5f9c9d','#9e5e9b'],deco:'topbar'},
 'Retrospect':{bg:'#f5efe2',tc:'#3a2d1a',bc:'#4a3d2a',major:'Calibri Light',minor:'Calibri',acc:['#e48312','#bd582c','#865640','#9b8357','#c2bc80','#94a088'],deco:'band'},
 'Slice':{bg:'#1b1b1b',tc:'#ffffff',bc:'#e6e6e6',major:'Arial',minor:'Arial',acc:['#052f61','#a50e82','#14967c','#6a9e1f','#e87d37','#c62324'],deco:'corner'},
 'Integral':{bg:'#ffffff',tc:'#1f497d',bc:'#333333',major:'Tw Cen MT',minor:'Tw Cen MT',acc:['#1cade4','#2683c6','#27ced7','#42ba97','#3e8853','#62a39f'],deco:'leftbar'},
 'Wisp':{bg:'#fbf6ec',tc:'#6a3a1a',bc:'#4a3a2a',major:'Century Gothic',minor:'Century Gothic',acc:['#a53010','#de7e18','#9f8351','#728653','#92aa4c','#6aac91'],deco:'corner'},
 'Dividend':{bg:'#ffffff',tc:'#4d1434',bc:'#333333',major:'Gill Sans MT',minor:'Gill Sans MT',acc:['#4d1434','#903163','#b2324b','#969fa7','#66b1ce','#40619d'],deco:'band'}
};
const FONTS=['Calibri','Calibri Light','Arial','Times New Roman','Verdana','Georgia','Tahoma','Segoe UI','Cambria','Impact','Comic Sans MS','Century Gothic','Trebuchet MS','Courier New','Garamond','Tw Cen MT','Gill Sans MT'];
function theme(){const t=deck.theme||(deck.theme={name:'Office',variant:0});const b=THEMES[t.name]||THEMES.Office;return b}
function themeCols(){const b=theme(),v=deck.theme.variant||0;let bg=b.bg,tc=b.tc,bc=b.bc;const dark=c=>{const n=parseInt(c.slice(1),16);return ((n>>16&255)*.3+(n>>8&255)*.59+(n&255)*.11)<110};
 if(v===1&&!dark(bg)){bg='#1f1f1f';tc='#ffffff';bc='#e8e8e8'}else if(v===1&&dark(bg)){bg='#ffffff';tc='#1b1b1b';bc='#333333'}
 else if(v===2){bg=mix(b.acc[0],'#ffffff',.88);if(dark(b.bg)){tc='#1b1b1b';bc='#333333'}}else if(v===3){bg=b.acc[0];tc='#ffffff';bc='#ffffff'}
 return{bg,tc,bc,acc:b.acc,major:b.major,minor:b.minor,deco:b.deco}}
function mix(a,b,t){const p=c=>[1,3,5].map(i=>parseInt(c.slice(i,i+2),16));const x=p(a),y=p(b);return '#'+x.map((v,i)=>Math.round(v*(1-t)+y[i]*t).toString(16).padStart(2,'0')).join('')}
function col(c){if(!c)return c;const m=/^acc(\d)$/.exec(c);if(m)return themeCols().acc[+m[1]-1]||'#c43e1c';if(c==='tx1')return themeCols().bc;if(c==='bg1')return themeCols().bg;return c}
function hex(c){return (col(c)||'#000000').replace('#','').toUpperCase()}
function applyTheme(n,v){deck.theme={name:n,variant:v==null?(deck.theme?deck.theme.variant:0):v};const T=themeCols();deck.slides.forEach(s=>{if(!s.bgc)s.bg=T.bg;s.items.forEach(i=>{if(i.ph&&i.tcol===undefined){}})});}
/* ---------------- layouts ---------------- */
const LAYOUTS={
 title:{n:'Title Slide',p:[['ctrTitle',.083,.26,.834,.22,{size:54,align:'center',valign:'b'}],['sub',.167,.5,.666,.13,{size:26,align:'center',valign:'t'}]]},
 content:{n:'Title and Content',p:[['title',.0625,.05,.875,.16,{size:38}],['body',.0625,.24,.875,.64,{size:26,bullet:'dot'}]]},
 section:{n:'Section Header',p:[['title',.0625,.30,.875,.2,{size:50,bold:true,valign:'b'}],['body',.0625,.52,.875,.12,{size:24,valign:'t',muted:1}]]},
 two:{n:'Two Content',p:[['title',.0625,.05,.875,.16,{size:38}],['body',.0625,.24,.43,.64,{size:24,bullet:'dot'}],['body2',.5075,.24,.43,.64,{size:24,bullet:'dot'}]]},
 compare:{n:'Comparison',p:[['title',.0625,.05,.875,.16,{size:38}],['hdr',.0625,.24,.43,.09,{size:24,bold:true,valign:'b'}],['body',.0625,.34,.43,.54,{size:22,bullet:'dot'}],['hdr2',.5075,.24,.43,.09,{size:24,bold:true,valign:'b'}],['body2',.5075,.34,.43,.54,{size:22,bullet:'dot'}]]},
 titleonly:{n:'Title Only',p:[['title',.0625,.05,.875,.16,{size:38}]]},
 blank:{n:'Blank',p:[]},
 caption:{n:'Content with Caption',p:[['title',.0625,.08,.31,.24,{size:30,bold:true,valign:'b'}],['body',.42,.08,.52,.8,{size:26,bullet:'dot'}],['sub',.0625,.34,.31,.5,{size:16,valign:'t'}]]},
 picture:{n:'Picture with Caption',p:[['pic',.0625,.07,.875,.5,{}],['title',.0625,.6,.875,.11,{size:30,bold:true,valign:'b'}],['sub',.0625,.72,.875,.14,{size:18,valign:'t'}]]}
};
const PROMPT={title:'Click to add title',ctrTitle:'Click to add title',sub:'Click to add subtitle',body:'Click to add text',body2:'Click to add text',hdr:'Click to add text',hdr2:'Click to add text',pic:'Click icon to add picture'};
function mkPh(kind,x,y,w,h,o){const it=Object.assign({id:nid(),type:'text',ph:kind,x:Math.round(x*deck.w),y:Math.round(y*deck.h),w:Math.round(w*deck.w),h:Math.round(h*deck.h),text:''},o||{});if(kind==='pic')it.type='phpic';return it}
function applyLayout(s,key,reset){const L=LAYOUTS[key]||LAYOUTS.blank;s.layout=key;const old=s.items.filter(i=>i.ph);const used=new Set();
 L.p.forEach(([k,x,y,w,h,o])=>{let it=old.find(i=>i.ph===k&&!used.has(i.id));if(!it&&(k==='title'||k==='ctrTitle'))it=old.find(i=>/^(title|ctrTitle)$/.test(i.ph)&&!used.has(i.id));if(!it&&/^(sub|body)$/.test(k))it=old.find(i=>/^(sub|body)$/.test(i.ph)&&!used.has(i.id)&&i.ph!==k&&!L.p.some(q=>q[0]===i.ph));
  if(it){used.add(it.id);const g=mkPh(k,x,y,w,h,o);if(reset){Object.keys(it).forEach(q=>{if(!/^(id|text|html|ph)$/.test(q))delete it[q]});Object.assign(it,g,{id:it.id,text:it.text,html:it.html})}else{it.ph=k;it.x=g.x;it.y=g.y;it.w=g.w;it.h=g.h}}
  else{const g=mkPh(k,x,y,w,h,o);s.items.unshift(g);used.add(g.id)}});
 s.items=s.items.filter(i=>!i.ph||used.has(i.id)||(i.text&&String(i.text).trim())||i.html);
 s.items.sort((a,b)=>(b.ph?1:0)-(a.ph?1:0)||0);
 const order=['title','ctrTitle','pic','sub','body','hdr','body2','hdr2'];const ph=s.items.filter(i=>i.ph).sort((a,b)=>order.indexOf(a.ph)-order.indexOf(b.ph));s.items=ph.concat(s.items.filter(i=>!i.ph));}
function mkSlide(layout){const s={bg:themeCols().bg,items:[],notes:'',layout:'blank',comments:[]};if(layout)applyLayout(s,layout);return s}
/* ---------------- shape catalogue ---------------- */
const SH={
 rect:['Rectangle','M0 0H100V100H0Z','rect'],roundRect:['Rounded Rectangle',null,'roundRect'],ellipse:['Oval',null,'ellipse'],triangle:['Isosceles Triangle','M50 0L100 100L0 100Z','triangle'],rtTriangle:['Right Triangle','M0 0L100 100L0 100Z','rtTriangle'],
 diamond:['Diamond','M50 0L100 50L50 100L0 50Z','diamond'],parallelogram:['Parallelogram','M22 0H100L78 100H0Z','parallelogram'],trapezoid:['Trapezoid','M22 0H78L100 100H0Z','trapezoid'],pentagon:['Pentagon','M50 0L100 38L81 100L19 100L0 38Z','pentagon'],hexagon:['Hexagon','M25 0H75L100 50L75 100H25L0 50Z','hexagon'],octagon:['Octagon','M30 0H70L100 30V70L70 100H30L0 70V30Z','octagon'],
 star4:['4-Point Star','M50 0L62 38L100 50L62 62L50 100L38 62L0 50L38 38Z','star4'],star5:['5-Point Star','M50 0L61 35L98 35L68 57L79 91L50 70L21 91L32 57L2 35L39 35Z','star5'],star6:['6-Point Star','M50 0L63 25L93 25L77 50L93 75L63 75L50 100L37 75L7 75L23 50L7 25L37 25Z','star6'],
 rightArrow:['Right Arrow','M0 30H62V0L100 50L62 100V70H0Z','rightArrow'],leftArrow:['Left Arrow','M100 30H38V0L0 50L38 100V70H100Z','leftArrow'],upArrow:['Up Arrow','M30 100V38H0L50 0L100 38H70V100Z','upArrow'],downArrow:['Down Arrow','M30 0V62H0L50 100L100 62H70V0Z','downArrow'],leftRightArrow:['Left-Right Arrow','M0 50L28 10V32H72V10L100 50L72 90V68H28V90Z','leftRightArrow'],
 chevron:['Chevron','M0 0H65L100 50L65 100H0L35 50Z','chevron'],homePlate:['Pentagon Arrow','M0 0H72L100 50L72 100H0Z','homePlate'],plus:['Cross','M35 0H65V35H100V65H65V100H35V65H0V35H35Z','plus'],
 heart:['Heart','M50 92C10 60 0 40 0 25C0 8 14 0 28 0C38 0 46 6 50 14C54 6 62 0 72 0C86 0 100 8 100 25C100 40 90 60 50 92Z','heart'],cloud:['Cloud','M25 80C8 80 0 66 8 54C2 40 14 28 28 32C32 14 58 10 66 28C84 22 100 38 90 54C100 68 90 82 74 80Z','cloud'],lightning:['Lightning Bolt','M58 0L12 56H44L34 100L88 40H56Z','lightningBolt'],moon:['Moon','M62 0C30 6 10 28 10 50C10 74 32 96 62 100C38 90 28 70 28 50C28 30 40 10 62 0Z','moon'],
 callout:['Speech Callout','M0 0H100V72H58L26 100V72H0Z','wedgeRectCallout'],donut:['Donut','M50 0A50 50 0 1 0 50 100A50 50 0 1 0 50 0ZM50 24A26 26 0 1 1 50 76A26 26 0 1 1 50 24Z','donut'],
 line:['Line',null,'line'],arrow:['Arrow',null,'line']
};
const SHCAT=[['Lines',['line','arrow']],['Rectangles & Basic',['rect','roundRect','ellipse','triangle','rtTriangle','diamond','parallelogram','trapezoid','pentagon','hexagon','octagon','donut','plus','heart','cloud','moon','lightning']],['Block Arrows',['rightArrow','leftArrow','upArrow','downArrow','leftRightArrow','chevron','homePlate']],['Stars & Callouts',['star4','star5','star6','callout']]];
function shapeSVG(it,w,h){const sh=it.shape||'rect';const fill=it.fillN?'none':(it.grad?'url(#g'+it.id+')':col(it.fill||'acc1'));const op=it.op==null?1:it.op;const sw=it.sw==null?0:it.sw;const dash=it.dash?({dash:`${sw*3},${sw*2}`,dot:`${sw},${sw*1.5}`,dd:`${sw*4},${sw*2},${sw},${sw*2}`}[it.dash]||''):'';
 let g='';if(it.grad){const a=(it.grad.ang||0)*Math.PI/180;g=`<defs><linearGradient id="g${it.id}" x1="${50-50*Math.cos(a)}%" y1="${50-50*Math.sin(a)}%" x2="${50+50*Math.cos(a)}%" y2="${50+50*Math.sin(a)}%"><stop offset="0" stop-color="${col(it.grad.c1)}"/><stop offset="1" stop-color="${col(it.grad.c2)}"/></linearGradient></defs>`}
 const st=`fill="${fill}" fill-opacity="${op}" stroke="${sw?col(it.stroke||'#333333'):'none'}" stroke-width="${sw}" ${dash?`stroke-dasharray="${dash}"`:''} stroke-linejoin="round" vector-effect="non-scaling-stroke"`;
 if(sh==='line'||sh==='arrow'){const c=col(it.stroke||'#333333'),lw=it.sw||3;const ar=it.arrow||(sh==='arrow'?'end':'');const mk=`<defs><marker id="m${it.id}" markerWidth="6" markerHeight="6" refX="5" refY="3" orient="auto-start-reverse"><path d="M0 0L6 3L0 6Z" fill="${c}"/></marker></defs>`;
  return `<svg viewBox="0 0 ${w} ${Math.max(h,1)}" preserveAspectRatio="none">${mk}<line x1="0" y1="0" x2="${w}" y2="${h}" stroke="${c}" stroke-width="${lw}" ${it.dash?`stroke-dasharray="${({dash:lw*3+','+lw*2,dot:lw+','+lw*1.5,dd:lw*4+','+lw*2+','+lw+','+lw*2})[it.dash]||''}"`:''} ${ar==='end'||ar==='both'?`marker-end="url(#m${it.id})"`:''} ${ar==='start'||ar==='both'?`marker-start="url(#m${it.id})"`:''}/></svg>`}
 if(sh==='ellipse')return `<svg viewBox="0 0 100 100" preserveAspectRatio="none">${g}<ellipse cx="50" cy="50" rx="50" ry="50" ${st}/></svg>`;
 if(sh==='roundRect'){const r=it.r==null?16:it.r;return `<svg viewBox="0 0 ${w} ${h}" preserveAspectRatio="none">${g}<rect x="0" y="0" width="${w}" height="${h}" rx="${Math.min(w,h)*r/100}" ${st}/></svg>`}
 const d=(SH[sh]||SH.rect)[1];return `<svg viewBox="0 0 100 100" preserveAspectRatio="none">${g}<path d="${d}" fill-rule="evenodd" ${st}/></svg>`}
/* ---------------- text rendering ---------------- */
function ptx(it){return (it.size||(it.ph?theme()&&20:20))}
function isEmptyText(it){return !(it.html?it.html.replace(/<[^>]*>/g,'').trim():String(it.text||'').trim())}
function textHTML(it){if(it.html)return it.html;const L=String(it.text==null?'':it.text).split('\n');return L.map(l=>`<div>${l===''?'<br>':esc(l)}</div>`).join('')}
function baseTextStyle(it){const T=themeCols();const tit=/^(title|ctrTitle)$/.test(it.ph);const c=it.color?col(it.color):(it.ph?(tit?T.tc:(it.muted?mix(T.bc,T.bg,.4):T.bc)):T.bc);const f=it.font||(tit?T.major:T.minor);
 let s=`font-size:${(it.size||20)*(it.shr||1)/0.75}px;color:${c};font-family:'${f}',Arial,sans-serif;font-weight:${it.bold?700:400};font-style:${it.italic?'italic':'normal'};text-align:${it.align||'left'};line-height:${it.ls||1.1};letter-spacing:${(it.cs||0)}px;`;
 const td=[];if(it.underline)td.push('underline');if(it.strike)td.push('line-through');if(td.length)s+=`text-decoration:${td.join(' ')};`;if(it.shadowT)s+='text-shadow:2px 2px 3px rgba(0,0,0,.45);';if(it.caps)s+='text-transform:uppercase;';
 const va=it.valign||(it.type==='shape'?'m':'t');s+=`--spb:${it.spb||0}px;--spa:${it.spa||0}px;justify-content:${va==='m'?'center':va==='b'?'flex-end':'flex-start'};`;const m=it.mar==null?(it.type==='shape'?8:4):it.mar;s+=`padding:${it.pt==null?m:it.pt}px ${m}px;`;if(it.vert)s+='writing-mode:vertical-rl;';if(it.cols>1)s+=`column-count:${it.cols};`;return s}
function wrapText(it,mode){const cls=['tx'];if(it.bullet)cls.push({dot:'bl-dot',dash:'bl-dash',chk:'bl-chk',num:'bl-num'}[it.bullet]||'');
 const empty=isEmptyText(it);const prompt=it.ph&&empty&&mode==='edit';const body=prompt?`<div>${esc(PROMPT[it.ph]||'Click to add text')}</div>`:textHTML(it);
  return `<div class="${cls.join(' ')}" style="${baseTextStyle(it)}" data-tx="1" data-sp="${it.spb||0},${it.spa||0}">${body}</div>`}
/* ---------------- tables ---------------- */
const TSTYLES=[{n:'Medium Style 2 - Accent 1',h:'acc1',ht:'#ffffff',b1:'acc1:.78',b2:'acc1:.9',bd:'#ffffff'},{n:'Medium Style 2 - Accent 2',h:'acc2',ht:'#ffffff',b1:'acc2:.78',b2:'acc2:.9',bd:'#ffffff'},{n:'Medium Style 2 - Accent 3',h:'acc3',ht:'#ffffff',b1:'acc3:.78',b2:'acc3:.9',bd:'#ffffff'},{n:'Light Style 1',h:null,ht:'#000000',b1:'#e8e8e8',b2:'#ffffff',bd:'#999999'},{n:'Dark Style',h:'#333333',ht:'#ffffff',b1:'#555555',b2:'#444444',bd:'#222222',tt:'#ffffff'},{n:'Grid (no fill)',h:null,ht:'#000000',b1:'#ffffff',b2:'#ffffff',bd:'#333333'}];
function tint(spec){if(!spec)return null;const p=spec.split(':');return p[1]?mix(col(p[0]),'#ffffff',+p[1]):col(p[0])}
function tblStyle(it){return TSTYLES[it.ts||0]||TSTYLES[0]}
function cellFill(it,r,c,cell){if(cell&&cell.fill)return col(cell.fill);const s=tblStyle(it);if(it.hdr!==false&&r===0&&s.h)return col(s.h);if(it.hdr!==false&&r===0)return null;const k=it.hdr!==false?r-1:r;return it.band===false?tint(s.b1):tint(k%2===0?s.b1:s.b2)}
function cellTxt(it,r,c,cell){const s=tblStyle(it);if(cell&&cell.color)return col(cell.color);if(it.hdr!==false&&r===0)return s.ht;return s.tt||'#000000'}
function tableHTML(it){const R=it.cells.length,C=it.cells[0].length;const tw=it.cw||Array(C).fill(1/C);let h=`<table style="width:100%;height:100%"><colgroup>${tw.map(x=>`<col style="width:${x*100}%">`).join('')}</colgroup>`;
 it.cells.forEach((row,r)=>{h+=`<tr style="height:${(it.rh&&it.rh[r])||Math.max(30,it.h/R)}px">`;row.forEach((cell,c)=>{if(cell.hm)return;const fl=cellFill(it,r,c,cell);const s=tblStyle(it);const hd=it.hdr!==false&&r===0;
  h+=`<td data-r="${r}" data-c="${c}" ${cell.cs>1?`colspan="${cell.cs}"`:''} ${cell.rs>1?`rowspan="${cell.rs}"`:''} style="${fl?`background:${fl};`:''}color:${cellTxt(it,r,c,cell)};border-color:${s.bd};font-size:${(cell.size||it.size||18)/0.75}px;font-weight:${(hd||cell.bold)?700:400};font-style:${cell.italic?'italic':'normal'};text-align:${cell.al||(hd?'center':'left')};vertical-align:${cell.va||'middle'};font-family:'${it.font||themeCols().minor}',Arial,sans-serif">${cell.html||esc(cell.t||'').replace(/\n/g,'<br>')}</td>`});h+='</tr>'});return h+'</table>'}
function mkTable(R,C,w,h){const cells=[];for(let r=0;r<R;r++){const row=[];for(let c=0;c<C;c++)row.push({t:''});cells.push(row)}return{type:'table',x:Math.round((deck.w-w)/2),y:Math.round(deck.h*.3),w,h,cells,hdr:true,band:true,ts:0,size:18,cw:Array(C).fill(1/C)}}
/* ---------------- charts ---------------- */
const CHT=[['col','Clustered Column'],['colStacked','Stacked Column'],['bar','Clustered Bar'],['line','Line'],['area','Area'],['pie','Pie'],['doughnut','Doughnut'],['scatter','Scatter']];
function mkChart(type){return{type:'chart',x:Math.round(deck.w*.12),y:Math.round(deck.h*.22),w:Math.round(deck.w*.76),h:Math.round(deck.h*.68),ch:{t:type||'col',title:'Chart Title',cats:['Category 1','Category 2','Category 3','Category 4'],ser:[{n:'Series 1',v:[4.3,2.5,3.5,4.5]},{n:'Series 2',v:[2.4,4.4,1.8,2.8]},{n:'Series 3',v:[2,2,3,5]}],legend:'b',labels:false}}}
function chartSVG(ch,W,H){const cols=themeCols().acc;const T=themeCols();const tx=mix(T.bc,T.bg,.1);let o=`<svg viewBox="0 0 ${W} ${H}" width="100%" height="100%" font-family="Calibri,Arial" font-size="13" fill="${tx}">`;let top=10,bot=H-10,lft=10,rgt=W-10;
 if(ch.title){o+=`<text x="${W/2}" y="24" text-anchor="middle" font-size="18" font-weight="600">${esc(ch.title)}</text>`;top=40}
 if(ch.legend&&ch.legend!=='n'){const items=/pie|doughnut/.test(ch.t)?ch.cats:ch.ser.map(s=>s.n);if(ch.legend==='b'||ch.legend==='t'){let x=0;const ws=items.map(s=>String(s).length*7+26);const tot=ws.reduce((a,b)=>a+b,0);x=(W-tot)/2;const y=ch.legend==='b'?H-14:top+6;items.forEach((s,i)=>{o+=`<rect x="${x}" y="${y-9}" width="11" height="11" fill="${cols[i%6]}"/><text x="${x+15}" y="${y}">${esc(s)}</text>`;x+=ws[i]});if(ch.legend==='b')bot=H-30;else top+=22}else{let y=top+10;const x=W-110;items.forEach((s,i)=>{o+=`<rect x="${x}" y="${y-9}" width="11" height="11" fill="${cols[i%6]}"/><text x="${x+15}" y="${y}">${esc(s)}</text>`;y+=18});rgt=W-120}}
 const ns=ch.ser.length,nc=ch.cats.length;
 if(/pie|doughnut/.test(ch.t)){const v=(ch.ser[0]||{v:[]}).v.map(x=>Math.max(0,+x||0));const sum=v.reduce((a,b)=>a+b,0)||1;const cx=(lft+rgt)/2,cy=(top+bot)/2,r=Math.min(rgt-lft,bot-top)/2-6;let a=-Math.PI/2;v.forEach((x,i)=>{const b=a+x/sum*Math.PI*2;const big=b-a>Math.PI?1:0;const p=(rr,ang)=>[cx+rr*Math.cos(ang),cy+rr*Math.sin(ang)];const [x1,y1]=p(r,a),[x2,y2]=p(r,b);const ir=ch.t==='doughnut'?r*.55:0;const [x3,y3]=p(ir,b),[x4,y4]=p(ir,a);o+=`<path d="M${x1} ${y1}A${r} ${r} 0 ${big} 1 ${x2} ${y2}${ir?`L${x3} ${y3}A${ir} ${ir} 0 ${big} 0 ${x4} ${y4}Z`:`L${cx} ${cy}Z`}" fill="${cols[i%6]}" stroke="${T.bg}" stroke-width="2"/>`;if(ch.labels){const m=(a+b)/2;const [lx,ly]=p((r+ir)/2+(ir?0:r*.15),m);o+=`<text x="${lx}" y="${ly}" text-anchor="middle" fill="#fff" font-weight="600">${Math.round(x/sum*100)}%</text>`}a=b});return o+'</svg>'}
 let mx=0,mn=0;if(ch.t==='colStacked'){for(let i=0;i<nc;i++){let p=0;ch.ser.forEach(s=>p+=+s.v[i]||0);mx=Math.max(mx,p)}}else ch.ser.forEach(s=>s.v.forEach(x=>{mx=Math.max(mx,+x||0);mn=Math.min(mn,+x||0)}));if(mx===mn)mx=mn+1;const step=nice((mx-mn)/5);mx=Math.ceil(mx/step)*step;mn=Math.floor(mn/step)*step;
 const hor=ch.t==='bar';const L=lft+36,Rr=rgt-6,Tp=top+4,Bt=bot-22;const pw=Rr-L,ph=Bt-Tp;
 const yv=v=>hor?L+(v-mn)/(mx-mn)*pw:Bt-(v-mn)/(mx-mn)*ph;
 for(let g=mn;g<=mx+1e-9;g+=step){const p=yv(g);if(hor)o+=`<line x1="${p}" y1="${Tp}" x2="${p}" y2="${Bt}" stroke="#ddd"/><text x="${p}" y="${Bt+14}" text-anchor="middle">${nfmt(g)}</text>`;else o+=`<line x1="${L}" y1="${p}" x2="${Rr}" y2="${p}" stroke="#ddd"/><text x="${L-6}" y="${p+4}" text-anchor="end">${nfmt(g)}</text>`}
 if(ch.t==='scatter'){const xs=ch.cats.map(c=>+c||0);const xmx=Math.max(...xs,1),xmn=Math.min(...xs,0);ch.ser.forEach((s,i)=>s.v.forEach((y,j)=>{const x=L+((xs[j]||0)-xmn)/((xmx-xmn)||1)*pw;o+=`<circle cx="${x}" cy="${yv(+y||0)}" r="5" fill="${cols[i%6]}"/>`}));return o+'</svg>'}
 const band=(hor?ph:pw)/nc;ch.cats.forEach((c,i)=>{const p=(hor?Tp:L)+band*(i+.5);if(hor)o+=`<text x="${L-6}" y="${p+4}" text-anchor="end">${esc(c)}</text>`;else o+=`<text x="${p}" y="${Bt+16}" text-anchor="middle">${esc(c)}</text>`});
 if(ch.t==='line'||ch.t==='area'){ch.ser.forEach((s,i)=>{const pts=s.v.map((v,j)=>[L+band*(j+.5),yv(+v||0)]);if(ch.t==='area')o+=`<path d="M${pts[0][0]} ${Bt}L${pts.map(p=>p.join(' ')).join('L')}L${pts[pts.length-1][0]} ${Bt}Z" fill="${cols[i%6]}" fill-opacity=".55"/>`;o+=`<polyline points="${pts.map(p=>p.join(',')).join(' ')}" fill="none" stroke="${cols[i%6]}" stroke-width="3"/>`;if(ch.t==='line')pts.forEach(p=>o+=`<circle cx="${p[0]}" cy="${p[1]}" r="4" fill="${cols[i%6]}"/>`)});return o+'</svg>'}
 const gw=band*.7;ch.ser.forEach((s,i)=>s.v.forEach((v,j)=>{v=+v||0;const base=(hor?Tp:L)+band*j+(band-gw)/2;if(ch.t==='colStacked'){let p0=0;for(let k=0;k<i;k++)p0+=+ch.ser[k].v[j]||0;const y1=yv(p0),y2=yv(p0+v);o+=`<rect x="${base}" y="${Math.min(y1,y2)}" width="${gw}" height="${Math.abs(y1-y2)}" fill="${cols[i%6]}"/>`;return}const bw=gw/ns,bx=base+bw*i;if(hor){const x2=yv(v),x0=yv(0);o+=`<rect x="${Math.min(x0,x2)}" y="${bx}" width="${Math.abs(x2-x0)}" height="${bw}" fill="${cols[i%6]}"/>`;if(ch.labels)o+=`<text x="${x2+4}" y="${bx+bw/2+4}">${nfmt(v)}</text>`}else{const y2=yv(v),y0=yv(0);o+=`<rect x="${bx}" y="${Math.min(y0,y2)}" width="${bw}" height="${Math.abs(y2-y0)}" fill="${cols[i%6]}"/>`;if(ch.labels)o+=`<text x="${bx+bw/2}" y="${y2-4}" text-anchor="middle">${nfmt(v)}</text>`}}));return o+'</svg>'}
function nice(x){const e=Math.pow(10,Math.floor(Math.log10(x||1)));const f=x/e;return (f<1.5?1:f<3?2:f<7?5:10)*e}
function nfmt(v){return Math.abs(v)>=1000?Math.round(v).toLocaleString():(Math.round(v*100)/100)}
/* ---------------- slide decoration (theme) ---------------- */
function decoHTML(s){const T=themeCols();const W=deck.w,H=deck.h;if(s.nodeco||s.layout==='blank'&&false)return '';const a=T.acc;switch(T.deco){case 'leftbar':return `<div style="position:absolute;left:0;top:0;width:${W*.025}px;height:${H}px;background:${a[0]}"></div>`;case 'topbar':return `<div style="position:absolute;left:0;top:0;width:${W}px;height:${H*.02}px;background:${a[0]}"></div><div style="position:absolute;left:0;top:${H*.02}px;width:${W*.2}px;height:${H*.008}px;background:${a[1]}"></div>`;case 'band':return `<div style="position:absolute;left:0;bottom:0;width:${W}px;height:${H*.04}px;background:${a[0]}"></div>`;case 'corner':return `<svg style="position:absolute;right:0;bottom:0" width="${W*.22}" height="${H*.28}" viewBox="0 0 100 100" preserveAspectRatio="none"><path d="M100 0V100H0Z" fill="${a[0]}" fill-opacity=".85"/></svg>`}return ''}
/* ---------------- item rendering ---------------- */
function itemEl(it,mode){const d=document.createElement('div');d.className='it'+(it.ph?' ph':'')+(it.ph&&isEmptyText(it)?' empty':'');d.dataset.id=it.id;if(mode==='edit')d.id='x'+it.id;
 const isLine=it.type==='shape'&&(it.shape==='line'||it.shape==='arrow');let X=it.x,Y=it.y,Wd=it.w,Ht=it.h,lfx=false,lfy=false;if(isLine){X=Math.min(it.x,it.x+it.w);Y=Math.min(it.y,it.y+it.h);Wd=Math.abs(it.w);Ht=Math.abs(it.h);lfx=it.w<0;lfy=it.h<0}let st=`left:${X}px;top:${Y}px;width:${Wd}px;height:${Ht}px;`;if(it.rot)st+=`transform:rotate(${it.rot}deg);`;if(it.op!=null&&it.type!=='shape')st+=`opacity:${it.op};`;
 if(it.shadow)st+=`filter:drop-shadow(${it.shadow.x||3}px ${it.shadow.y||3}px ${it.shadow.b||4}px rgba(0,0,0,${it.shadow.o||.4}));`;
 if(it.hidden&&mode==='edit')st+='opacity:.35;';d.style.cssText=st+(isLine?'overflow:visible;':'')+(it.ph||it.type==='text'&&!it.fillN?'':'');
 if(isLine)d.classList.add('ln');
 if(it.type==='text'){if(it.fill&&!it.fillN)d.style.background=col(it.fill);if(it.sw){d.style.border=`${it.sw}px solid ${col(it.stroke||'#333')}`}d.innerHTML=wrapText(it,mode)}
 else if(it.type==='phpic'){if(mode==='edit')d.innerHTML=`<div style="width:100%;height:100%;border:2px dashed #aaa;display:flex;align-items:center;justify-content:center;color:#888;font-size:20px;background:rgba(0,0,0,.03)">&#128444; ${PROMPT.pic}</div>`}
 else if(it.type==='image'){const c=it.crop||{l:0,t:0,r:0,b:0};const cw=100-c.l-c.r,ch=100-c.t-c.b;let fl=[];if(it.bri!=null&&it.bri!==100)fl.push(`brightness(${it.bri}%)`);if(it.con!=null&&it.con!==100)fl.push(`contrast(${it.con}%)`);if(it.gray)fl.push('grayscale(1)');if(it.sepia)fl.push('sepia(1)');if(it.blur)fl.push(`blur(${it.blur}px)`);
  d.innerHTML=`<img src="${it.src}" style="position:absolute;left:${-c.l/cw*100}%;top:${-c.t/ch*100}%;width:${100/cw*100}%;height:${100/ch*100}%;${fl.length?`filter:${fl.join(' ')};`:''}${it.flipH||it.flipV?`transform:scale(${it.flipH?-1:1},${it.flipV?-1:1});`:''}">`;if(it.r)d.style.borderRadius=it.r+'px';if(it.sw)d.style.border=`${it.sw}px solid ${col(it.stroke||'#000')}`;d.style.overflow='hidden'}
 else if(it.type==='shape'){d.innerHTML=shapeSVG(it,Wd,Ht);const fx=(it.flipH?1:0)^(lfx?1:0),fy=(it.flipV?1:0)^(lfy?1:0);if(fx||fy)d.firstChild.style.transform=`scale(${fx?-1:1},${fy?-1:1})`;if(!isLine&&(!isEmptyText(it)||(mode==='edit'&&it.ph)))d.insertAdjacentHTML('beforeend',wrapText(it,mode));d.style.overflow='visible'}
 else if(it.type==='table'){d.innerHTML=tableHTML(it);d.style.overflow='visible'}
 else if(it.type==='chart'){d.innerHTML=chartSVG(it.ch,it.w,it.h);d.style.background=it.fill?col(it.fill):''}
 if(it.link&&mode==='show')d.style.cursor='pointer';return d}
/* ---------------- whole slide rendering ---------------- */
function paintSlide(el,s,mode,anim){el.style.width=deck.w+'px';el.style.height=deck.h+'px';el.style.background=s.bgImg?`url(${s.bgImg}) center/cover`:s.bg;if(s.grad)el.style.background=`linear-gradient(${s.grad.ang||90}deg,${col(s.grad.c1)},${col(s.grad.c2)})`;el.innerHTML=decoHTML(s);s.items.forEach(it=>{if(it.hidden&&mode!=='edit')return;el.appendChild(itemEl(it,mode))});
 const f=deck.hf;if(f&&mode!=='edit-skip'){const T=themeCols();const mk=(t,x,w,al)=>`<div style="position:absolute;left:${x}px;bottom:${deck.h*.02}px;width:${w}px;text-align:${al};font-size:${Math.round(deck.h/27)}px;color:${mix(T.bc,T.bg,.45)};font-family:'${T.minor}',Arial">${esc(t)}</div>`;const idx=deck.slides.indexOf(s)+1;if(!(f.notTitle&&s.layout==='title')){if(f.date&&f.dateText)el.insertAdjacentHTML('beforeend',mk(f.dateText,deck.w*.04,deck.w*.25,'left'));if(f.footer&&f.footerText)el.insertAdjacentHTML('beforeend',mk(f.footerText,deck.w*.34,deck.w*.32,'center'));if(f.num)el.insertAdjacentHTML('beforeend',mk(String(idx),deck.w*.71,deck.w*.25,'right'))}}}
function render(keepFpane){const st=$('stage');paintSlide(st,S(),'edit');const ov=document.createElement('div');ov.id='ov';st.appendChild(ov);if(PV.grid){const g=document.createElement('div');g.className='grid';st.insertBefore(g,ov)}
 fitStage();drawSel();renderThumbs();$('stL').textContent=`Slide ${cur+1} of ${deck.slides.length}`;$('stR').textContent=S().hidden?'(hidden)':'';if($('notesTa')&&document.activeElement!==$('notesTa'))$('notesTa').value=S().notes||'';if(window.syncRibbon)syncRibbon();if(PV.fpane&&window.fpaneDraw&&!keepFpane)fpaneDraw();if(view==='sorter')renderSorter();}
function setZoom(z){uz=Math.max(.3,Math.min(4,Math.round(z*100)/100));$('zs').value=uz*100;$('zp').textContent=Math.round(uz*100)+'%';fitStage();drawSel()}function zoomBy(d){setZoom(uz+d)}
function fitStage(){const w=$('stageWrap');const aw=w.clientWidth-20,ah=w.clientHeight-20;if(aw<=0)return;scale=Math.min(aw/deck.w,ah/deck.h)*uz;$('stage').style.transform=`scale(${scale})`;const b=$('stageBox');b.style.width=deck.w*scale+'px';b.style.height=deck.h*scale+'px'}
window.addEventListener('resize',()=>{fitStage();drawSel();renderThumbs()});
let thumbCache=null;
function renderThumbs(){const T=$('thumbs');T.innerHTML='';const land=matchMedia('(orientation:landscape)').matches;const tw=land?118:Math.round(80*deck.w/deck.h);const k=tw/deck.w;let lastSec=null;
 deck.slides.forEach((s,i)=>{if(s.section!=null&&s.section!==lastSec||(i===0&&s.section)){const sd=document.createElement('div');sd.className='sec';sd.textContent=s.section||'Section';sd.onclick=()=>sectionMenu(i);T.appendChild(sd);lastSec=s.section}
  const d=document.createElement('div');d.className='th'+(i===cur?' on':'')+(s.hidden?' hid':'');d.dataset.i=i;d.style.width=tw+'px';d.style.height=Math.round(deck.h*k)+'px';d.innerHTML=`<span class="n">${i+1}${s.tr&&s.tr.t&&s.tr.t!=='none'?' &#9733;':''}</span>`;const in_=document.createElement('div');in_.className='in';in_.style.cssText=`transform:scale(${k})`;paintSlide(in_,s,'thumb');d.appendChild(in_);T.appendChild(d);bindThumb(d,i)})}
