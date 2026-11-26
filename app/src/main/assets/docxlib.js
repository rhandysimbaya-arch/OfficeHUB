// Minimal .docx writer (needs JSZip). Blocks: {type:'p',runs:[{t,b,i,u,sz,color,field,br}],align,heading,bullet} | {type:'table',rows:[[ [blocks] | 'text' ]]} | {type:'pagebreak'}
window.Docx = (function(){
  const esc = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  function run(r){
    if (r.br) return '<w:r><w:br/></w:r>';
    let pr = '';
    if (r.b) pr += '<w:b/>'; if (r.i) pr += '<w:i/>'; if (r.u) pr += '<w:u w:val="single"/>';
    if (r.color) pr += '<w:color w:val="'+esc(r.color.replace('#',''))+'"/>';
    if (r.font) pr += '<w:rFonts w:ascii="'+esc(r.font)+'" w:hAnsi="'+esc(r.font)+'" w:cs="'+esc(r.font)+'"/>';
    if (r.sz) pr += '<w:sz w:val="'+Math.round(r.sz*2)+'"/><w:szCs w:val="'+Math.round(r.sz*2)+'"/>';
    pr = pr ? '<w:rPr>'+pr+'</w:rPr>' : '';
    if (r.field) return '<w:r>'+pr+'<w:fldChar w:fldCharType="begin"/></w:r><w:r>'+pr+'<w:instrText xml:space="preserve"> '+esc(r.field)+' </w:instrText></w:r><w:r>'+pr+'<w:fldChar w:fldCharType="separate"/></w:r><w:r>'+pr+'<w:t>1</w:t></w:r><w:r>'+pr+'<w:fldChar w:fldCharType="end"/></w:r>';
    return '<w:r>'+pr+'<w:t xml:space="preserve">'+esc(r.t==null?'':r.t)+'</w:t></w:r>';
  }
  function para(p){
    let pp = '';
    if (p.heading) pp += '<w:pStyle w:val="Heading'+p.heading+'"/>';
    if (p.bullet) pp += '<w:pStyle w:val="ListBullet"/>';
    if (p.align) pp += '<w:jc w:val="'+({left:'left',center:'center',right:'right',justify:'both',both:'both'}[p.align]||'left')+'"/>';
    return '<w:p>'+(pp?'<w:pPr>'+pp+'</w:pPr>':'')+(p.runs||[]).map(run).join('')+'</w:p>';
  }
  function block(b){
    if (b.type === 'pagebreak') return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    if (b.type === 'table') {
      const cols = Math.max.apply(null, b.rows.map(r=>r.length).concat([1]));
      let x = '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>'+['top','left','bottom','right','insideH','insideV'].map(s=>'<w:'+s+' w:val="single" w:sz="4" w:space="0" w:color="999999"/>').join('')+'</w:tblBorders></w:tblPr><w:tblGrid>'+'<w:gridCol w:w="'+Math.floor(9000/cols)+'"/>'.repeat(cols)+'</w:tblGrid>';
      b.rows.forEach(r=>{ x += '<w:tr>'; for (let c=0;c<cols;c++){ const cell = r[c]; let inner;
        if (cell == null) inner = '<w:p/>'; else if (typeof cell === 'string') inner = para({runs:[{t:cell}]}); else inner = cell.map(block).join('') || '<w:p/>';
        x += '<w:tc><w:tcPr><w:tcW w:w="'+Math.floor(9000/cols)+'" w:type="dxa"/></w:tcPr>'+inner+'</w:tc>'; } x += '</w:tr>'; });
      return x + '</w:tbl><w:p/>';
    }
    return para(b);
  }
  async function build(m){
    const zip = new JSZip(); const put=(n,c)=>zip.file(n,c,{createFolders:false}); const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const hasH = m.header && m.header.length, hasF = m.footer && m.footer.length;
    put('[Content_Types].xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'+(hasH?'<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>':'')+(hasF?'<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>':'')+'</Types>');
    put('_rels/.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    put('word/_rels/document.xml.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'+(hasH?'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>':'')+(hasF?'<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>':'')+'</Relationships>');
    put('word/styles.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles '+NS+'><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'+[1,2,3].map((n,i)=>'<w:style w:type="paragraph" w:styleId="Heading'+n+'"><w:name w:val="heading '+n+'"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/></w:pPr><w:rPr><w:b/><w:color w:val="2F5496"/><w:sz w:val="'+[32,26,24][i]+'"/></w:rPr></w:style>').join('')+'<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style></w:styles>');
    put('word/numbering.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering '+NS+'><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>');
    if (hasH) put('word/header1.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr '+NS+'>'+m.header.map(block).join('')+'</w:hdr>');
    if (hasF) put('word/footer1.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr '+NS+'>'+m.footer.map(block).join('')+'</w:ftr>');
    const pg = m.page || {w:12240,h:15840};
    const sect = '<w:sectPr>'+(hasH?'<w:headerReference w:type="default" r:id="rId3"/>':'')+(hasF?'<w:footerReference w:type="default" r:id="rId4"/>':'')+'<w:pgSz w:w="'+pg.w+'" w:h="'+pg.h+'"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
    put('word/document.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document '+NS+'><w:body>'+(m.blocks||[]).map(block).join('')+sect+'</w:body></w:document>');
    return zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  }
  return {build};
})();
