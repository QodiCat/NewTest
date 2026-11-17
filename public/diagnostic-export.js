// Generate a real OOXML/DOCX archive in the browser from the visible report.
// No extra server route or third-party service is needed for this download.
const encode=new TextEncoder();
const xml=value=>String(value??'').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
const run=(text,{bold=false,color,size}={})=>`<w:r><w:rPr>${bold?'<w:b/>':''}${color?`<w:color w:val="${color}"/>`:''}${size?`<w:sz w:val="${size}"/>`:''}</w:rPr>${String(text??'').split(/\r?\n/).map((line,i)=>`${i?'<w:br/>':''}<w:t xml:space="preserve">${xml(line)}</w:t>`).join('')}</w:r>`;
const p=(text,style='Normal',options={})=>`<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>${run(text,options)}</w:p>`;
const color=status=>status==='通过'?'26745C':status==='不通过'?'B5473B':'9A6C20';
const row=(cells,header=false)=>`<w:tr>${header?'<w:trPr><w:tblHeader/></w:trPr>':''}${cells.map(cell=>`<w:tc><w:tcPr>${header?'<w:shd w:fill="EAF0F8"/>':''}<w:vAlign w:val="center"/></w:tcPr>${p(cell,'TableText',{bold:header})}</w:tc>`).join('')}</w:tr>`;
const table=(heads,rows)=>`<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>${['top','left','bottom','right','insideH','insideV'].map(edge=>`<w:${edge} w:val="single" w:sz="4" w:color="CBD5E1"/>`).join('')}</w:tblBorders><w:tblCellMar><w:top w:w="90" w:type="dxa"/><w:left w:w="110" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/><w:right w:w="110" w:type="dxa"/></w:tblCellMar></w:tblPr>${row(heads,true)}${rows.map(c=>row(c)).join('')}</w:tbl>${p('')}`;
const timestamp=value=>value?new Date(value).toLocaleString('zh-CN',{timeZone:'Asia/Shanghai',hour12:false}):'—';
function documentXml(task,project,user){
 const results=task.results,summary={total:results.length,pass:results.filter(r=>r.status==='通过').length,fail:results.filter(r=>r.status==='不通过').length,unknown:results.filter(r=>r.status==='无法核验').length};
 const title=task.stage==='evaluation'?'评标报告诊断报告':'招标文件诊断报告';
 const body=[p(title,'Title'),p(project?.name||'采购项目','Subtitle'),p('一、任务与版本信息','Heading1'),table(['项目','内容'],[
  ['项目编号',project?.code||task.projectId],['检测任务',task.id],['目标文档',task.targetName],['任务状态',task.status],['执行时间',timestamp(task.created)],['完成时间',timestamp(task.completed)],['规则版本',task.rules?.version||'—'],['解析版本',task.parserVersion||'—'],['关联初检',task.parentTaskId||'无'],['导出人员',user?.name||'—'],['导出时间',timestamp(new Date().toISOString())]
 ]),p('二、诊断结果汇总','Heading1'),table(['检测项','通过','不通过','无法核验'],[[summary.total,summary.pass,summary.fail,summary.unknown]]),p(summary.fail?`发现 ${summary.fail} 项不通过，建议按逐项诊断意见整改后复检。`:summary.unknown?'存在无法核验项，请补齐依据后复核。':'本次纳入检测的全部检查项均通过。'),table(['编号','检测条件','结论／级别'],results.map(r=>[r.id,r.condition,`${r.status}／${r.severity||'—'}`])),p('三、逐项诊断与原文证据','Heading1')];
 for(const r of results){
  body.push(p(`${r.id}  ${r.condition}`,'Heading2'),p(`诊断结论：${r.status}`,'Normal',{bold:true,color:color(r.status)}),p(`诊断说明：${r.reason||'—'}`),p(`检查依据：${r.source||'见任务规则快照'}`),p(`处理建议：${r.suggestion||'—'}`));
  if(!r.evidence?.length)body.push(p('原文证据：当前检测项未取得可定位证据。'));
  else for(const [i,e] of r.evidence.entries()){
   const kind=e.documentId===task.targetId?'目标文档原文':'对照基准原文';
   body.push(p(`${i+1}. ${kind}`,'EvidenceHeading',{bold:true}),p(`文件：${e.name||e.documentId}  ${e.version?`· 版本 ${e.version}`:''}`,'EvidenceMeta'),p(`位置：${e.page?`第 ${e.page} 页 · `:''}${e.location||'未记录位置'}`,'EvidenceMeta'),p(e.text||'—','EvidenceText'));
  }
 }
 body.push(p('四、本次引用的材料版本','Heading1'),table(['文件名称','版本','内容摘要 SHA-256'],(task.selection?.selected||[]).map(d=>[d.name,d.version||'—',d.hash||'—'])),p('本报告使用检测任务中固化的目标文档、基准材料和规则结果生成。文档未建立固定页码时，使用章节、段落或表格行定位。','Note'));
 return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body.join('')}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr></w:body></w:document>`;
}
const styles=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/><w:sz w:val="21"/><w:color w:val="24364B"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="140" w:line="320" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${[
 ['Title','Title',38,'24446B',true],['Subtitle','Subtitle',24,'64748B',false],['Heading1','heading 1',28,'24446B',true],['Heading2','heading 2',24,'385984',true],['EvidenceHeading','Evidence Heading',21,'385984',true],['EvidenceMeta','Evidence Metadata',18,'64748B',false],['EvidenceText','Evidence Text',20,'334155',false],['TableText','Table Text',19,'334155',false],['Note','Note',18,'64748B',false]
].map(([id,name,size,c,b])=>`<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:pPr>${b?'<w:keepNext/>':''}<w:spacing w:before="${id.startsWith('Heading')?240:60}" w:after="120"/></w:pPr><w:rPr>${b?'<w:b/>':''}<w:sz w:val="${size}"/><w:color w:val="${c}"/></w:rPr></w:style>`).join('')}</w:styles>`;
const crcTable=Uint32Array.from({length:256},(_,n)=>{let c=n;for(let i=0;i<8;i++)c=c&1?0xEDB88320^(c>>>1):c>>>1;return c>>>0;});
function crc32(bytes){let crc=0xFFFFFFFF;for(const b of bytes)crc=crcTable[(crc^b)&255]^(crc>>>8);return (crc^0xFFFFFFFF)>>>0;}
function header(length){const bytes=new Uint8Array(length),view=new DataView(bytes.buffer);return {bytes,u16:(offset,n)=>view.setUint16(offset,n,true),u32:(offset,n)=>view.setUint32(offset,n,true)};}
export function createDocxArchive(files){
 const local=[],central=[];let offset=0,centralSize=0;
 const now=new Date(),dosTime=(now.getHours()<<11)|(now.getMinutes()<<5)|(now.getSeconds()>>1),dosDate=((Math.max(1980,now.getFullYear())-1980)<<9)|((now.getMonth()+1)<<5)|now.getDate();
 for(const [name,source] of Object.entries(files)){
  const filename=encode.encode(name),data=typeof source==='string'?encode.encode(source):source,crc=crc32(data),h=header(30);
  h.u32(0,0x04034B50);h.u16(4,20);h.u16(6,0x0800);h.u16(10,dosTime);h.u16(12,dosDate);h.u32(14,crc);h.u32(18,data.length);h.u32(22,data.length);h.u16(26,filename.length);local.push(h.bytes,filename,data);
  const c=header(46);c.u32(0,0x02014B50);c.u16(4,20);c.u16(6,20);c.u16(8,0x0800);c.u16(12,dosTime);c.u16(14,dosDate);c.u32(16,crc);c.u32(20,data.length);c.u32(24,data.length);c.u16(28,filename.length);c.u32(42,offset);central.push(c.bytes,filename);centralSize+=46+filename.length;offset+=30+filename.length+data.length;
 }
 const end=header(22),count=Object.keys(files).length;end.u32(0,0x06054B50);end.u16(8,count);end.u16(10,count);end.u32(12,centralSize);end.u32(16,offset);
 return new Blob([...local,...central,end.bytes],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
}
export function downloadDiagnosticReport(task,project,user){
 if(!task?.results?.length)throw new Error('请先完成检测并打开报告，再导出诊断文档');
 const blob=createDocxArchive({
  '[Content_Types].xml':'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
  '_rels/.rels':'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  'word/document.xml':documentXml(task,project,user),
  'word/styles.xml':styles,
  'word/_rels/document.xml.rels':'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
 });
 const title=task.stage==='evaluation'?'评标报告诊断报告':'招标文件诊断报告',base=String(task.targetName||project?.name||'采购项目').replace(/\.docx$/i,'').replace(/[<>:"/\\|?*\u0000-\u001f]/g,'_').slice(0,100);
 const url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download=`${base}_${title}_${task.id.slice(0,8)}.docx`;document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);
}
