import {createDocxArchive} from './diagnostic-export.js';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R='http://schemas.openxmlformats.org/package/2006/relationships';
const CT='http://schemas.openxmlformats.org/package/2006/content-types';
const OFFICE='http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const utf8=new TextDecoder();
const children=(node,name)=>[...node.children].filter(n=>n.namespaceURI===W&&n.localName===name);
const text=node=>node.namespaceURI===W&&node.localName==='t'?node.textContent:node.namespaceURI===W&&node.localName==='tab'?' ':[...node.childNodes].map(n=>n.nodeType===1?text(n):'').join('');
function parse(bytes){const doc=new DOMParser().parseFromString(typeof bytes==='string'?bytes:utf8.decode(bytes),'application/xml');if(doc.getElementsByTagName('parsererror').length)throw new Error('原文档 XML 无法读取');return doc;}
const serialize=doc=>'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+new XMLSerializer().serializeToString(doc).replace(/^<\?xml[^?]*\?>\s*/,'');
function element(doc,name,attributes={}){const e=doc.createElementNS(W,'w:'+name);for(const [k,v] of Object.entries(attributes))e.setAttributeNS(W,'w:'+k,String(v));return e;}
function commentParagraph(doc,value,bold=false){const p=element(doc,'p'),r=element(doc,'r');if(bold){const pr=element(doc,'rPr');pr.append(element(doc,'b'));r.append(pr);}const t=element(doc,'t');t.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve');t.textContent=String(value??'').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,'');r.append(t);p.append(r);return p;}

async function unzip(buffer){
 const bytes=new Uint8Array(buffer),view=new DataView(buffer),files=Object.create(null);let end=-1;
 for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--)if(view.getUint32(i,true)===0x06054B50&&i+22+view.getUint16(i+20,true)===bytes.length){end=i;break;}
 if(end<0)throw new Error('原文件不是有效的 DOCX 压缩包');
 const count=view.getUint16(end+10,true);let cursor=view.getUint32(end+16,true),total=0;
 if(view.getUint16(end+4,true)||view.getUint16(end+6,true)||count===65535)throw new Error('暂不支持分卷或 ZIP64 文档');
 for(let i=0;i<count;i++){
  if(cursor+46>bytes.length||view.getUint32(cursor,true)!==0x02014B50)throw new Error('DOCX 目录已损坏');
  const flags=view.getUint16(cursor+8,true),method=view.getUint16(cursor+10,true),compressed=view.getUint32(cursor+20,true),size=view.getUint32(cursor+24,true),nameLength=view.getUint16(cursor+28,true),extraLength=view.getUint16(cursor+30,true),commentLength=view.getUint16(cursor+32,true),offset=view.getUint32(cursor+42,true);
  const name=utf8.decode(bytes.slice(cursor+46,cursor+46+nameLength));cursor+=46+nameLength+extraLength+commentLength;
  if(name.endsWith('/'))continue;
  if(flags&1)throw new Error('请先解除 DOCX 加密后再导出批注版');
  total+=size;if(size>64*1024*1024||total>128*1024*1024)throw new Error('原文档解压后过大，暂不支持批注导出');
  if(offset+30>bytes.length||view.getUint32(offset,true)!==0x04034B50)throw new Error('DOCX 文件内容已损坏');
  const start=offset+30+view.getUint16(offset+26,true)+view.getUint16(offset+28,true);
  if(start+compressed>bytes.length)throw new Error('DOCX 文件数据不完整');
  const source=bytes.slice(start,start+compressed);
  if(method===0)files[name]=source;
  else if(method===8){
   let stream;try{stream=new DecompressionStream('deflate-raw');}catch{throw new Error('当前浏览器不支持 DOCX 解压，请使用新版 Edge 或 Chrome');}
   const reader=new Blob([source]).stream().pipeThrough(stream).getReader(),chunks=[];let length=0;
   try{while(true){const {value,done}=await reader.read();if(done)break;length+=value.length;if(length>size||length>64*1024*1024)throw new Error('DOCX 解压数据超出声明大小');chunks.push(value);}}catch(error){await reader.cancel().catch(()=>{});throw error;}
   const output=new Uint8Array(length);let position=0;for(const chunk of chunks){output.set(chunk,position);position+=chunk.length;}files[name]=output;
  }else throw new Error('暂不支持该 DOCX 的压缩方式');
  if(files[name].length!==size)throw new Error('DOCX 解压后的文件长度不符');
 }
 return files;
}
function anchors(doc){
 const body=doc.getElementsByTagNameNS(W,'body')[0];if(!body)throw new Error('原文档缺少正文');
 const out=[];
 for(const node of body.children){
  if(node.namespaceURI!==W)continue;
  if(node.localName==='p'){const value=text(node).trim();if(value)out.push({id:'p'+(out.length+1),node,text:value});}
  if(node.localName==='tbl')for(const tr of children(node,'tr')){
   const cells=children(tr,'tc'),paragraphs=[...tr.getElementsByTagNameNS(W,'p')];
   out.push({id:'p'+(out.length+1),node:paragraphs.find(p=>text(p).trim())||paragraphs[0],text:cells.map(c=>text(c).trim()).join(' | ')});
  }
 }
 return out;
}
function ensureRelation(doc,type,target){const root=doc.documentElement;const existing=[...root.children].find(r=>r.getAttribute('Type')===OFFICE+type);if(existing)return existing;const relation=doc.createElementNS(R,'Relationship'),ids=new Set([...root.children].map(r=>r.getAttribute('Id')));let n=1;while(ids.has('rIdAnnotation'+n))n++;relation.setAttribute('Id','rIdAnnotation'+n);relation.setAttribute('Type',OFFICE+type);relation.setAttribute('Target',target);root.append(relation);return relation;}
function ensureType(doc,part,type){if([...doc.documentElement.children].some(e=>e.getAttribute('PartName')===part))return;const item=doc.createElementNS(CT,'Override');item.setAttribute('PartName',part);item.setAttribute('ContentType',type);doc.documentElement.append(item);}
function commentPath(relation){return new URL(relation.getAttribute('Target'),'https://docx.invalid/word/document.xml').pathname.slice(1);}

export async function downloadAnnotatedDocument(task,user){
 if(!task?.results?.length)throw new Error('请先完成检测并打开报告');
 const target=task.selection?.selected?.find(d=>d.id===task.targetId);
 if(!/\.docx$/i.test(target?.name||task.targetName||''))throw new Error('导出批注版仅支持 DOCX 原文档');
 const response=await fetch('/api/documents/'+encodeURIComponent(task.targetId)+'/download');
 if(!response.ok){const error=await response.json().catch(()=>null);throw new Error(error?.error||'无法取得原 DOCX 文档，请重新登录后再试');}
 const buffer=await response.arrayBuffer();
 if(target?.hash&&globalThis.crypto?.subtle){const digest=await crypto.subtle.digest('SHA-256',buffer),hash=[...new Uint8Array(digest)].map(b=>b.toString(16).padStart(2,'0')).join('');if(hash!==target.hash)throw new Error('原文档与本次检测快照不一致，请重新检测后导出');}
 const files=await unzip(buffer);if(!files['word/document.xml']||!files['[Content_Types].xml'])throw new Error('原文件缺少 DOCX 必要内容');
 const doc=parse(files['word/document.xml']),locations=anchors(doc),first=locations.find(a=>a.node)?.node;if(!first)throw new Error('原文档没有可添加批注的正文段落');
 const relations=parse(files['word/_rels/document.xml.rels']||`<Relationships xmlns="${R}"/>`),types=parse(files['[Content_Types].xml']);
 const relation=ensureRelation(relations,'comments','comments.xml'),path=commentPath(relation);
 const comments=parse(files[path]||`<w:comments xmlns:w="${W}"/>`),existing=[...comments.getElementsByTagNameNS(W,'comment'),...doc.getElementsByTagNameNS(W,'commentRangeStart'),...doc.getElementsByTagNameNS(W,'commentReference')];
 let nextId=Math.max(-1,...existing.map(e=>Number(e.getAttributeNS(W,'id'))).filter(Number.isFinite))+1;
 const groups=new Map();
 const add=(paragraph,rule,missing=false)=>{if(!groups.has(paragraph))groups.set(paragraph,new Map());groups.get(paragraph).set(rule.id,{rule,missing});};
 for(const rule of task.results){
  const evidence=(rule.evidence||[]).filter(e=>e.documentId===task.targetId);let located=false;
  for(const e of evidence){const a=locations.find(a=>a.id===e.fragmentId&&a.text===e.text)||locations.find(a=>a.text===e.text)||locations.find(a=>e.text&&a.text.includes(e.text));if(a?.node){add(a.node,rule);located=true;}}
  if(!located)add(first,rule,true);
 }
 const created=new Date().toISOString();
 for(const [paragraph,entries] of groups){
  const id=nextId++,comment=element(comments,'comment',{id,author:`因构智采 · ${user?.name||'审核人员'}`,initials:'AI',date:created});
  for(const {rule,missing} of entries.values()){
   comment.append(commentParagraph(comments,`[${rule.id}] ${rule.status} · ${rule.severity||'一般'}级`,true),commentParagraph(comments,rule.condition));
   if(missing)comment.append(commentParagraph(comments,'定位说明：未取得精确的目标原文位置，此处为待核验提醒。'));
   comment.append(commentParagraph(comments,'诊断：'+(rule.reason||'—')),commentParagraph(comments,'建议：'+(rule.suggestion||'—')));
   for(const e of rule.evidence||[])if(e.documentId!==task.targetId)comment.append(commentParagraph(comments,`基准：${e.name} · ${e.location||'正文'}`),commentParagraph(comments,e.text||''));
  }
  comment.append(commentParagraph(comments,`检测任务：${task.id}；规则版本：${task.rules?.version||'—'}`));comments.documentElement.append(comment);
  const start=element(doc,'commentRangeStart',{id}),end=element(doc,'commentRangeEnd',{id}),reference=element(doc,'r'),properties=element(doc,'rPr');properties.append(element(doc,'rStyle',{val:'CommentReference'}));reference.append(properties,element(doc,'commentReference',{id}));
  const paragraphProperties=children(paragraph,'pPr')[0];paragraph.insertBefore(start,paragraphProperties?paragraphProperties.nextSibling:paragraph.firstChild);paragraph.append(end,reference);
 }
 ensureType(types,'/'+path,'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');
 const settingsRelation=ensureRelation(relations,'settings','settings.xml'),settingsPath=commentPath(settingsRelation),settings=parse(files[settingsPath]||`<w:settings xmlns:w="${W}"/>`);
 let revision=settings.getElementsByTagNameNS(W,'revisionView')[0];if(!revision){revision=element(settings,'revisionView');settings.documentElement.append(revision);}revision.setAttributeNS(W,'w:comments','true');
 ensureType(types,'/'+settingsPath,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
 files['word/document.xml']=serialize(doc);files[path]=serialize(comments);files['word/_rels/document.xml.rels']=serialize(relations);files['[Content_Types].xml']=serialize(types);files[settingsPath]=serialize(settings);
 const blob=createDocxArchive(files),url=URL.createObjectURL(blob),link=document.createElement('a');
 link.href=url;link.download=String(target?.name||task.targetName).replace(/\.docx$/i,'')+'_批注版.docx';document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);
 return groups.size;
}
