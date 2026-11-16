import AdmZip from 'adm-zip';
import { XMLParser } from 'fast-xml-parser';
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const parser = new XMLParser({preserveOrder:true,ignoreAttributes:false,trimValues:false});
const nodes = (a,key) => (a||[]).flatMap(n=>n[key] ? [n[key]] : Object.values(n).filter(Array.isArray).flatMap(v=>nodes(v,key)));
const text = a => (a||[]).map(n=> n['#text']!==undefined ? String(n['#text']) : n['w:tab'] ? ' ' : Object.values(n).filter(Array.isArray).map(text).join('')).join('');
export async function parseDocument(path, name) {
  const buffer=await readFile(path), hash=createHash('sha256').update(buffer).digest('hex');
  const fragments=[]; let section='首页';
  if (/\.docx$/i.test(name)) {
    const zip=new AdmZip(buffer), entry=zip.getEntry('word/document.xml');
    if(!entry || entry.header.size>30e6) throw new Error('DOCX 损坏或解压后正文过大');
    const body=nodes(parser.parse(zip.readAsText(entry)),'w:body')[0]||[];
    let table=0;
    for(const node of body) {
      if(node['w:p']) {
        const value=text(node['w:p']).trim(); if(!value)continue;
        if(/^(\d+\.|[A-D]\.\d|附件\s*[A-Z\d]|第\s*\d+\s*条)/.test(value))section=value;
        fragments.push({id:`p${fragments.length+1}`,text:value,section,location:`${section} · 段落 ${fragments.length+1}`,page:null});
      } else if(node['w:tbl']) {
        table++;let row=0;
        for(const tr of node['w:tbl'].filter(n=>n['w:tr'])) {
          row++; const cells=tr['w:tr'].filter(n=>n['w:tc']).map(n=>text(n['w:tc']).trim());
          fragments.push({id:`p${fragments.length+1}`,text:cells.join(' | '),cells,section,location:`${section} · 表 ${table} 行 ${row}`,page:null});
        }
      }
    }
  } else if (/\.pdf$/i.test(name)) {
    const {getDocument}=await import('pdfjs-dist/legacy/build/pdf.mjs');
    const pdf=await getDocument({data:new Uint8Array(buffer),useSystemFonts:true,isEvalSupported:false}).promise;
    for(let page=1;page<=pdf.numPages;page++) {
      const p=await pdf.getPage(page), content=await p.getTextContent();
      const value=content.items.map(i=>i.str||'').join(' ');
      fragments.push({id:`p${page}`,page,section:`第 ${page} 页`,location:`第 ${page} 页`,text:value,coordinates:content.items.filter(i=>i.str).map(i=>({text:i.str,x:i.transform[4],y:i.transform[5],width:i.width,height:i.height}))});
    }
    await pdf.destroy();
    if(fragments.some(f=>f.text.trim().length<20) && process.env.OCR_COMMAND) {
      const {stdout}=await promisify(execFile)(process.env.OCR_COMMAND,[path],{timeout:180000,maxBuffer:20e6});
      const pages=JSON.parse(stdout).pages;
      for(const p of pages||[]) {const f=fragments.find(f=>f.page===p.page);if(f && f.text.trim().length<20){f.text=p.text;f.ocr=true;}}
    }
  } else if (/\.(txt|md)$/i.test(name)) {
    buffer.toString('utf8').split(/\r?\n/).filter(Boolean).forEach((t,i)=>fragments.push({id:`p${i+1}`,text:t,section:'正文',location:`行 ${i+1}`,page:null}));
  } else throw new Error('仅支持 DOCX、PDF、TXT、MD');
  if(!fragments.length)throw new Error('未提取到正文，请补充可读版本');
  const low=fragments.filter(f=>f.text.trim().length<20 && f.page).length;
  return {hash,fragments,text:fragments.map(f=>f.text).join('\n'),quality:low?'需补录或配置 OCR':'可读取',parserVersion:'local-parser-1.0',pageNote:/\.docx$/i.test(name)?'DOCX 使用章节、段落及表格行定位；未建立固定分页，页码不作推断。':'PDF 保留原页码与文字坐标'};
}

export function classify(name, text='') {
  const s=name+'\n'+text.slice(0,500);
  if(/测试路径|验收预期|人工核对|测试方式|PRD|^19[_\s]/i.test(name))return 'excluded';
  if(/检测清单/.test(s))return /评标/.test(s)?'evaluation_rules':'tender_rules';
  if(/采购制度|制度节选/.test(s))return 'policy';
  if(/评标报告/.test(s))return /空白|可编辑模板/.test(s)?'evaluation_template':'evaluation';
  if(/投标文件[_＿\s]?[A-D]|采购投标文件/.test(s))return 'bid';
  if(/资格要求|资格条件_已确认/.test(name))return 'qualification';
  if(/技术需求书/.test(name))return 'technical';
  if(/合同/.test(name))return 'contract';
  if(/审核函/.test(name))return 'letter';
  if(/项目.*登记/.test(name))return 'project';
  if(/招标文件/.test(s))return /空白|可编辑模板/.test(s)?'tender_template':'tender';
  return 'other';
}

const stageKinds={compile:['project','technical','qualification','contract','tender_template','letter'],tender:['project','technical','qualification','contract','tender_template','tender_rules'],evaluation:['project','tender','evaluation_rules'],bids:['tender','bid'],qa:['policy'],archive:['evaluation','project']};
export function selectDocuments(docs,stage,targetId) {
  if(!stageKinds[stage])throw new Error('未知业务环节');
  const target=docs.find(d=>d.id===targetId); const selected=[],excluded=[];
  const latest=new Map();
  for(const d of docs.filter(d=>d.active && d.confirmed && d.kind!=='excluded')) {
    if(!latest.has(d.kind) || d.created>latest.get(d.kind).created)latest.set(d.kind,d);
  }
  for(const d of docs) {
    let reason='';
    if(d.kind==='excluded')reason='验收／测试说明，仅供人工核对';
    else if(!d.active)reason='文档已撤下';
    else if(d.id===targetId) {selected.push({...d,selectionReason:'本次目标文档'});continue;}
    else if(d.kind===target?.kind)reason='同类其他版本，防止答案及未来版本泄漏';
    else if(!stageKinds[stage].includes(d.kind))reason='与当前环节无关';
    else if(d.kind!=='bid'&&!d.confirmed)reason='未确认为基准／正式版本';
    else if(d.kind!=='bid' && latest.has(d.kind) && latest.get(d.kind).id!==d.id)reason='已由最新确认版本替代';
    if(reason)excluded.push({id:d.id,name:d.name,kind:d.kind,reason});
    else selected.push({...d,selectionReason:'当前环节适用的有效材料'});
  }
  return {stage,targetId,selected,excluded};
}
export function extractRules(doc) {
  const candidates=doc.fragments.filter(f=>f.cells && /^[RE]\d{2}$/.test(f.cells[0]));
  const ruleTable=candidates[0]?.location.split(' 行 ')[0];
  const rules=candidates.filter(f=>f.location.split(' 行 ')[0]===ruleTable).map(f=>({id:f.cells[0],condition:f.cells[1],source:f.cells[2],enabled:true,severity:['R09','E04'].includes(f.cells[0])?'高':'中'}));
  if(!rules.length)throw new Error('未找到规则表格，第一列应为 R01／E01 格式');
  if(new Set(rules.map(r=>r.id)).size!==rules.length)throw new Error('存在重复规则编号');
  return rules;
}
