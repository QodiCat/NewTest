import AdmZip from 'adm-zip';
import {XMLParser} from 'fast-xml-parser';
import {readFile} from 'node:fs/promises';
import {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,HeadingLevel,WidthType} from 'docx';

const xml=new XMLParser({preserveOrder:true,ignoreAttributes:false,trimValues:false});
const child=(nodes,key)=>nodes?.find(n=>n[key])?.[key]||[];
const attr=(nodes,key,name='w:val')=>nodes?.find(n=>n[key])?.[':@']?.['@_'+name];
const has=(nodes,key)=>nodes?.some(n=>n[key]&& !['0','false','off','none'].includes(String(n[':@']?.['@_w:val'])));
function find(nodes,key){for(const n of nodes||[]){if(n[key])return n[key];for(const v of Object.values(n))if(Array.isArray(v)){const found=find(v,key);if(found)return found;}}}
function runs(nodes,style={}){
 const out=[];
 for(const n of nodes||[]){
  if(n['w:r']){const p=child(n['w:r'],'w:rPr');out.push(...runs(n['w:r'],{bold:has(p,'w:b'),italic:has(p,'w:i'),underline:has(p,'w:u')}));}
  else if(n['w:t'])out.push({text:n['w:t'].map(t=>t['#text']??'').join(''),...style});
  else if(n['w:br']||n['w:cr'])out.push({text:'\n',...style});
  else if(n['w:tab'])out.push({text:'\t',...style});
  else for(const [k,v] of Object.entries(n))if(Array.isArray(v)&&!['w:pPr','w:rPr','w:del','w:instrText'].includes(k))out.push(...runs(v,style));
 }
 return out;
}
function blocks(nodes){return (nodes||[]).flatMap(n=>{
 if(n['w:p']){const p=child(n['w:p'],'w:pPr'),style=attr(p,'w:pStyle')||'',heading=style.match(/(?:heading|标题)\s*([1-3])/i);return [{type:'paragraph',tag:heading?'h'+heading[1]:/title|标题/i.test(style)?'h1':'p',align:attr(p,'w:jc')||'left',list:child(p,'w:numPr').length?'bullet':null,runs:runs(n['w:p'])}];}
 if(n['w:tbl'])return [{type:'table',rows:n['w:tbl'].filter(r=>r['w:tr']).map(r=>r['w:tr'].filter(c=>c['w:tc']).map(c=>({colSpan:Number(attr(child(c['w:tc'],'w:tcPr'),'w:gridSpan'))||1,blocks:blocks(c['w:tc'])})))}];
 return [];
});}
export async function readRichDocx(file){const zip=new AdmZip(await readFile(file));return normalizeContent(blocks(find(xml.parse(zip.readAsText('word/document.xml')),'w:body')));}

export function normalizeContent(input){
 let count=0,characters=0;
 const visit=(items,depth=0)=>{
  if(!Array.isArray(items)||depth>3)throw new Error('文档结构无效');
  return items.map(b=>{
   if(++count>15000)throw new Error('文档内容过多');
   if(b?.type==='table'){
    if(!Array.isArray(b.rows)||!b.rows.length||b.rows.length>300)throw new Error('表格最多支持 300 行');
    return {type:'table',rows:b.rows.map(row=>{if(!Array.isArray(row)||!row.length||row.length>30)throw new Error('表格最多支持 30 列');return row.map(c=>({colSpan:Math.max(1,Math.min(30,Number(c.colSpan)||1)),blocks:visit(c.blocks,depth+1)}));})};
   }
   if(b?.type!=='paragraph'||!Array.isArray(b.runs))throw new Error('段落格式无效');
   return {type:'paragraph',tag:['h1','h2','h3'].includes(b.tag)?b.tag:'p',align:['left','center','right','justify'].includes(b.align)?b.align:'left',list:['bullet','number'].includes(b.list)?b.list:null,runs:b.runs.map(r=>{if(++count>15000||typeof r.text!=='string')throw new Error('文字格式无效');characters+=r.text.length;if(characters>300000)throw new Error('文档文字过多');return {text:r.text,bold:!!r.bold,italic:!!r.italic,underline:!!r.underline};})};
  });
 };
 return visit(input);
}
export async function richDocx(input){
 const content=normalizeContent(input);
 const convert=items=>items.map(b=>b.type==='table'?new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:b.rows.map(row=>new TableRow({children:row.map(c=>new TableCell({columnSpan:c.colSpan,children:[...convert(c.blocks),new Paragraph('')]}))}))}):new Paragraph({heading:b.tag==='h1'?HeadingLevel.HEADING_1:b.tag==='h2'?HeadingLevel.HEADING_2:b.tag==='h3'?HeadingLevel.HEADING_3:undefined,alignment:b.align,bullet:b.list==='bullet'?{level:0}:undefined,numbering:b.list==='number'?{reference:'editor-numbered',level:0}:undefined,spacing:{after:160},children:b.runs.flatMap(r=>r.text.split('\n').map((text,i)=>new TextRun({text,break:i?1:undefined,bold:r.bold,italics:r.italic,underline:r.underline?{}:undefined,font:'Microsoft YaHei',size:22})))}));
 const buffer=await Packer.toBuffer(new Document({numbering:{config:[{reference:'editor-numbered',levels:[{level:0,format:'decimal',text:'%1.',alignment:'left'}]}]},sections:[{children:convert(content)}]}));
 return {buffer,content};
}
