const escapeText=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const plainContent=text=>{
 const blocks=[];let table=null;
 for(const line of String(text??'').split('\n')){
  if(line.includes(' | ')){
   if(!table){table={type:'table',rows:[]};blocks.push(table);}
   table.rows.push(line.split(' | ').map(text=>({colSpan:1,blocks:[{type:'paragraph',tag:'p',runs:[{text}]}]})));
  }else{table=null;blocks.push({type:'paragraph',tag:!blocks.length?'h1':/^(?:\d+\.\s|附件\s*[A-Z\d])/.test(line)?'h2':'p',runs:[{text:line}]});}
 }
 return blocks;
};
function htmlContent(blocks){
 let html='',listTag=null;const closeList=()=>{if(listTag){html+=`</${listTag}>`;listTag=null;}};
 for(const b of blocks){
  if(b.type==='table'){closeList();html+=`<table><tbody>${b.rows.map(row=>`<tr>${row.map(c=>`<td colspan="${Number(c.colSpan)||1}">${htmlContent(c.blocks)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;continue;}
  const tag=['h1','h2','h3'].includes(b.tag)?b.tag:'p';
  const text=b.runs.map(r=>{let s=escapeText(r.text).replace(/\n/g,'<br>');if(r.bold)s=`<strong>${s}</strong>`;if(r.italic)s=`<em>${s}</em>`;if(r.underline)s=`<u>${s}</u>`;return s;}).join('')||'<br>';
  const paragraph=`<${tag} style="text-align:${['center','right','justify'].includes(b.align)?b.align:'left'}">${text}</${tag}>`;
  if(b.list){const next=b.list==='number'?'ol':'ul';if(listTag!==next){closeList();listTag=next;html+=`<${next}>`;}html+=`<li>${paragraph}</li>`;}else{closeList();html+=paragraph;}
 }
 closeList();return html;
}
function inline(node,marks={}){
 if(node.nodeType===3)return [{text:node.textContent,...marks}];
 if(node.nodeType!==1)return [];
 if(node.tagName==='BR')return [{text:'\n',...marks}];
 const style=node.style;
 const next={bold:marks.bold||['B','STRONG'].includes(node.tagName)||/bold|[6-9]00/.test(style.fontWeight),italic:marks.italic||['I','EM'].includes(node.tagName)||style.fontStyle==='italic',underline:marks.underline||node.tagName==='U'||style.textDecoration.includes('underline')};
 return [...node.childNodes].flatMap(n=>inline(n,next));
}
function readBlocks(parent,list=null){
 const result=[];let pending=[];
 const flush=()=>{if(pending.length){result.push({type:'paragraph',tag:'p',list,runs:pending});pending=[];}};
 for(const node of parent.childNodes){
  if(node.nodeType===3){if(node.textContent.trim())pending.push(...inline(node));continue;}
  if(node.nodeType!==1)continue;
  const tag=node.tagName;
  if(tag==='TABLE'){flush();result.push({type:'table',rows:[...node.rows].map(row=>[...row.cells].map(c=>({colSpan:c.colSpan,blocks:readBlocks(c)})))});}
  else if(['UL','OL'].includes(tag)){flush();result.push(...readBlocks(node,tag==='OL'?'number':'bullet'));}
  else if(['P','DIV','H1','H2','H3','LI','BLOCKQUOTE'].includes(tag)){
   flush();
   if(node.querySelector('table,p,div,ul,ol,h1,h2,h3'))result.push(...readBlocks(node,list));
   else result.push({type:'paragraph',tag:/^H[1-3]$/.test(tag)?tag.toLowerCase():'p',align:node.style.textAlign||'left',list,runs:[...node.childNodes].flatMap(n=>inline(n))});
  }else pending.push(...inline(node));
 }
 flush();return result;
}
export function mountRichEditor(host,{id,title,content,text,downloadId}){
 const button=(cmd,label,html=label)=>`<button type="button" data-edit="${cmd}" aria-label="${label}" title="${label}">${html}</button>`;
 host.innerHTML=`<div class="document-editor"><div class="document-editor-heading"><div><span class="document-kicker">DOCUMENT EDITOR</span><h2>${escapeText(title)}</h2></div><span class="editor-save-state">已保存</span></div><div class="editor-toolbar" role="toolbar" aria-label="${escapeText(title)}格式工具"><select aria-label="段落样式" data-style><option value="p">正文</option><option value="h1">标题 1</option><option value="h2">标题 2</option><option value="h3">标题 3</option></select><span class="editor-separator"></span>${button('bold','加粗','<b>B</b>')}${button('italic','斜体','<i>I</i>')}${button('underline','下划线','<u>U</u>')}<span class="editor-separator"></span>${button('insertUnorderedList','项目符号','• 列表')}${button('insertOrderedList','编号列表','1. 列表')}${button('justifyLeft','左对齐','左对齐')}${button('justifyCenter','居中','居中')}<span class="editor-separator"></span>${button('table','插入表格','＋ 表格')}${button('row','在下方插入行','＋ 行')}${button('column','在右侧插入列','＋ 列')}${button('removeRow','删除当前行','删行')}${button('undo','撤销','↶')}${button('redo','重做','↷')}</div><div class="editor-canvas"><div id="${id}" class="rich-document" contenteditable="true" role="textbox" aria-multiline="true" aria-label="${escapeText(title)}正文" spellcheck="false">${htmlContent(content||plainContent(text))}</div></div><div class="editor-footer"><span>选中文字设置格式 · 点击表格可编辑单元格</span><a class="button" href="/api/documents/${downloadId}/download">下载已保存的 DOCX</a></div></div>`;
 const editor=host.querySelector('.rich-document'),toolbar=host.querySelector('.editor-toolbar'),status=host.querySelector('.editor-save-state');
 let range=null,baseline=JSON.stringify(readBlocks(editor)),dirty=false;
 const remember=()=>{const s=window.getSelection();if(s.rangeCount&&editor.contains(s.anchorNode)){range=s.getRangeAt(0).cloneRange();for(const b of toolbar.querySelectorAll('[data-edit]')){const cmd=b.dataset.edit;if(['bold','italic','underline','insertUnorderedList','insertOrderedList','justifyLeft','justifyCenter'].includes(cmd)){const active=document.queryCommandState(cmd);b.classList.toggle('is-active',active);b.setAttribute('aria-pressed',String(active));}}const style=String(document.queryCommandValue('formatBlock')).toLowerCase();toolbar.querySelector('[data-style]').value=['h1','h2','h3'].includes(style)?style:'p';}};
 const changed=()=>{dirty=JSON.stringify(readBlocks(editor))!==baseline;status.textContent=dirty?'有未保存修改':'已保存';status.classList.toggle('is-dirty',dirty);};
 const restore=()=>{editor.focus();if(range&&editor.contains(range.commonAncestorContainer)){const s=window.getSelection();s.removeAllRanges();s.addRange(range);}};
 editor.addEventListener('input',()=>{remember();changed();});
 editor.addEventListener('keyup',remember);editor.addEventListener('mouseup',remember);
 editor.addEventListener('paste',e=>{e.preventDefault();document.execCommand('insertText',false,e.clipboardData.getData('text/plain'));remember();changed();});
 toolbar.addEventListener('mousedown',e=>{if(e.target.closest('button')){remember();e.preventDefault();}});
 toolbar.querySelector('select').addEventListener('mousedown',remember);
 toolbar.querySelector('select').addEventListener('change',e=>{restore();document.execCommand('formatBlock',false,e.target.value);remember();changed();});
 toolbar.addEventListener('click',e=>{
  const cmd=e.target.closest('[data-edit]')?.dataset.edit;if(!cmd)return;restore();
  const selected=window.getSelection()?.anchorNode,cell=(selected?.nodeType===1?selected:selected?.parentElement)?.closest('td,th');
  if(cmd==='table')document.execCommand('insertHTML',false,'<table><tbody><tr><td><p>列标题 1</p></td><td><p>列标题 2</p></td><td><p>列标题 3</p></td></tr><tr><td><p><br></p></td><td><p><br></p></td><td><p><br></p></td></tr></tbody></table><p><br></p>');
  else if(['row','column','removeRow'].includes(cmd)){
   if(!cell||!editor.contains(cell)){status.textContent='请先点击要修改的表格单元格';return;}
   const row=cell.parentElement,table=cell.closest('table');
   if(cmd==='row'){const next=row.cloneNode(true);for(const c of next.cells)c.innerHTML='<p><br></p>';row.after(next);}
   if(cmd==='column'){if([...table.rows].some(r=>[...r.cells].some(c=>c.colSpan>1))){status.textContent='合并单元格的表格请先在 Word 中调整列';return;}const index=cell.cellIndex;for(const r of table.rows){const c=r.insertCell(Math.min(index+1,r.cells.length));c.innerHTML='<p><br></p>';}}
   if(cmd==='removeRow'){if(table.rows.length===1)table.remove();else row.remove();range=null;}
  }else document.execCommand(cmd,false);
  remember();changed();
 });
 return {content:()=>readBlocks(editor),changed:()=>dirty};
}
