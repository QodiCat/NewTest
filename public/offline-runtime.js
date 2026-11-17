export function installOfflineRuntime(seed,helpers){
 const {checkRules,compareBids,responseMatrix,selectDocuments,extractRules,createDocxArchive}=helpers;
 const records={project:[seed.project],document:seed.documents,ruleset:seed.rulesets,task:[],issue:[],notification:[],audit:[],comparison:[],review:[]},assets={...seed.assets};let user=null;
 const clone=x=>structuredClone(x),all=t=>records[t]||[],find=(t,id)=>all(t).find(x=>x.id===id),put=(t,x)=>{const i=all(t).findIndex(y=>y.id===x.id);i<0?all(t).push(x):all(t)[i]=x;return x;},id=()=>crypto.randomUUID?.()||'local-'+Date.now()+'-'+Math.random().toString(16).slice(2),now=()=>new Date().toISOString();
 const json=(x,status=200)=>new Response(JSON.stringify(x),{status,headers:{'Content-Type':'application/json'}}),error=(message,status=400)=>Object.assign(new Error(message),{status});
 const bytes=s=>Uint8Array.from(atob(s),c=>c.charCodeAt(0));
 const project=pid=>{const p=find('project',pid);if(!p||p.deletedAt)throw error('项目不存在或已删除',404);return p;};
 const docs=pid=>all('document').filter(d=>d.projectId===pid),selection=(pid,stage,targetId)=>selectDocuments(docs(pid),stage,targetId);
 const log=(pid,action,detail)=>put('audit',{id:id(),projectId:pid,created:now(),user:user?.name||'系统',action,detail});
 const deny=()=>{if(['audit','agency'].includes(user.role))throw error('当前角色无操作权限',403);};
 const people={reviewer:['张审核','reviewer'],business:['李经办','business'],agency:['代理经办','agency'],audit:['监督审计','audit'],admin:['系统管理员','admin']};
 const richXml=blocks=>(blocks||[]).map(b=>b.type==='table'?'<w:tbl><w:tblPr><w:tblBorders>'+['top','left','bottom','right','insideH','insideV'].map(k=>`<w:${k} w:val="single" w:sz="4"/>`).join('')+'</w:tblBorders></w:tblPr>'+b.rows.map(row=>'<w:tr>'+row.map(c=>'<w:tc><w:tcPr/>'+richXml(c.blocks)+'</w:tc>').join('')+'</w:tr>').join('')+'</w:tbl>':'<w:p><w:pPr><w:jc w:val="'+(b.align||'left')+'"/></w:pPr>'+(b.runs||[]).map(r=>'<w:r><w:rPr>'+(r.bold?'<w:b/>':'')+(r.italic?'<w:i/>':'')+(r.underline?'<w:u w:val="single"/>':'')+'</w:rPr><w:t xml:space="preserve">'+String(r.text||'').replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]))+'</w:t></w:r>').join('')+'</w:p>').join('');
 async function editDoc(source,content){
  const d={...clone(source),id:id(),created:now(),editorContent:content,version:source.version+1};
  d.text=content.map(b=>b.type==='table'?b.rows.map(r=>r.map(c=>c.blocks.map(p=>p.runs.map(x=>x.text).join('')).join(' ')).join(' | ')).join('\n'):b.runs.map(r=>r.text).join('')).join('\n');
  d.fragments=d.text.split('\n').map((text,i)=>({id:'p'+(i+1),text,section:'正文',location:'段落 '+(i+1),page:null}));d.hash=null;
  assets[d.id]=await createDocxArchive({'[Content_Types].xml':'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>','_rels/.rels':'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>','word/document.xml':'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'+richXml(content)+'<w:sectPr/></w:body></w:document>'}).arrayBuffer();d.assetId=d.id;put('document',d);return d;
 }
 async function handle(url,options={}){
  const path=url.pathname.slice(4),parts=path.split('/').filter(Boolean),method=options.method||'GET',body=typeof options.body==='string'?JSON.parse(options.body):{},pid=parts[0]==='projects'?parts[1]:null;
  if(path==='/login'){const p=people[body.username];if(!p||body.password!=='LocalDemo2026!')throw error('账号或密码错误',401);user={id:body.username,name:p[0],role:p[1]};return json(user);}
  if(!user)throw error('请登录本地工作台',401);
  if(path==='/me')return json(user);if(path==='/logout'){user=null;return json({ok:true});}
  if(path==='/config')return json({enabled:false,baseUrl:'',model:'',hasKey:false,timeout:120000,configSource:'单文件离线版'});
  if(path==='/model/test'||parts.at(-1)==='qa')throw error('真实 AI 需使用本地服务：启动 start.ps1 后打开 http://127.0.0.1:3000，后端会读取项目 .env。',503);
  if(path==='/projects'&&method==='GET')return json(all('project').filter(p=>!p.deletedAt));
  if(path==='/projects'&&method==='POST'){deny();return json(put('project',{...body,id:id(),created:now(),stage:'草稿',members:Object.keys(people)}));}
  if(pid)project(pid);
  if(parts[0]==='projects'&&parts.length===2){if(method==='DELETE'){deny();put('project',{...project(pid),deletedAt:now()});return json({ok:true});}return json({...project(pid),documents:docs(pid).map(d=>({...d,fragmentCount:d.fragments.length})),tasks:all('task').filter(t=>t.projectId===pid),issues:all('issue').filter(i=>i.projectId===pid),notifications:all('notification').filter(n=>n.projectId===pid)});}
  if(parts[0]==='documents'){
   const d=find('document',parts[1]);if(!d)throw error('文档不存在',404);project(d.projectId);
   if(parts[2]==='download'){const asset=assets[d.assetId||d.id];if(!asset)throw error('此文件仅登记了名称，暂无可下载原件');return new Response(typeof asset==='string'?bytes(asset):asset,{headers:{'Content-Type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document','Content-Disposition':`attachment; filename*=UTF-8''${encodeURIComponent(d.name)}`}});}
   if(method==='PATCH'){deny();put('document',{...d,...body});return json({ok:true});}return json(d);
  }
  if(parts[2]==='register-files'){deny();const imported=[];for(const f of body.files||[]){const src=seed.documents.find(d=>d.name===f.name);const d={...(src?clone(src):{kind:'other',fragments:[],text:'',confirmed:false}),id:id(),projectId:pid,name:f.name,active:true,created:now(),version:docs(pid).length+1,assetId:src?.id};put('document',d);imported.push(d);if(['tender_rules','evaluation_rules'].includes(d.kind))put('ruleset',{id:id(),projectId:pid,stage:d.kind==='tender_rules'?'tender':'evaluation',sourceId:d.id,rules:extractRules(d),version:now(),created:now()});}log(pid,'登记文件',{names:imported.map(d=>d.name)});return json(imported);}
  if(parts[2]==='selection')return json(selection(pid,url.searchParams.get('stage'),url.searchParams.get('targetId')));
  if(parts[2]==='rules'){if(method==='POST'){deny();return json(put('ruleset',{...body,id:id(),projectId:pid,version:now(),created:now()}));}return json(all('ruleset').filter(r=>r.projectId===pid));}
  if(parts[2]==='tasks'){
   deny();const previous=find('task',body.parentTaskId),rules=find('ruleset',body.rulesetId)||previous?.rules||all('ruleset').filter(r=>r.projectId===pid&&r.stage===body.stage).at(-1);if(!rules)throw error('请先导入本环节检测清单');
   const selected=selection(pid,body.stage,body.targetId);if(previous&&!body.refreshBaseline)selected.selected=[docs(pid).find(d=>d.id===body.targetId),...previous.selection.selected.filter(d=>d.id!==previous.targetId)];
   const results=checkRules(selected,rules.rules),summary={total:results.length,pass:results.filter(r=>r.status==='通过').length,fail:results.filter(r=>r.status==='不通过').length,unknown:results.filter(r=>r.status==='无法核验').length};
   const task=put('task',{id:id(),projectId:pid,stage:body.stage,targetId:body.targetId,targetName:find('document',body.targetId).name,selection:clone(selected),rules:clone(rules),results,summary,created:now(),completed:now(),status:'已完成',progress:100,parserVersion:'offline-1.0',parentTaskId:body.parentTaskId});
   for(const issue of all('issue').filter(i=>i.taskId===body.parentTaskId)){issue.status=results.find(r=>r.id===issue.ruleId)?.status==='通过'?'已解决':'复检未通过';}log(pid,'检测完成',{taskId:task.id,summary});return json(task);
  }
  if(parts[0]==='tasks'){const t=find('task',parts[1]);if(!t)throw error('任务不存在',404);project(t.projectId);if(parts[2]==='issues'){deny();const r=t.results.find(r=>r.id===body.ruleId);return json(put('issue',{id:id(),projectId:t.projectId,taskId:t.id,ruleId:r.id,status:'待确认',reason:r.reason,created:now()}));}if(parts[2]==='reviews')return json(all('review').filter(r=>r.taskId===t.id));if(parts[2]==='results'){deny();put('review',{id:id(),projectId:t.projectId,taskId:t.id,ruleId:parts[3],...body,created:now()});return json({ok:true});}return json(t);}
  if(parts[0]==='issues'){deny();const i=find('issue',parts[1]);if(!i)throw error('问题不存在');if(parts[2]==='notify'){if(!i.notificationId){const n=put('notification',{id:id(),issueId:i.id,projectId:i.projectId,created:now(),status:'已送达'});i.notificationId=n.id;i.status='待整改';}return json(find('notification',i.notificationId));}Object.assign(i,{feedback:body,status:'已反馈'});return json({ok:true});}
  if(parts[2]==='compile'){
   deny();const sources=selection(pid,'compile').selected;const sourceFor=n=>docs(pid).find(d=>d.name===n)||seed.documents.find(d=>d.name===n);
   let draft={...clone(sourceFor('08_招标文件_送审稿.docx')),id:id(),projectId:pid,created:now(),version:1},letter={...clone(sourceFor('05_招标文件审核函.docx')),id:id(),projectId:pid,created:now(),version:1};
   draft.assetId=seed.documents.find(d=>d.name==='08_招标文件_送审稿.docx').id;letter.assetId=seed.documents.find(d=>d.name==='05_招标文件审核函.docx').id;put('document',draft);put('document',letter);
   if(body.draftContent)draft=await editDoc(draft,body.draftContent);if(body.letterContent)letter=await editDoc(letter,body.letterContent);
   return json({documentId:draft.id,letterId:letter.id,text:draft.text,letterText:letter.text,draftContent:draft.editorContent,letterContent:letter.editorContent,pending:[],sources});
  }
  if(parts[2]==='compare'){deny();const s=selection(pid,'bids');return json(put('comparison',{id:id(),projectId:pid,pairs:compareBids(s),matrix:responseMatrix(s),created:now()}));}
  if(parts[2]==='comparison')return json(all('comparison').filter(c=>c.projectId===pid).at(-1)||null);
  if(parts[0]==='comparisons'){deny();const c=find('comparison',parts[1]),pair=c.pairs.find(p=>p.id===body.pairId);Object.assign(pair,{status:'已排除',reason:body.reason});return json(c);}
  if(parts[2]==='publish'){deny();put('project',{...project(pid),stage:'已发布',publication:body});return json({ok:true});}
  if(parts[2]==='award'){const t=all('task').filter(t=>t.projectId===pid&&t.stage==='evaluation').at(-1),d=t?.selection.selected.find(d=>d.id===t.targetId);return json({taskId:t?.id,blocked:!t||t.summary.fail>0||t.summary.unknown>0,reportConfirmed:true,decision:project(pid).decision,issues:all('issue').filter(i=>i.projectId===pid&&i.status!=='已解决'),candidates:d?.fragments.filter(f=>f.cells?.length===6&&/^[A-Z]$/.test(f.cells[0])).map(f=>({supplier:f.cells[0],score:+f.cells[4],rank:+f.cells[5],evidence:{documentId:d.id,fragmentId:f.id,name:d.name,location:f.location,text:f.text}})).sort((a,b)=>a.rank-b.rank)||[]});}
  if(parts[2]==='archive'){deny();put('project',{...project(pid),stage:'已归档',decision:{...body,user:user.name,created:now()}});return json({ok:true});}
  if(parts[2]==='audit')return json(all('audit').filter(a=>a.projectId===pid));
  if(parts[2]==='export')return new Response(JSON.stringify({project:project(pid),tasks:all('task').filter(t=>t.projectId===pid)},null,2),{headers:{'Content-Type':'application/json','Content-Disposition':'attachment; filename="audit-export.json"'}});
  throw error('离线版暂不支持该操作');
 }
 const nativeFetch=window.fetch.bind(window);window.fetch=async(input,options)=>{const value=typeof input==='string'?input:input.url;if(!value.startsWith('/api/'))return nativeFetch(input,options);try{return await handle(new URL(value,'http://offline.local'),options);}catch(e){return json({error:e.message},e.status||400);}};
 document.addEventListener('click',async e=>{const a=e.target.closest('a');if(!a?.getAttribute('href')?.startsWith('/api/'))return;e.preventDefault();e.stopImmediatePropagation();const response=await fetch(a.getAttribute('href'));if(!response.ok){alert((await response.json()).error);return;}const blob=await response.blob(),url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download=a.download||decodeURIComponent(response.headers.get('Content-Disposition')?.match(/UTF-8''(.+)/)?.[1]||response.headers.get('Content-Disposition')?.match(/filename="([^"]+)"/)?.[1]||'download.docx');document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);},true);
}
