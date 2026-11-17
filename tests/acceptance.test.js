import {test} from 'node:test';
import assert from 'node:assert/strict';
import {spawn} from 'node:child_process';
import {mkdtemp,readdir} from 'node:fs/promises';
import path from 'node:path';
test('真实样例：检测、整改、隔离、问答及权限闭环',async t=>{
 const dir=await mkdtemp(path.resolve('artifacts-test-')),port=13000+Math.floor(Math.random()*3000);
 const child=spawn(process.execPath,['server.js'],{env:{...process.env,PORT:String(port),DATA_DIR:dir,LLM_ENABLED:'false'},stdio:['ignore','pipe','pipe']});
 let logs='';child.stdout.on('data',d=>logs+=d);child.stderr.on('data',d=>logs+=d);t.after(()=>child.kill());
 const base=`http://127.0.0.1:${port}/api`;for(let i=0;i<100;i++){try{await fetch(base+'/me');break;}catch{await new Promise(r=>setTimeout(r,100));}if(i===99)throw new Error(logs);}
 let cookie='';async function call(url,body,method){const r=await fetch(base+url,{method:method||(body?'POST':'GET'),headers:{'Content-Type':'application/json',cookie},body:body?JSON.stringify(body):undefined});if(url==='/login')cookie=r.headers.get('set-cookie')?.split(';')[0];const json=await r.json();return {status:r.status,...json};}
 assert.equal((await call('/login',{username:'reviewer',password:'LocalDemo2026!'})).status,200);
 const pid='DEMO-CCAS-2026-001',p=await call('/projects/'+pid);assert.equal(p.documents.length,18,logs);
 const doc=n=>p.documents.find(d=>d.name.startsWith(n+'_'));
 async function run(n,stage,extra={}){const job=await call(`/projects/${pid}/tasks`,{targetId:doc(n).id,stage,...extra});assert.equal(job.status==='待执行'||job.status==='检测中',true,JSON.stringify(job));for(let i=0;i<50;i++){const r=await call('/tasks/'+job.id);if(['已完成','部分完成','失败'].includes(r.status))return r;await new Promise(r=>setTimeout(r,50));}throw Error('超时');}
 const a=await run('08','tender');assert.deepEqual(a.summary,{pass:14,fail:1,unknown:0,total:15},JSON.stringify(a.results));const r09=a.results.find(r=>r.id==='R09');assert.equal(r09.evidence.some(e=>/至少具有 3 个/.test(e.text)),true);assert.equal(r09.evidence.some(e=>/至少具有 1 个/.test(e.text)),true);
 assert.equal(a.selection.excluded.some(d=>d.id===doc('09').id),true);
 const issue=await call('/tasks/'+a.id+'/issues',{ruleId:'R09'});const nt1=await call('/issues/'+issue.id+'/notify',{}),nt2=await call('/issues/'+issue.id+'/notify',{});assert.equal(nt1.id,nt2.id);
 const b=await run('09','tender',{parentTaskId:a.id});assert.deepEqual(b.summary,{pass:15,fail:0,unknown:0,total:15});assert.equal((await call('/projects/'+pid)).issues[0].status,'已解决');
 await call('/documents/'+doc('03').id,{active:false},'PATCH');const missing=await run('08','tender');assert.equal(missing.results.find(r=>r.id==='R09').status,'无法核验');assert.equal((await call('/tasks/'+a.id)).summary.fail,1);await call('/documents/'+doc('03').id,{active:true},'PATCH');
 const e1=await run('16','evaluation');assert.deepEqual(e1.summary,{pass:6,fail:1,unknown:0,total:7},JSON.stringify(e1.results));const e2=await run('17','evaluation');assert.deepEqual(e2.summary,{pass:7,fail:0,unknown:0,total:7});
 const qa=await call(`/projects/${pid}/qa`,{question:'评标报告由谁签字？'});assert.match(qa.answer,/参加评标的委员会成员签字/);assert.match(qa.fragments[0].section,/第 7 条/);
 const comp=await call(`/projects/${pid}/compare`,{});assert.equal(comp.matrix.length,4);assert.ok(comp.pairs.some(p=>p.matches.some(m=>m.typo)));const cd=comp.pairs.find(pair=>pair.left===doc('12').id&&pair.right===doc('13').id);assert.equal(cd.matches.some(m=>m.type==='全文重复'&&m.evidence.some(e=>e.location.includes('2. 技术实施方案'))),false);
 const draft=await call(`/projects/${pid}/compile`,{});assert.equal(draft.status,200);assert.match(draft.text,/至少具有 1 个/);assert.doesNotMatch(draft.text,/至少具有 3 个/);
 const filesBefore=await readdir(dir+'/uploads');
 const registered=await call(`/projects/${pid}/register-files`,{files:[{name:'不读取内容.docx',size:123},{name:doc('08').name,size:456}]});assert.equal(registered.status,200);
 const registration=await call('/projects/'+pid);const unknownDoc=registration.documents.find(d=>d.name==='不读取内容.docx');assert.equal(unknownDoc.parserVersion,'metadata-only');assert.equal(unknownDoc.fragmentCount,0);assert.equal(unknownDoc.originalSize,123);
 const readyDoc=registration.documents.filter(d=>d.name===doc('08').name).at(-1);assert.equal(readyDoc.hash,doc('08').hash);assert.equal(readyDoc.preparedSourceId,doc('08').id);assert.deepEqual(await readdir(dir+'/uploads'),filesBefore);
 const pub=await call(`/projects/${pid}/publish`,{documentId:doc('09').id,platform:'指定演示平台',url:'https://example.test/demo',start:'2026-12-01T09:00:00+08:00',end:'2026-12-07T17:00:00+08:00',price:0});assert.equal(pub.status,200);
 const award=await call(`/projects/${pid}/archive`,{supplier:'A',reason:'确认版报告评分及候选人排序'});assert.equal(award.status,200);assert.equal((await call('/projects/'+pid)).stage,'已归档');
 const disposable=await call('/projects',{name:'删除功能验收项目',code:'DELETE-TEST-001',department:'测试部门',owner:'张审核'});assert.equal(disposable.status,200);
 await call(`/projects/${disposable.id}/register-files`,{files:[{name:doc('08').name,size:123}]});const disposableDoc=(await call('/projects/'+disposable.id)).documents[0];
 const retainedFiles=await readdir(dir+'/uploads');
 for(const username of ['business','audit','agency']){await call('/login',{username,password:'LocalDemo2026!'});assert.equal((await call('/projects/'+disposable.id,undefined,'DELETE')).status,403);}
 await call('/login',{username:'reviewer',password:'LocalDemo2026!'});assert.equal((await call('/projects/'+disposable.id,undefined,'DELETE')).status,200);assert.equal((await call('/projects/'+disposable.id)).status,404);assert.equal((await call('/documents/'+disposableDoc.id)).status,404);assert.equal((await call('/projects/'+disposable.id+'/selection?stage=tender')).status,404);assert.equal((await call('/projects/'+disposable.id,undefined,'DELETE')).status,404);assert.deepEqual(await readdir(dir+'/uploads'),retainedFiles);
 const remainingProjects=await fetch(base+'/projects',{headers:{cookie}}).then(r=>r.json());assert.equal(remainingProjects.some(p=>p.id===disposable.id),false);assert.equal(remainingProjects.some(p=>p.id===pid),true);
 await call('/login',{username:'audit',password:'LocalDemo2026!'});assert.equal((await call('/documents/'+doc('03').id,{active:false},'PATCH')).status,403);await call('/login',{username:'agency',password:'LocalDemo2026!'});assert.equal((await call('/tasks/'+a.id)).status,403);assert.equal((await call('/documents/'+doc('03').id)).status,403);
});
