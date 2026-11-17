import test from 'node:test';
import assert from 'node:assert/strict';
import {DemoSession} from '../public/demo-session.js';
test('login hides historical files and all results; upload and analysis reveal independently',()=>{
 const d=new DemoSession();d.reset('reviewer');
 const data={id:'p',documents:[{id:'doc1'},{id:'doc2'}],tasks:[{id:'old'},{id:'new'}],issues:[{id:'oldIssue'},{id:'newIssue'}],notifications:[{issueId:'oldIssue'}],publication:{url:'previous'},decision:{supplier:'previous'}};
 let visible=d.projectView(data);assert.equal(visible.documents.length,0);assert.equal(visible.tasks.length,0);assert.equal(visible.issues.length,0);assert.equal(visible.notifications.length,0);assert.equal(visible.decision,null);assert.equal(visible.publication,null);
 assert.deepEqual(d.selectionView({selected:data.documents,excluded:[{id:'doc3'}]},'p'),{selected:[],excluded:[]});
 d.revealDocuments('p',data.documents);visible=d.projectView(data);assert.equal(visible.documents.length,2);assert.equal(visible.tasks.length,0);
 d.recordTask({id:'new',projectId:'p'});d.project('p').issues.push('newIssue');visible=d.projectView(data);assert.deepEqual(visible.tasks,[{id:'new'}]);assert.deepEqual(visible.issues,[{id:'newIssue'}]);
 const refresh=new DemoSession(JSON.parse(JSON.stringify(d.state)));assert.equal(refresh.projectView(data).tasks.length,1);
 refresh.reset('reviewer');visible=refresh.projectView(data);assert.equal(visible.documents.length,0);assert.equal(visible.tasks.length,0);assert.equal(visible.issues.length,0);
 assert.equal(data.documents.length,2);assert.equal(data.tasks.length,2);
});
