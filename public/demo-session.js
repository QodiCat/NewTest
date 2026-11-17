export class DemoSession {
  constructor(saved) {this.state=saved?.version===1?saved:{version:1,userId:null,started:new Date().toISOString(),projects:{}};}
  reset(userId){this.state={version:1,userId,started:new Date().toISOString(),projects:{}};}
  project(id){return this.state.projects[id]??={documents:[],tasks:[],issues:[],comparison:null,publication:false,decision:false};}
  revealDocuments(pid,documents){const p=this.project(pid);p.documents=[...new Set([...p.documents,...documents.map(d=>d.id)])];}
  recordTask(task){const p=this.project(task.projectId);p.tasks=[...new Set([...p.tasks,task.id])];}
  projectView(data){const p=this.project(data.id);return {...data,documents:data.documents.filter(d=>p.documents.includes(d.id)),tasks:data.tasks.filter(t=>p.tasks.includes(t.id)),issues:data.issues.filter(i=>p.issues.includes(i.id)),notifications:data.notifications.filter(n=>p.issues.includes(n.issueId)),publication:p.publication?data.publication:null,decision:p.decision?data.decision:null};}
  selectionView(data,pid){const ids=this.project(pid).documents;return {...data,selected:data.selected.filter(d=>ids.includes(d.id)),excluded:data.excluded.filter(d=>ids.includes(d.id))};}
}
