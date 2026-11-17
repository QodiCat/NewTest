// Use the established registration/rules/task APIs so this flow also works
// with servers started before the dedicated agent-repair endpoint existed.
export async function prepareReleaseAndRecheck({request,projectId,originalTaskId,cached,stage='tender'}) {
 const isEvaluation=stage==='evaluation';
 if(!['tender','evaluation'].includes(stage))throw new Error('当前环节不支持一键处理');
 const names=isEvaluation?['01_项目信息登记表.docx','09_招标文件_发布版.docx','15_评标报告检测清单.docx','17_评标报告_复核版.docx']:['01_项目信息登记表.docx','02_技术需求书.docx','03_供应商资格要求_已确认版.docx','04_招标文件_空白模板.docx','06_银行采购合同_模板.docx','07_招标文件检测清单_15项.docx','09_招标文件_发布版.docx'];
 const ruleCount=isEvaluation?7:15,ruleFilename=isEvaluation?names[2]:names[5];
 if(cached){
  const report=await request('/tasks/'+cached.taskId);
  return {...cached,filename:names.at(-1),report:{...report,agentRepair:true}};
 }
 const documents=await request(`/projects/${projectId}/register-files`,{files:names.map(name=>({name,size:0}))});
 const target=documents.find(d=>d.name===names.at(-1)&&!d.error);
 const rulesDocument=documents.find(d=>d.name===ruleFilename&&!d.error);
 if(!target?.id||!target.hash)throw new Error((isEvaluation?'17 复核版报告':'09 发布版文件')+'尚未就绪，请重新启动本地服务后重试');
 const rulesets=await request(`/projects/${projectId}/rules`);
 const rules=rulesets.findLast(r=>r.stage===stage&&r.sourceId===rulesDocument?.id)||rulesets.findLast(r=>r.stage===stage&&r.rules?.length===ruleCount);
 if(!rules)throw new Error(isEvaluation?'缺少评标报告 7 项核验规则，请先导入 15 号规则文件':'缺少招标文件 15 项检测规则，请先导入 07 号规则文件');
 let report=await request(`/projects/${projectId}/tasks`,{stage,targetId:target.id,rulesetId:rules.id,parentTaskId:originalTaskId,refreshBaseline:true,useModel:false});
 const deadline=Date.now()+120000;
 while(!['已完成','部分完成','失败'].includes(report.status)){
  if(Date.now()>deadline)throw new Error('复检任务仍在处理，请稍后从检测记录查看');
  await new Promise(resolve=>setTimeout(resolve,300));report=await request('/tasks/'+report.id);
 }
 if(report.status==='失败')throw new Error(report.error||'发布版复检失败');
 return {documentId:target.id,taskId:report.id,filename:target.name,documents,report:{...report,agentRepair:true}};
}
