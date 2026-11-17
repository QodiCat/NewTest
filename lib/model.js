export function modelConfig(){return {enabled:process.env.LLM_ENABLED==='true',baseUrl:process.env.LLM_BASE_URL||'',model:process.env.LLM_MODEL||'',hasKey:!!process.env.LLM_API_KEY,timeout:Number(process.env.LLM_TIMEOUT_MS)||120000};}
export async function askModel(system,user,history=[]) {
 const c=modelConfig();if(!c.enabled)throw new Error('模型未启用，请配置 .env 后重启');
 if(!c.baseUrl||!c.model)throw new Error('缺少 LLM_BASE_URL 或 LLM_MODEL');
 const context=history.filter(m=>['user','assistant'].includes(m.role)&&typeof m.content==='string').slice(-20).map(m=>({role:m.role,content:m.content.slice(0,12000)}));
 const response=await fetch(c.baseUrl.replace(/\/$/,'')+'/chat/completions',{method:'POST',headers:{'Content-Type':'application/json',...(process.env.LLM_API_KEY?{Authorization:`Bearer ${process.env.LLM_API_KEY}`}:{})},body:JSON.stringify({model:c.model,messages:[{role:'system',content:system},...context,{role:'user',content:user}],temperature:0.1}),signal:AbortSignal.timeout(c.timeout)});
 if(!response.ok)throw new Error(`模型接口返回 HTTP ${response.status}，请检查服务、配额及模型配置`);
 const json=await response.json();const result=json.choices?.[0]?.message?.content;
 if(typeof result!=='string'||!result.trim())throw new Error('模型未返回有效的回答内容');return result;
}
