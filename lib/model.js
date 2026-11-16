export function modelConfig(){return {enabled:process.env.LLM_ENABLED==='true',baseUrl:process.env.LLM_BASE_URL||'',model:process.env.LLM_MODEL||'',hasKey:!!process.env.LLM_API_KEY,timeout:Number(process.env.LLM_TIMEOUT_MS)||120000};}
export async function askModel(system,user) {
 const c=modelConfig();if(!c.enabled)throw new Error('模型未启用，请配置 .env 后重启');
 if(!c.baseUrl||!c.model)throw new Error('缺少 LLM_BASE_URL 或 LLM_MODEL');
 const response=await fetch(c.baseUrl.replace(/\/$/,'')+'/chat/completions',{method:'POST',headers:{'Content-Type':'application/json',...(process.env.LLM_API_KEY?{Authorization:`Bearer ${process.env.LLM_API_KEY}`}:{})},body:JSON.stringify({model:c.model,messages:[{role:'system',content:system},{role:'user',content:user}],temperature:0.1}),signal:AbortSignal.timeout(c.timeout)});
 if(!response.ok)throw new Error(`模型接口返回 HTTP ${response.status}，请检查服务、配额及模型配置`);
 const json=await response.json();const result=json.choices?.[0]?.message?.content;
 if(typeof result!=='string')throw new Error('模型未返回兼容的文本内容');return result;
}
