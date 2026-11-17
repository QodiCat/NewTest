import {readFileSync} from 'node:fs';
import {fileURLToPath} from 'node:url';
import {parse} from 'dotenv';
const envPath=fileURLToPath(new URL('../.env',import.meta.url));
const inheritedEnv={...process.env};
function readSettings(){
 let local={};try{local=parse(readFileSync(envPath));}catch(error){if(error.code!=='ENOENT')throw new Error('无法读取项目根目录 .env，请检查文件权限');}
 const value=(name,old)=>local[name]??local[old]??inheritedEnv[name]??inheritedEnv[old];
 const baseUrl=String(value('AI_BASE_URL','LLM_BASE_URL')||'').trim(),model=String(value('AI_MODEL','LLM_MODEL')||'').trim(),apiKey=String(value('AI_API_KEY','LLM_API_KEY')||'').trim();
 return {baseUrl,model,apiKey,enabled:(local.AI_ENABLED??inheritedEnv.AI_ENABLED)!=='false'&&!!baseUrl&&!!model,timeout:Math.max(1000,Number(value('AI_TIMEOUT_MS','LLM_TIMEOUT_MS'))||120000),configSource:Object.keys(local).length?'.env':'环境变量'};
}
export function modelConfig(){const {apiKey,...publicConfig}=readSettings();return {...publicConfig,hasKey:!!apiKey};}
export async function askModel(system,user,history=[]) {
 const c=readSettings();if(!c.enabled)throw new Error('请在项目根目录 .env 配置 AI_BASE_URL、AI_MODEL 和所需的 AI_API_KEY');
 if(!c.baseUrl||!c.model)throw new Error('缺少 AI_BASE_URL 或 AI_MODEL');
 const context=history.filter(m=>['user','assistant'].includes(m.role)&&typeof m.content==='string').slice(-20).map(m=>({role:m.role,content:m.content.slice(0,12000)}));
 const endpoint=c.baseUrl.replace(/\/+$/,'');
 const response=await fetch(endpoint.endsWith('/chat/completions')?endpoint:endpoint+'/chat/completions',{method:'POST',headers:{'Content-Type':'application/json',...(c.apiKey?{Authorization:`Bearer ${c.apiKey}`}:{})},body:JSON.stringify({model:c.model,messages:[{role:'system',content:system},...context,{role:'user',content:user}],temperature:0.1}),signal:AbortSignal.timeout(c.timeout)});
 if(!response.ok)throw new Error(`模型接口返回 HTTP ${response.status}，请检查服务、配额及模型配置`);
 const json=await response.json();const result=json.choices?.[0]?.message?.content;
 if(typeof result!=='string'||!result.trim())throw new Error('模型未返回有效的回答内容');return result;
}
