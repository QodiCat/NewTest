const compact=s=>s.replace(/[\s|。；，、：:]/g,'');
const find=(d,re)=>d?.fragments.filter(f=>re.test(f.text))||[];
const evidence=(d,fs)=>fs.map(f=>({documentId:d.id,name:d.name,hash:d.hash,version:d.version,fragmentId:f.id,location:f.location,page:f.page,text:f.text}));
const missing='缺少可解析证据，请补充材料或人工核验';
export function checkRules(selection,rules) {
  const docs=selection.selected,t=docs.find(d=>d.id===selection.targetId);
  if(!t)throw new Error('请选择有效的目标文件');
  const by=k=>docs.find(d=>d.kind===k), all=t.fragments.map(f=>f.text).join('\n');
  const take=re=>find(t,re), has=re=>re.test(all);
  return rules.filter(r=>r.enabled).map(rule=>{
    let ok=null,reason=missing,tf=[],bf=[],b=null;
    const use=(re,baseline,bre=re)=>{tf=take(re);b=baseline;bf=find(b,bre);};
    const required=kind=>{b=by(kind);if(!b){reason=`缺少已确认的${kind}基准材料`;return false;}return true;};
    switch(rule.semantic?'CUSTOM':rule.id) {
      case 'R01':case 'R02':
        if(required('project')){use(rule.id==='R01'?/采购项目/:/项目编号/,b); const value=rule.id==='R02'?b.fragments.map(f=>f.text).join(' ').match(/(?:DEMO-|ZB-)[A-Z0-9-]+/)?.[0]:b.fragments.find(f=>/采购项目/.test(f.text))?.text.split(/项目编号|\|/)[0].trim();ok=value?all.includes(value):null;reason=value?`项目档案参照：${value}`:missing;}break;
      case 'R03':case 'R04':case 'R05':
        if(required('project')){const re=rule.id==='R03'?/招标人|代理机构/:rule.id==='R04'?/采购标的/:/获取|售价/;use(re,b);ok=rule.id==='R03'?has(/招标人[：:].+代理机构[：:]/):rule.id==='R04'?has(/信用卡审批系统一套/):has(/售价\s*0\s*元/)&&has(/\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}\s*至/) &&has(/平台/);reason='核对项目登记与公告字段';}break;
      case 'R06':{tf=take(/投标截止为/);const times=tf.flatMap(f=>[...f.text.matchAll(/投标截止为\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2})/g)].map(m=>m[1]));ok=times.length>=2?new Set(times).size===1:null;reason=times.length>=2?`公告／须知时间：${times.join('；')}`:missing;break;}
      case 'R07':tf=take(/提交 DOCX|提交.*PDF/);ok=tf.length?has(/平台.*提交.*(?:DOCX|PDF)/):null;reason='提交渠道与格式检查';break;
      case 'R08':tf=take(/澄清|有效期/);ok=tf.length?has(/澄清.*渠道|澄清.*平台|澄清通过平台/)&&has(/有效期.*\d+\s*个?自然日/):null;reason='澄清方式及有效期检查';break;
      case 'R09':
        if(required('qualification')) {
          tf=take(/^Q0[123]\s/);bf=find(b,/^Q0[123]\s*\|/).filter(f=>!f.text.includes("□"));
          if(tf.length>=3&&bf.length>=3) {
            const targetQ=tf.find(f=>/^Q02/.test(f.text)),baseQ=bf.find(f=>/^Q02/.test(f.text));
            const a=targetQ.text.match(/至少具有\s*(\d+)\s*个/),c=baseQ.text.match(/至少具有\s*(\d+)\s*个/);
            ok=a&&c ? a[1]===c[1] && bf.every(f=> {const id=f.text.slice(0,3),tr=tf.find(x=>x.text.startsWith(id));return tr && f.cells.slice(1).every(v=>compact(tr.text).includes(compact(v)));}):null;
            reason=a&&c?`目标最低 ${a[1]} 个案例；已确认基准最低 ${c[1]} 个案例。逐项核对 Q01–Q03 及证明材料。`:missing;
          }
        }break;
      case 'R10':if(required('technical')){use(/T0[1-5]/,b);ok=['T01','T02','T03','T04','T05'].every(v=>all.includes(v));reason='技术需求编号覆盖检查；内容是否充分仍需人工复核';}break;
      case 'R11':if(required('technical')){use(/120|验收方法/,b,/120/);const days=bf.map(f=>f.text).join(' ').match(/(\d+)\s*(?:个自然日|日内完成项目验收)/);ok=days?new RegExp(`${days[1]}\\s*个自然日`).test(all)&&has(/验收.*技术需求书/):null;reason='交付周期及验收基准核对';}break;
      case 'R12':{tf=take(/综合评分满分|按总分/);const f=tf[0]?.text,m=f?.match(/技术\s*(\d+)\s*分、商务\s*(\d+)\s*分、价格\s*(\d+)\s*分/);ok=m?m.slice(1).reduce((a,b)=>a+Number(b),0)===100&&has(/推荐候选人/):null;reason=m?`${m.slice(1).join(' + ')} = ${m.slice(1).reduce((a,b)=>a+Number(b),0)}；核对推荐规则`:missing;break;}
      case 'R13':{tf=take(/附件清单|投标响应表|合同主要/);const needed=['technical','qualification','contract'];ok=needed.every(k=>by(k))?has(/附件清单/)&&has(/投标响应表/):null;reason=ok===null?'附件实际材料缺失：'+needed.filter(k=>!by(k)).join('、'):'合同、响应表与实际提交附件核对';break;}
      case 'R14':tf=take(/【演示签章：代理机构】/);ok=tf.length>0;reason='仅核对可见演示签章文本，未验证电子签名真伪';break;
      case 'R15':if(required('tender_template')){use(/^[1-6]\.\s/,b);ok=bf.length>=6?bf.slice(0,6).every(f=>tf.some(x=>compact(x.text)===compact(f.text))):null;reason='六节标题与标准模板对照';}break;
      case 'E01':if(required('project')){use(/项目编号|招标人|代理机构/,b);const id=b.fragments.map(f=>f.text).join(' ').match(/DEMO-[A-Z0-9-]+/)?.[0];const agency=b.fragments.find(f=>f.cells?.[0]==='代理机构')?.cells[1];ok=id&&agency?all.includes(id)&&all.includes(agency):null;reason='项目编号、名称及代理机构对照';}break;
      case 'E02':tf=take(/均通过资格|资格／符合性/);ok=tf.length?['A','B','C','D'].every(c=>has(new RegExp(`${c}.*通过|A、B、C、D.*通过`))):null;reason='四家资格与符合性审查结论';break;
      case 'E03':{tf=t.fragments.filter(f=>f.cells?.length===6&&/^[A-Z]$/.test(f.cells[0]));ok=tf.length?tf.every(f=>f.cells.slice(1,4).reduce((a,v)=>a+Number(v),0)===Number(f.cells[4])):null;reason=tf.map(f=>`${f.cells[0]}: ${f.cells.slice(1,4).join('+')}=${f.cells[4]}`).join('；')||missing;break;}
      case 'E04':case 'E05':{
        tf=take(/推荐第.*总分/);bf=t.fragments.filter(f=>f.cells?.length===6&&/^[A-Z]$/.test(f.cells[0]));b=t;
        const rec=[...tf.map(f=>f.text).join(' ').matchAll(/推荐第([一二三四])中标候选人为供应商\s*([A-Z])，总分\s*(\d+)\s*分/g)];
        const pairs=rec.map(m=>({name:m[2],score:+m[3],rank:'一二三四'.indexOf(m[1])+1,row:bf.find(f=>f.cells[0]===m[2])}));
        ok=pairs.length&&pairs.every(p=>p.row)?pairs.every(p=>rule.id==='E04'?p.score===+p.row.cells[4]:p.rank===+p.row.cells[5]):null;
        reason=pairs.map(p=>rule.id==='E04'?`${p.name}：正文 ${p.score}，汇总 ${p.row?.cells[4]??'缺失'}`:`${p.name}：推荐 ${p.rank}，汇总 ${p.row?.cells[5]??'缺失'}`).join('；')||missing;break;}
      case 'E06':tf=take(/演示签字|演示签章/);ok=['甲','乙','丙'].every(v=>all.includes(`【演示签字：成员${v}】`))&&has(/【演示签章：代理机构】/);reason='三名成员及代理机构演示签署位置检查';break;
      case 'E07':tf=take(/^附件\s*[12]\s*(评分明细|资格审查记录)/);ok=tf.length>=2?tf.every(f=>f.text.split('：')[1]?.length>10):null;reason='核对报告内实际附载的评分明细及资格审查记录，单纯目录不计为已提供';break;
      default:reason='自定义规则需要语义模型或人工核验';
    }
    return {...rule,status:ok===null?'无法核验':ok?'通过':'不通过',reason,evidence:[...evidence(t,tf),...(b?evidence(b,bf):[])],suggestion:ok===false?'请代理机构按基准更正后提交新版本复检。':ok===null?'请补齐有效基准、证据或提交人工复核。':'保留证据以备复核。',engine:'deterministic-1.0'};
  });
}

export function compareBids(selection) {
 const bids=selection.selected.filter(d=>d.kind==='bid'),published=selection.selected.find(d=>d.kind==='tender'),pairs=[];
 for(let i=0;i<bids.length;i++)for(let j=i+1;j<bids.length;j++){
  const a=bids[i],b=bids[j]; const matches=[],templates=[];
  for(const f of a.fragments.filter(f=>f.text.length>30))for(const g of b.fragments.filter(g=>g.text.length>30)) {
   const ca=compact(f.text),cb=compact(g.text);const grams=s=>new Set(Array.from({length:Math.max(0,s.length-3)},(_,i)=>s.slice(i,i+4)));const ga=grams(ca),gb=grams(cb);const sim=2*[...ga].filter(x=>gb.has(x)).length/(ga.size+gb.size||1);
   if(sim<0.8)continue;
   const template=/资格及商务|技术条款响应表|实施资源|附件 C|项目风险|培训/.test(f.section)||f.cells||published?.fragments.some(p=>compact(p.text).includes(ca));
   const item={type:ca===cb?'全文重复':'近似段落',similarity:Math.round(sim*100),typo:f.text.includes('申批回朔')&&g.text.includes('申批回朔')?'共同异常用词：申批回朔':null,evidence:[...evidence(a,[f]),...evidence(b,[g])]};
   (template?templates:matches).push(item);
  }
  pairs.push({id:`${a.id}:${b.id}`,left:a.id,right:b.id,leftName:a.name,rightName:b.name,matches,templates,status:'待复核'});
 }
 return pairs;
}

export function responseMatrix(selection) {
 const base=selection.selected.find(d=>d.kind==='tender');if(!base)throw new Error('缺少已确认发布版');
 const requirements=['Q01','Q02','Q03','T01','T02','T03','T04','T05'];
 return selection.selected.filter(d=>d.kind==='bid').map(d=>({id:d.id,name:d.name,supplier:d.fragments.find(f=>/投标人：/.test(f.text))?.text,items:requirements.map(id=>{const re=id==='Q01'?/Z-01.*主体证明.*登记名称/:id==='Q02'?/Z-02[abc].*合同/:id==='Q03'?/Z-03.*独立投标承诺函.*我方/:new RegExp('^'+id+'\\s*\\|');const fs=find(d,re);return {id,status:fs.length?'已定位':'缺证据',baseline:evidence(base,find(base,new RegExp('^'+id+'\\s'))),evidence:evidence(d,fs)};})}));
}


