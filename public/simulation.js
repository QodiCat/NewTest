// Presentation-only timing. Real parsing and analysis run concurrently.
export async function simulateWorkflow({title,steps,files=[],durationMs=10000},work=async()=>null) {
  if(document.querySelector('#simulationDialog'))throw new Error('已有任务正在执行，请等待完成');
  const dialog=document.createElement('dialog');dialog.id='simulationDialog';dialog.className='simulation-dialog';
  const escape=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  dialog.innerHTML=`<div class="simulation-top"><span class="simulation-spinner" aria-hidden="true"></span><span class="badge">处理中</span></div><h2>${escape(title)}</h2><p class="simulation-description">正在处理所选文档，请稍候。完成后将显示材料及处理结果。</p><div class="simulation-line"><strong id="simulationStage"></strong><span id="simulationPercent">0%</span></div><progress id="simulationProgress" max="100" value="0" aria-label="处理进度"></progress><ol class="simulation-steps">${steps.map((s,i)=>`<li data-step="${i}"><span>${i+1}</span>${escape(s)}</li>`).join('')}</ol><div id="simulationFile" class="muted"></div><p class="muted" id="simulationElapsed">已用时 0 秒</p>`;
  document.body.append(dialog);dialog.addEventListener('cancel',e=>e.preventDefault());dialog.showModal();
  const start=performance.now(),duration=durationMs;
  const tick=()=>{const elapsed=performance.now()-start,p=Math.min(96,Math.floor(elapsed/duration*100)),stage=Math.min(steps.length-1,Math.floor(p/100*steps.length));dialog.querySelector('#simulationProgress').value=p;dialog.querySelector('#simulationPercent').textContent=p+'%';dialog.querySelector('#simulationStage').textContent=steps[stage];dialog.querySelector('#simulationElapsed').textContent=`已用时 ${Math.floor(elapsed/1000)} 秒${elapsed>=duration?' · 正在整理处理结果…':''}`;dialog.querySelectorAll('[data-step]').forEach((li,i)=>{li.className=i<stage?'done':i===stage?'running':'';});if(files.length)dialog.querySelector('#simulationFile').textContent=`${files[Math.min(files.length-1,Math.floor(p/100*files.length))]} · ${Math.min(files.length,Math.floor(p/100*files.length)+1)} / ${files.length}`;};
  tick();const timer=setInterval(tick,100);
  try {
    // Catch immediately so a quick request error cannot become an unhandled rejection.
    const result=Promise.resolve().then(work).then(value=>({value}),error=>({error}));
    const [out]=await Promise.all([result,new Promise(r=>setTimeout(r,duration))]);
    if(out.error)throw out.error;
    clearInterval(timer);dialog.querySelector('#simulationProgress').value=100;dialog.querySelector('#simulationPercent').textContent='100%';dialog.querySelector('#simulationStage').textContent='处理完成';dialog.querySelectorAll('[data-step]').forEach(li=>li.className='done');
    await new Promise(r=>setTimeout(r,250));return out.value;
  } finally {clearInterval(timer);dialog.close();dialog.remove();}
}
