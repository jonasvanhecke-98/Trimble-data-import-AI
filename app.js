(() => {
  "use strict";
  const BACKEND = "http://localhost:8000";
  const channel = new BroadcastChannel("trimble-data-import-ai-v1");
  const insideTrimble = window.self !== window.top;
  const $ = id => document.getElementById(id);

  let api=null, project=null, token=null, explorerApi=null, ifcMeta=null, pdfMeta=null, mode="trimble", analysis=null;

  $("excel-link").href = `${BACKEND}/api/audit.xlsx`;

  if (insideTrimble) startExtension();
  else startCompanion();

  async function startExtension(){
    $("extension-ui").hidden=false;
    $("status").textContent="Verbinden met Trimble Connect…";
    try{
      api=await TrimbleConnectWorkspace.connect(window.parent,onEvent,30000);
      project=await api.project.getProject();
      channel.postMessage({type:"project",project});
      const p=await api.extension.requestPermission("accesstoken");
      if(typeof p==="string"&&p.length>20){token=p;channel.postMessage({type:"token",token});}
      $("status").textContent=`Verbonden: ${project?.name||project?.id||"Trimble Connect"}`;
      $("open-companion").onclick=()=>{
        const u=new URL(location.href);u.searchParams.set("companion","1");
        window.open(u.toString(),"trimble-data-import-ai-control","width=1450,height=950");
      };
      channel.addEventListener("message",e=>{
        if(e.data?.type==="companion-ready"){
          if(project) channel.postMessage({type:"project",project});
          if(token) channel.postMessage({type:"token",token});
        }
        if(e.data?.type==="focus"&&e.data.guid) focusGuid(e.data.guid,e.data.modelId);
      });
    }catch(e){$("status").textContent="Trimble verbinding mislukt: "+(e?.message||e);}
  }

  function onEvent(event,args){
    const data=args?.data??args;
    if(event==="extension.accessToken"&&typeof data==="string"&&data.length>20){
      token=data;channel.postMessage({type:"token",token});
    }
  }

  async function focusGuid(guid,preferredModelId){
    try{
      const models=await api.viewer.getModels();
      const search=preferredModelId?models.filter(m=>m.id===preferredModelId):models;
      for(const model of (search.length?search:models)){
        const found=await api.viewer.getObjects({modelObjectIds:[{modelId:model.id}],parameter:{properties:{GlobalId:guid}}});
        if(!found?.length) continue;
        const selector={modelObjectIds:found.map(g=>({modelId:g.modelId,objectIds:(g.objects||[]).map(o=>o.id||o.objectId||o.externalId).filter(Boolean)})).filter(x=>x.objectIds.length)};
        if(selector.modelObjectIds.length){
          await api.viewer.setSelection(selector,"set");
          try{await api.viewer.setCamera(selector,{animationTime:350});}catch{}
          return;
        }
      }
    }catch(e){console.error(e);}
  }

  function startCompanion(){
    $("standalone-ui").hidden=false;$("status").textContent="Controlevenster";
    $("tab-trimble").onclick=()=>setMode("trimble");$("tab-local").onclick=()=>setMode("local");
    $("analyze").onclick=analyze;$("confirm-high").onclick=confirmHigh;$("apply").onclick=applyConfirmed;
    $("local-ifc").onchange=updateAnalyze;$("local-pdf").onchange=updateAnalyze;
    channel.addEventListener("message",e=>{
      if(e.data?.type==="token"){token=e.data.token;maybeExplorer();}
      if(e.data?.type==="project"){project=e.data.project;maybeExplorer();}
    });
    channel.postMessage({type:"companion-ready"});
    updateAnalyze();
  }

  function setMode(m){
    mode=m;$("trimble-pane").hidden=m!=="trimble";$("local-pane").hidden=m!=="local";updateAnalyze();
  }

  async function maybeExplorer(){
    if(!project?.id||!token||explorerApi)return;
    try{
      const iframe=$("trimble-explorer");iframe.src=TrimbleConnectWorkspace.getConnectEmbedUrl();
      explorerApi=await TrimbleConnectWorkspace.connect(iframe,(event,args)=>{
        if(event!=="extension.fileSelected")return;
        const file=args?.data?.file||args?.file;if(!file||file.type!=="FILE")return;
        const n=(file.name||"").toLowerCase();
        if(n.endsWith(".ifc")){ifcMeta=file;$("ifc-name").textContent=file.name;}
        if(n.endsWith(".pdf")){pdfMeta=file;$("pdf-name").textContent=file.name;}
        updateAnalyze();
      },30000);
      await explorerApi.embed.setTokens({accessToken:token});
      await explorerApi.embed.initFileExplorer({projectId:project.id,enableSelect:true,enableUploadFiles:false,enableCreateFolder:false,enableExplorerKebabMenu:false,fileTypeFilter:["ifc","pdf"]});
      $("explorer-status").textContent=`Project: ${project.name||project.id}`;
    }catch(e){$("explorer-status").textContent="Explorer fout: "+(e?.message||e);}
  }

  function updateAnalyze(){
    $("analyze").disabled=mode==="trimble"?!(token&&project?.id&&ifcMeta&&pdfMeta):!($("local-ifc").files?.[0]&&$("local-pdf").files?.[0]);
  }

  async function analyze(){
    setMessage("Analyseren…");
    try{
      let r;
      if(mode==="trimble"){
        r=await fetch(`${BACKEND}/api/analyze/trimble`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({access_token:token,project_id:project.id,ifc_file:ifcMeta,pdf_file:pdfMeta})});
      }else{
        const f=new FormData();f.append("ifc",$("local-ifc").files[0]);f.append("pdf",$("local-pdf").files[0]);
        r=await fetch(`${BACKEND}/api/analyze/local`,{method:"POST",body:f});
      }
      if(!r.ok)throw new Error(await r.text());
      analysis=await r.json();renderResults();setMessage("");
    }catch(e){setMessage(e?.message||String(e));}
  }

  function renderResults(){
    $("results").hidden=false;$("count-elements").textContent=analysis.element_count||0;$("count-assignments").textContent=analysis.assignment_count||0;$("count-proposals").textContent=analysis.rows?.length||0;
    const body=$("results-body");body.innerHTML="";
    (analysis.rows||[]).forEach((row,i)=>{
      const tr=document.createElement("tr"),cc=row.confidence>=95?"green":row.confidence>=75?"amber":"red";
      tr.innerHTML=`<td><b>${esc(row.element_ref||"—")}</b><small>${esc(row.guid)}</small></td><td><input data-v="${i}" value="${attr(row.value||"")}"></td><td><span class="conf ${cc}">${row.confidence}%</span></td><td>${esc(row.method||"")}</td><td><select data-s="${i}"><option value="proposed">Voorstel</option><option value="confirmed">Bevestigd</option><option value="needs_review">Controle nodig</option><option value="rejected">Afgewezen</option></select></td><td><button data-f="${i}">Bekijk</button></td>`;
      body.appendChild(tr);tr.querySelector(`[data-s="${i}"]`).value=row.status||"proposed";
    });
    body.querySelectorAll("[data-v]").forEach(el=>el.oninput=e=>analysis.rows[+e.target.dataset.v].value=e.target.value);
    body.querySelectorAll("[data-s]").forEach(el=>el.onchange=e=>{analysis.rows[+e.target.dataset.s].status=e.target.value;updateApply();});
    body.querySelectorAll("[data-f]").forEach(el=>el.onclick=e=>{const r=analysis.rows[+e.target.dataset.f];channel.postMessage({type:"focus",guid:r.guid,modelId:r.model_id});});
    updateApply();
  }

  function confirmHigh(){(analysis.rows||[]).forEach(r=>{if(r.confidence>=95)r.status="confirmed";});renderResults();}
  function updateApply(){$("apply").disabled=!(analysis?.rows||[]).some(r=>r.status==="confirmed");}
  async function applyConfirmed(){
    if(!token||!project?.id){setMessage("Open dit venster vanuit de Trimble extension.");return;}
    try{
      const r=await fetch(`${BACKEND}/api/apply`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({access_token:token,project_id:project.id,model_name:analysis.model_name,rows:analysis.rows,confirmed_by:"Trimble user"})});
      if(!r.ok)throw new Error(await r.text());const x=await r.json();setMessage(`Klaar: ${x.written||0} verwerkt. Mode: ${x.mode||"onbekend"}`);
    }catch(e){setMessage(e?.message||String(e));}
  }
  function setMessage(t){$("message").textContent=t||"";$("message").hidden=!t;}
  function esc(v){return String(v).replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;");}
  function attr(v){return esc(v).replaceAll('"',"&quot;");}
})();
