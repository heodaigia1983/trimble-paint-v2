/**
 * Paint Approval Tool v8.0
 * ─────────────────────────────────────
 * v7 đã work: ẩn hết → hiện+xanh → hiện lại
 * v8 thêm: tô xám phần còn lại
 * 
 * CHIẾN LƯỢC:
 *   1. Reset
 *   2. Ẩn (hide) TẤT CẢ objects
 *   3. Hiện + tô XANH cho matched GUIDs
 *   4. Hiện + tô XÁM cho phần còn lại
 *      → Dùng runtimeIds từ getObjects (sub-level)
 *      → Loại trừ tất cả runtimeIds ĐÃ tô xanh
 *      → Không dùng selector undefined (gây đè)
 * ─────────────────────────────────────
 */

var COLOR_GREEN    = "#00FF00";
var COLOR_GRAY     = "#888888";
var RETRY_MAX      = 7;
var RETRY_DELAY_MS = 2000;
var BATCH_CONVERT  = 500;
var BATCH_COLOR    = 300;
var PAINT_DELAY    = 150;

var _api = null;
var _excelGuids = [];

/* ═══ UI ═══ */
function log(m,t){var e=document.getElementById("log");if(!e){console.log(m);return;}var s=document.createElement("span");if(t)s.className=t;s.textContent=m+"\n";e.appendChild(s);e.scrollTop=e.scrollHeight;console.log("["+(t||"log")+"] "+m);}
function clearLog(){var e=document.getElementById("log");if(e)e.innerHTML="";}
function setStat(id,v){var e=document.getElementById(id);if(e)e.textContent=(v!==null&&v!==undefined)?v:"—";}
function resetStats(){["s-total","s-excel","s-green","s-gray"].forEach(function(id){setStat(id,"—");});}
function setProgress(p){var w=document.getElementById("progWrap"),b=document.getElementById("progBar");if(!w||!b)return;if(p<=0){w.classList.remove("on");b.style.width="0%";return;}w.classList.add("on");b.style.width=Math.min(p,100)+"%";}
function lockUI(y){["applyBtn","resetBtn","saveBtn"].forEach(function(id){var e=document.getElementById(id);if(e)e.disabled=y;});}
function sleep(ms){return new Promise(function(r){setTimeout(r,ms);});}
function pad2(n){return String(n).padStart(2,"0");}
function fmtN(n){return(typeof n==="number")?n.toLocaleString():String(n);}

/* ═══ UUID ↔ IFC GUID ═══ */
var B64="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_$";
function to64(n,d){var r=[];for(var i=0;i<d;i++){r.push(B64.charAt(n%64));n=Math.floor(n/64);}return r.reverse().join("");}
function from64(s){var r=0;for(var i=0;i<s.length;i++){var x=B64.indexOf(s.charAt(i));if(x<0)return-1;r=r*64+x;}return r;}
function uuid2ifc(u){if(!u)return null;var h=String(u).replace(/-/g,"").toLowerCase();if(h.length!==32||!/^[0-9a-f]{32}$/.test(h))return null;var n=[parseInt(h.substr(0,2),16)];for(var i=0;i<5;i++)n.push(parseInt(h.substr(2+i*6,6),16));var r=to64(n[0],2);for(var i=1;i<6;i++)r+=to64(n[i],4);return r;}
function ifc2uuid(c){if(!c||c.length!==22)return null;var p=[from64(c.substr(0,2))];for(var i=0;i<5;i++)p.push(from64(c.substr(2+i*4,4)));if(p.some(function(x){return x<0;}))return null;var h=p[0].toString(16).padStart(2,"0");for(var i=1;i<6;i++)h+=p[i].toString(16).padStart(6,"0");return h.substr(0,8)+"-"+h.substr(8,4)+"-"+h.substr(12,4)+"-"+h.substr(16,4)+"-"+h.substr(20,12);}
function detectFmt(g){if(!g)return"x";var s=String(g).trim();if(s.length===36&&/^[0-9a-f]{8}-/i.test(s))return"uuid";if(s.length===32&&/^[0-9a-f]{32}$/i.test(s))return"uuid-nd";if(s.length===22)return"ifc";return"x";}

/* ═══ API ═══ */
async function getAPI(){if(_api)return _api;_api=await TrimbleConnectWorkspace.connect(window.parent,function(e,d){console.log("[T]",e,d);});log("Đã kết nối Trimble API.","ok");return _api;}

/* ═══ Excel ═══ */
function readWB(f){return new Promise(function(ok,no){var r=new FileReader();r.onload=function(e){try{ok(XLSX.read(e.target.result,{type:"array"}));}catch(err){no(err);}};r.onerror=no;r.readAsArrayBuffer(f);});}
function extractGuids(wb){
  if(!wb||!wb.SheetNames||!wb.SheetNames.length)throw new Error("Excel không có sheet.");
  var sn=wb.SheetNames[0];
  var rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{defval:""});
  if(!rows.length)throw new Error("Sheet trống.");
  var keys=Object.keys(rows[0]);
  var gk=keys.find(function(k){return k.trim().toUpperCase()==="GUID";});
  if(!gk){gk=keys[0];log('⚠ Dùng cột đầu: "'+gk+'"',"warn");}
  var seen={},out=[];
  rows.forEach(function(r){var g=String(r[gk]||"").trim();if(g&&!seen[g]){seen[g]=true;out.push(g);}});
  log('Sheet "'+sn+'": '+rows.length+' dòng, '+out.length+' GUID',"info");
  if(out.length) log('GUID[0]: '+out[0]+' ('+detectFmt(out[0])+')',"info");
  return out;
}

/* ═══ Model groups — lấy CẢ runtimeIds sub-level ═══ */
function parseIds(g){
  var ids={};if(!g)return[];
  if(Array.isArray(g.objects))g.objects.forEach(function(o){var v=(o&&o.id!==undefined)?o.id:(o&&o.runtimeId);if(typeof v==="number")ids[v]=1;else if(typeof v==="string"&&v&&!isNaN(+v))ids[+v]=1;});
  if(Array.isArray(g.objectRuntimeIds))g.objectRuntimeIds.forEach(function(v){if(typeof v==="number")ids[v]=1;else if(typeof v==="string"&&v&&!isNaN(+v))ids[+v]=1;});
  if(Array.isArray(g.ids))g.ids.forEach(function(v){if(typeof v==="number")ids[v]=1;else if(typeof v==="string"&&v&&!isNaN(+v))ids[+v]=1;});
  return Object.keys(ids).map(Number);
}

async function getModelGroups(){
  var api=await getAPI();
  for(var a=1;a<=RETRY_MAX;a++){
    var raw;
    try{raw=await api.viewer.getObjects();}catch(e){
      log("getObjects err ("+a+"): "+(e&&e.message?e.message:String(e)),"warn");
      if(a<RETRY_MAX){await sleep(RETRY_DELAY_MS);continue;}throw e;
    }
    if(!Array.isArray(raw)||!raw.length){
      log("Chưa có object ("+a+"/"+RETRY_MAX+")...","warn");
      if(a<RETRY_MAX){await sleep(RETRY_DELAY_MS);continue;}throw new Error("Viewer trống.");
    }
    var groups=[];
    raw.forEach(function(g){
      if(!g||!g.modelId)return;
      var existing=groups.find(function(x){return x.modelId===g.modelId;});
      var ids=parseIds(g);
      if(existing){existing.runtimeIds=existing.runtimeIds.concat(ids);}
      else{groups.push({modelId:g.modelId,runtimeIds:ids});}
    });
    groups=groups.filter(function(g){return g.runtimeIds.length>0;});
    if(!groups.length){if(a<RETRY_MAX){await sleep(RETRY_DELAY_MS);continue;}throw new Error("Không parse được.");}
    groups.forEach(function(g){log("  model "+g.modelId.substr(0,12)+"...: "+fmtN(g.runtimeIds.length)+" sub-objects","info");});
    return groups;
  }
}

/* ═══ Convert GUIDs → runtimeIds (assembly level) ═══ */
function flat(v){if(v===null||v===undefined)return[];if(typeof v==="number")return[v];if(Array.isArray(v)){var o=[];v.forEach(function(x){if(typeof x==="number")o.push(x);else if(Array.isArray(x))x.forEach(function(y){if(typeof y==="number")o.push(y);});});return o;}return[];}
async function batchConvert(api,mid,guids){var out=[];for(var i=0;i<guids.length;i+=BATCH_CONVERT){var c=guids.slice(i,i+BATCH_CONVERT);var r;try{r=await api.viewer.convertToObjectRuntimeIds(mid,c);}catch(e){for(var k=0;k<c.length;k++)out.push(null);continue;}if(!Array.isArray(r)){for(var k=0;k<c.length;k++)out.push(null);continue;}out=out.concat(r);}return out;}

async function smartConvert(api,modelIds,guids){
  var m=new Map();
  var uuids=[],ifcs=[],others=[];
  guids.forEach(function(g){var f=detectFmt(g);if(f==="uuid"||f==="uuid-nd")uuids.push(g);else if(f==="ifc")ifcs.push(g);else others.push(g);});
  var u2i=uuids.map(uuid2ifc).filter(Boolean);
  var i2u=ifcs.map(ifc2uuid).filter(Boolean);
  log("  UUID="+uuids.length+", U→I="+u2i.length,"info");
  for(var mi=0;mi<modelIds.length;mi++){
    var mid=modelIds[mi];var all=[];
    async function tryL(list,label){
      if(!list.length)return;var conv=await batchConvert(api,mid,list);var hit=0;
      for(var i=0;i<list.length;i++){var ids=flat(conv[i]);if(ids.length){hit++;all=all.concat(ids);}}
      if(hit>0) log("  ["+label+"] "+hit+"/"+list.length+" → "+all.length+" runtimeIds","ok");
    }
    await tryL(uuids,"UUID");await tryL(ifcs,"IFC ");await tryL(u2i,"U→I ");await tryL(i2u,"I→U ");await tryL(others,"RAW ");
    if(all.length){var u={};all.forEach(function(id){u[id]=1;});m.set(mid,Object.keys(u).map(Number));}
  }
  return m;
}

/* ═══ Batch paint with delay ═══ */
async function paintBatch(api,mid,ids,state,label){
  for(var i=0;i<ids.length;i+=BATCH_COLOR){
    var chunk=ids.slice(i,i+BATCH_COLOR);
    try{await api.viewer.setObjectState({modelObjectIds:[{modelId:mid,objectRuntimeIds:chunk}]},state);}
    catch(e){log("  ✗ "+label+" lỗi: "+(e&&e.message?e.message:String(e)),"err");}
    if(i+BATCH_COLOR<ids.length)await sleep(PAINT_DELAY);
  }
}

/* ═══════════════════════════════════════
   MAIN v8
   
   1. Reset
   2. Ẩn TẤT CẢ (setObjectState undefined visible:false)
   3. Hiện + XANH cho matched GUIDs (assembly-level IDs)
   4. Hiện + XÁM cho phần còn lại
      → Dùng sub-object IDs từ getObjects
      → Loại trừ IDs đã tô xanh
      → Paint TỪNG BATCH, không dùng undefined selector
   5. Tô XANH lại lần 2 (đảm bảo)
═══════════════════════════════════════ */
async function applyColors(){
  lockUI(true);clearLog();setProgress(5);
  try{
    if(!_excelGuids.length)throw new Error("Chưa có GUID.");
    var api=await getAPI();

    // RESET
    log("Bước 1: Reset...","info");
    try{await api.viewer.setObjectState(undefined,{color:"reset",visible:"reset"});}catch(e){}
    await sleep(500);
    setProgress(8);

    // GET MODEL GROUPS (sub-object level)
    log("Bước 2: Lấy model info...","info");
    var modelGroups=await getModelGroups();
    var totalObjects=modelGroups.reduce(function(s,g){return s+g.runtimeIds.length;},0);
    setStat("s-total",fmtN(totalObjects));
    setStat("s-excel",fmtN(_excelGuids.length));
    setProgress(18);

    // CONVERT GUIDs → runtimeIds (assembly level)
    log("Bước 3: Map GUID → runtimeId...","info");
    var modelIds=modelGroups.map(function(g){return g.modelId;});
    var matchMap=await smartConvert(api,modelIds,_excelGuids);
    var greenTotal=0;
    matchMap.forEach(function(ids){greenTotal+=ids.length;});
    setStat("s-green",fmtN(greenTotal));
    setStat("s-gray",fmtN(totalObjects-greenTotal));
    if(greenTotal===0){log("✗ Không match!","err");setProgress(0);lockUI(false);return;}
    setProgress(30);

    // ẨN TẤT CẢ
    log("Bước 4: Ẩn toàn bộ model...","info");
    try{await api.viewer.setObjectState(undefined,{visible:false});}catch(e){log("  ⚠ ẩn fallback","warn");}
    await sleep(800);
    setProgress(38);

    // HIỆN + XANH cho matched
    log("Bước 5: Hiện + tô XANH "+fmtN(greenTotal)+" objects...","info");
    for(var gi=0;gi<modelGroups.length;gi++){
      var g=modelGroups[gi];
      var greenIds=matchMap.get(g.modelId);
      if(!greenIds||!greenIds.length)continue;
      await paintBatch(api,g.modelId,greenIds,{visible:true,color:COLOR_GREEN},"Xanh");
      log("  ▪ Xanh "+fmtN(greenIds.length)+" ["+g.modelId.substr(0,10)+"...]","ok");
    }
    setProgress(55);
    await sleep(500);

    // HIỆN + XÁM cho phần còn lại
    // Dùng sub-object runtimeIds từ getObjects, loại trừ greenIds
    log("Bước 6: Hiện + tô XÁM phần còn lại...","info");
    var grayCount=0;
    for(var gi=0;gi<modelGroups.length;gi++){
      var g=modelGroups[gi];
      var greenIds=matchMap.get(g.modelId)||[];
      var greenSet={};
      greenIds.forEach(function(id){greenSet[id]=1;});
      
      // Sub-object IDs mà KHÔNG nằm trong greenIds
      var grayIds=g.runtimeIds.filter(function(id){return !greenSet[id];});
      if(!grayIds.length)continue;
      
      await paintBatch(api,g.modelId,grayIds,{visible:true,color:COLOR_GRAY},"Xám");
      grayCount+=grayIds.length;
      log("  ▫ Xám "+fmtN(grayIds.length)+" ["+g.modelId.substr(0,10)+"...]","ok");
    }
    setProgress(85);
    await sleep(500);

    // TÔ XANH LẠI LẦN 2 (đảm bảo không bị đè)
    log("Bước 7: Confirm xanh...","info");
    for(var gi=0;gi<modelGroups.length;gi++){
      var g=modelGroups[gi];
      var greenIds=matchMap.get(g.modelId);
      if(!greenIds||!greenIds.length)continue;
      await paintBatch(api,g.modelId,greenIds,{color:COLOR_GREEN},"Xanh2");
    }
    setProgress(100);

    log("","info");
    log("✓ HOÀN TẤT!","ok");
    log("  Xanh: "+fmtN(greenTotal)+" cấu kiện (trong Excel)","ok");
    log("  Xám:  "+fmtN(grayCount)+" cấu kiện (còn lại)","ok");
    setTimeout(function(){setProgress(0);},2000);

  }catch(err){
    log("✗ "+(err&&err.message?err.message:String(err)),"err");setProgress(0);
  }finally{
    lockUI(false);if(!_excelGuids.length)document.getElementById("applyBtn").disabled=true;
  }
}

/* ═══ Reset ═══ */
async function resetViewer(){
  lockUI(true);clearLog();setProgress(10);
  try{var api=await getAPI();try{await api.viewer.setObjectState(undefined,{color:"reset",visible:"reset"});}catch(e){}await api.viewer.reset();resetStats();setProgress(100);log("✓ Reset OK.","ok");setTimeout(function(){setProgress(0);},1000);}
  catch(e){log("✗ "+(e&&e.message?e.message:String(e)),"err");setProgress(0);}
  finally{lockUI(false);if(!_excelGuids.length)document.getElementById("applyBtn").disabled=true;}
}

/* ═══ Save View ═══ */
async function saveView(){
  try{var api=await getAPI();var inp=document.getElementById("viewName");var name=inp?inp.value.trim():"";
  if(!name){var n=new Date();name="Approval "+n.getFullYear()+"-"+pad2(n.getMonth()+1)+"-"+pad2(n.getDate())+" "+pad2(n.getHours())+":"+pad2(n.getMinutes());if(inp)inp.value=name;}
  var c=await api.view.createView({name:name,description:"Paint Approval Tool v8.0 | Le Van Thao"});
  if(!c||!c.id)throw new Error("No view ID.");await api.view.updateView({id:c.id});await api.view.selectView(c.id);
  log('✓ View: "'+name+'"',"ok");}catch(e){log("✗ Save: "+(e&&e.message?e.message:String(e)),"err");}
}

/* ═══ Events ═══ */
document.getElementById("fileInput").addEventListener("change",async function(e){
  var f=e.target.files&&e.target.files[0];if(!f)return;
  document.getElementById("fileName").textContent=f.name;clearLog();
  log('Đang đọc "'+f.name+'"...',"info");
  try{var wb=await readWB(f);_excelGuids=extractGuids(wb);setStat("s-excel",fmtN(_excelGuids.length));
  if(_excelGuids.length>0){document.getElementById("applyBtn").disabled=false;log('✓ Nhấn "Áp màu" để bắt đầu.',"ok");}
  else log("⚠ Không thấy GUID.","warn");}
  catch(e){log("✗ "+(e&&e.message?e.message:String(e)),"err");_excelGuids=[];document.getElementById("applyBtn").disabled=true;}
});
(function(){var z=document.getElementById("uploadZone");z.addEventListener("dragover",function(e){e.preventDefault();z.classList.add("over");});z.addEventListener("dragleave",function(){z.classList.remove("over");});z.addEventListener("drop",function(e){e.preventDefault();z.classList.remove("over");var f=e.dataTransfer&&e.dataTransfer.files&&e.dataTransfer.files[0];if(f){document.getElementById("fileInput").files=e.dataTransfer.files;document.getElementById("fileInput").dispatchEvent(new Event("change"));}});})();
document.getElementById("applyBtn").addEventListener("click",applyColors);
document.getElementById("resetBtn").addEventListener("click",resetViewer);
document.getElementById("saveBtn").addEventListener("click",saveView);
