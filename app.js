/**
 * Paint Approval Tool v4.0 - DEBUG MODE
 * ─────────────────────────────────────
 * Mục đích: tìm ra FORMAT MÀU mà Trimble API chịu nhận
 * Tô thử 10 object đầu tiên của mỗi model với NHIỀU FORMAT MÀU khác nhau
 * Nhìn vào model, format nào hiện màu → đó là format đúng
 * ─────────────────────────────────────
 */

var RETRY_MAX      = 7;
var RETRY_DELAY_MS = 2000;

var _api        = null;
var _excelGuids = [];

function log(msg, type) {
  var el = document.getElementById("log");
  if (!el) { console.log(msg); return; }
  var span = document.createElement("span");
  if (type) span.className = type;
  span.textContent = msg + "\n";
  el.appendChild(span);
  el.scrollTop = el.scrollHeight;
  console.log("[" + (type||"log") + "] " + msg);
}
function clearLog(){ var el=document.getElementById("log"); if(el) el.innerHTML=""; }
function setStat(id,v){ var el=document.getElementById(id); if(el) el.textContent=(v!==null&&v!==undefined)?v:"—"; }
function resetStats(){ ["s-total","s-excel","s-green","s-gray"].forEach(function(id){setStat(id,"—");}); }
function setProgress(pct){
  var w=document.getElementById("progWrap"),b=document.getElementById("progBar");
  if(!w||!b) return;
  if(pct<=0){w.classList.remove("on");b.style.width="0%";return;}
  w.classList.add("on"); b.style.width=Math.min(pct,100)+"%";
}
function lockUI(yes){ ["applyBtn","resetBtn","saveBtn"].forEach(function(id){var el=document.getElementById(id);if(el) el.disabled=yes;}); }
function sleep(ms){ return new Promise(function(r){setTimeout(r,ms);}); }
function pad2(n){ return String(n).padStart(2,"0"); }
function fmtNum(n){ return (typeof n==="number")?n.toLocaleString():String(n); }

/* UUID ↔ IFC GUID */
var IFC_B64 = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_$";
function _to64(num, n){var d=[];for(var i=0;i<n;i++){d.push(IFC_B64.charAt(num%64));num=Math.floor(num/64);}return d.reverse().join("");}
function _from64(s){var r=0;for(var i=0;i<s.length;i++){var x=IFC_B64.indexOf(s.charAt(i));if(x<0)return -1;r=r*64+x;}return r;}
function uuidToIfcGuid(uuid){
  if(!uuid) return null;
  var hex=String(uuid).replace(/-/g,"").toLowerCase();
  if(hex.length!==32 || !/^[0-9a-f]{32}$/.test(hex)) return null;
  var num=[parseInt(hex.substr(0,2),16)];
  for(var i=0;i<5;i++) num.push(parseInt(hex.substr(2+i*6,6),16));
  var r=_to64(num[0],2);
  for(var i=1;i<6;i++) r+=_to64(num[i],4);
  return r;
}
function ifcGuidToUuid(ifc){
  if(!ifc||ifc.length!==22) return null;
  var p=[_from64(ifc.substr(0,2))];
  for(var i=0;i<5;i++) p.push(_from64(ifc.substr(2+i*4,4)));
  if(p.some(function(x){return x<0;})) return null;
  var hex=p[0].toString(16).padStart(2,"0");
  for(var i=1;i<6;i++) hex+=p[i].toString(16).padStart(6,"0");
  return hex.substr(0,8)+"-"+hex.substr(8,4)+"-"+hex.substr(12,4)+"-"+hex.substr(16,4)+"-"+hex.substr(20,12);
}
function detectFmt(g){
  if(!g) return "unknown";
  var s=String(g).trim();
  if(s.length===36 && /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(s)) return "uuid";
  if(s.length===32 && /^[0-9a-f]{32}$/i.test(s)) return "uuid-nodash";
  if(s.length===22) return "ifc";
  return "unknown";
}

async function getAPI(){
  if(_api) return _api;
  _api = await TrimbleConnectWorkspace.connect(window.parent, function(ev,data){ console.log("[Trimble]",ev,data); });
  log("Đã kết nối Trimble Workspace API.","ok");
  return _api;
}

function readWorkbook(file){
  return new Promise(function(resolve,reject){
    var r=new FileReader();
    r.onload=function(e){try{resolve(XLSX.read(e.target.result,{type:"array"}));}catch(err){reject(err);}};
    r.onerror=reject; r.readAsArrayBuffer(file);
  });
}
function extractGuids(wb){
  if(!wb||!wb.SheetNames||!wb.SheetNames.length) throw new Error("File Excel không có sheet.");
  var sn=wb.SheetNames[0];
  var rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{defval:""});
  if(!rows.length) throw new Error("Sheet đầu tiên không có dữ liệu.");
  var keys=Object.keys(rows[0]);
  var gk=keys.find(function(k){return k.trim().toUpperCase()==="GUID";});
  if(!gk){ gk=keys[0]; log('⚠ Dùng cột đầu: "'+gk+'"',"warn"); }
  var seen={},out=[];
  rows.forEach(function(r){var g=String(r[gk]||"").trim(); if(g&&!seen[g]){seen[g]=true;out.push(g);}});
  var fmt=out.length?detectFmt(out[0]):"none";
  log('Sheet: "'+sn+'" | '+rows.length+' dòng | '+out.length+' GUID duy nhất',"info");
  log('GUID[0]: '+(out[0]||"N/A")+' ('+fmt+')',"info");
  return out;
}

function parseIds(g){
  var ids={}; if(!g) return [];
  if(Array.isArray(g.objects)) g.objects.forEach(function(o){var v=(o&&o.id!==undefined)?o.id:(o&&o.runtimeId);if(typeof v==="number")ids[v]=true;else if(typeof v==="string"&&v&&!isNaN(+v))ids[+v]=true;});
  if(Array.isArray(g.objectRuntimeIds)) g.objectRuntimeIds.forEach(function(v){if(typeof v==="number")ids[v]=true;else if(typeof v==="string"&&v&&!isNaN(+v))ids[+v]=true;});
  if(Array.isArray(g.ids)) g.ids.forEach(function(v){if(typeof v==="number")ids[v]=true;else if(typeof v==="string"&&v&&!isNaN(+v))ids[+v]=true;});
  return Object.keys(ids).map(Number);
}
async function getModelGroups(){
  var api=await getAPI();
  for(var attempt=1;attempt<=RETRY_MAX;attempt++){
    var raw;
    try{raw=await api.viewer.getObjects();}catch(err){log("getObjects err: "+(err&&err.message?err.message:String(err)),"warn");if(attempt<RETRY_MAX){await sleep(RETRY_DELAY_MS);continue;}throw err;}
    if(!Array.isArray(raw)||!raw.length){log("Chưa có object (attempt "+attempt+")...","warn");if(attempt<RETRY_MAX){await sleep(RETRY_DELAY_MS);continue;}throw new Error("Viewer trống.");}
    var groups=raw.map(function(g){return {modelId:g&&g.modelId,runtimeIds:parseIds(g)};}).filter(function(g){return g.modelId&&g.runtimeIds.length>0;});
    if(!groups.length){if(attempt<RETRY_MAX){await sleep(RETRY_DELAY_MS);continue;}throw new Error("Không parse được runtimeId.");}
    groups.forEach(function(g){log("  ✓ model "+g.modelId.substr(0,10)+"...: "+fmtNum(g.runtimeIds.length)+" objects","info");});
    return groups;
  }
}

function flatten(v){
  if(v===null||v===undefined) return [];
  if(typeof v==="number") return [v];
  if(Array.isArray(v)){var o=[];v.forEach(function(x){if(typeof x==="number")o.push(x);else if(Array.isArray(x))x.forEach(function(y){if(typeof y==="number")o.push(y);});});return o;}
  return [];
}
async function tryConvert(api,modelId,guids){
  var out=[];
  for(var i=0;i<guids.length;i+=500){
    var c=guids.slice(i,i+500);
    var r;
    try{r=await api.viewer.convertToObjectRuntimeIds(modelId,c);}catch(e){for(var k=0;k<c.length;k++)out.push(null);continue;}
    if(!Array.isArray(r)){for(var k=0;k<c.length;k++)out.push(null);continue;}
    out=out.concat(r);
  }
  return out;
}
async function smartConvert(api,modelGroups,guids){
  var m=new Map();
  var uuidList=[],ifcList=[],others=[];
  guids.forEach(function(g){var f=detectFmt(g);if(f==="uuid"||f==="uuid-nodash")uuidList.push(g);else if(f==="ifc")ifcList.push(g);else others.push(g);});
  log("  Format: UUID="+uuidList.length+", IFC="+ifcList.length+", other="+others.length,"info");
  var u2i=uuidList.map(uuidToIfcGuid).filter(function(v){return !!v;});
  var i2u=ifcList.map(ifcGuidToUuid).filter(function(v){return !!v;});

  for(var gi=0;gi<modelGroups.length;gi++){
    var group=modelGroups[gi];
    var mSet=new Set(group.runtimeIds);
    var matches=new Set();

    async function tav(lst,lbl){
      if(!lst.length) return;
      var conv=await tryConvert(api,group.modelId,lst);
      var hit=0;
      for(var i=0;i<lst.length;i++){
        var ids=flatten(conv[i]);
        if(!ids.length) continue;
        ids.forEach(function(id){if(mSet.has(id)){matches.add(id);hit++;}});
      }
      log("  ["+lbl+"] "+group.modelId.substr(0,10)+"...: verified="+hit+"/"+lst.length, hit>0?"ok":"warn");
    }
    if(uuidList.length) await tav(uuidList,"UUID");
    if(ifcList.length)  await tav(ifcList, "IFC ");
    if(u2i.length)      await tav(u2i,     "U→I ");
    if(i2u.length)      await tav(i2u,     "I→U ");
    if(others.length)   await tav(others,  "RAW ");

    if(matches.size>0) m.set(group.modelId, matches);
  }
  return m;
}

/* ═══════════════════════════════════════
   ★★★ DEBUG MODE ★★★
   Thử tô bằng nhiều format màu khác nhau
═══════════════════════════════════════ */

/** Thử tô 10 object đầu bằng các format khác nhau */
async function debugPaintTest(api, modelId, runtimeIds) {
  if (runtimeIds.length < 10) {
    log("  ⚠ Ít object quá để test, dùng " + runtimeIds.length, "warn");
  }

  // Chia ra 6 nhóm nhỏ, mỗi nhóm 1 format khác
  var testSize = Math.min(5, Math.floor(runtimeIds.length / 6));
  if (testSize < 1) testSize = 1;

  var groups = [];
  for (var i = 0; i < 6; i++) {
    groups.push(runtimeIds.slice(i * testSize, (i + 1) * testSize));
  }

  var tests = [
    {
      name: "1. String hex '#FF0000' (đỏ)",
      ids: groups[0],
      state: { color: "#FF0000" }
    },
    {
      name: "2. String hex không # 'FF0000' (đỏ)",
      ids: groups[1],
      state: { color: "FF0000" }
    },
    {
      name: "3. RGB object {r,g,b,a} (xanh lá)",
      ids: groups[2],
      state: { color: { r: 0, g: 255, b: 0, a: 255 } }
    },
    {
      name: "4. RGB object {r,g,b} (vàng)",
      ids: groups[3],
      state: { color: { r: 255, g: 255, b: 0 } }
    },
    {
      name: "5. RGB 0-1 float (tím)",
      ids: groups[4],
      state: { color: { r: 1, g: 0, b: 1, a: 1 } }
    },
    {
      name: "6. Array [r,g,b,a] (cyan)",
      ids: groups[5],
      state: { color: [0, 255, 255, 255] }
    }
  ];

  log("━━━ BẮT ĐẦU TEST 6 FORMAT MÀU ━━━", "info");
  log("Mỗi nhóm " + testSize + " object ở 1 góc khác nhau của model.", "info");
  log("Nhìn model → format nào hiện đúng màu → em biết format đúng!", "info");
  log("", "info");

  for (var t = 0; t < tests.length; t++) {
    var test = tests[t];
    if (!test.ids.length) continue;
    try {
      log("Test " + test.name + "...", "info");
      log("  state: " + JSON.stringify(test.state), "info");
      await api.viewer.setObjectState(
        { modelObjectIds: [{ modelId: modelId, objectRuntimeIds: test.ids }] },
        test.state
      );
      log("  → Lệnh đã gửi OK. IDs tô: " + test.ids.slice(0,3).join(",") + "...", "ok");
    } catch (err) {
      log("  ✗ LỖI: " + (err && err.message ? err.message : String(err)), "err");
    }
    await sleep(500);
  }

  log("", "info");
  log("━━━ HOÀN TẤT TEST ━━━", "ok");
  log("👉 Anh NHÌN VÀO MODEL xem có nhóm nào hiện màu không!", "ok");
  log("   - Nếu thấy ĐỎ  → format 1 hoặc 2 đúng", "info");
  log("   - Nếu thấy XANH LÁ → format 3 đúng", "info");
  log("   - Nếu thấy VÀNG → format 4 đúng", "info");
  log("   - Nếu thấy TÍM  → format 5 đúng", "info");
  log("   - Nếu thấy CYAN → format 6 đúng", "info");
  log("   - Nếu KHÔNG thấy gì → API setObjectState có vấn đề", "warn");
  log("", "info");
  log("📸 Anh chụp màn hình + log gửi lại em!", "info");
}

async function applyColors() {
  lockUI(true);
  clearLog();
  setProgress(5);
  try {
    if (!_excelGuids.length) throw new Error("Chưa có GUID.");
    var api = await getAPI();

    log("🔬 CHẾ ĐỘ DEBUG v4.0", "warn");
    log("", "info");

    setProgress(15);
    var modelGroups = await getModelGroups();
    var totalObjects = modelGroups.reduce(function(s,g){return s+g.runtimeIds.length;},0);
    setStat("s-total", fmtNum(totalObjects));
    setStat("s-excel", fmtNum(_excelGuids.length));

    setProgress(30);
    log("Map GUID → runtimeId...", "info");
    var matchMap = await smartConvert(api, modelGroups, _excelGuids);

    var greenTotal = 0;
    matchMap.forEach(function(s){greenTotal+=s.size;});
    setStat("s-green", fmtNum(greenTotal));
    setStat("s-gray",  fmtNum(totalObjects - greenTotal));

    if (greenTotal === 0) {
      log("✗ Không match object nào. Không thể test.", "err");
      setProgress(0); lockUI(false); return;
    }

    setProgress(60);

    // Lấy model đầu tiên có match, test trên runtime IDs đã verified
    var targetModelId = null;
    var targetIds = null;
    matchMap.forEach(function(set, modelId) {
      if (!targetModelId) {
        targetModelId = modelId;
        targetIds = Array.from(set);
      }
    });

    log("Test trên model: " + targetModelId, "info");
    log("Số object có thể test: " + targetIds.length, "info");
    log("", "info");

    // Reset trước
    try {
      await api.viewer.setObjectState(undefined, { color: "reset" });
      await sleep(500);
    } catch(e) { log("reset fallback: "+(e&&e.message?e.message:String(e)), "warn"); }

    // Test
    await debugPaintTest(api, targetModelId, targetIds);

    setProgress(100);
    setTimeout(function(){setProgress(0);}, 1500);

  } catch (err) {
    log("✗ " + (err && err.message ? err.message : String(err)), "err");
    setProgress(0);
  } finally {
    lockUI(false);
    if (!_excelGuids.length) document.getElementById("applyBtn").disabled = true;
  }
}

async function resetViewer(){
  lockUI(true); clearLog(); setProgress(10);
  try{
    var api=await getAPI();
    try{await api.viewer.setObjectState(undefined,{color:"reset",visible:"reset"});}catch(e){log("reset fallback: "+(e&&e.message?e.message:String(e)),"warn");}
    await api.viewer.reset();
    resetStats();
    setProgress(100); log("✓ Đã reset viewer.","ok");
    setTimeout(function(){setProgress(0);},1000);
  }catch(err){log("✗ "+(err&&err.message?err.message:String(err)),"err");setProgress(0);}
  finally{lockUI(false); if(!_excelGuids.length) document.getElementById("applyBtn").disabled=true;}
}
async function saveView(){
  try{
    var api=await getAPI();
    var inp=document.getElementById("viewName");
    var name=inp?inp.value.trim():"";
    if(!name){var now=new Date();name="Approval "+now.getFullYear()+"-"+pad2(now.getMonth()+1)+"-"+pad2(now.getDate())+" "+pad2(now.getHours())+":"+pad2(now.getMinutes());if(inp)inp.value=name;}
    var c=await api.view.createView({name:name,description:"Paint Approval Tool v4.0 DEBUG | Le Van Thao"});
    if(!c||!c.id) throw new Error("No view ID.");
    await api.view.updateView({id:c.id});
    await api.view.selectView(c.id);
    log('✓ Đã lưu view: "'+name+'"',"ok");
  }catch(err){log("✗ Save view: "+(err&&err.message?err.message:String(err)),"err");}
}

document.getElementById("fileInput").addEventListener("change", async function(e){
  var file=e.target.files&&e.target.files[0];
  if(!file) return;
  document.getElementById("fileName").textContent=file.name;
  clearLog();
  log('Đang đọc "'+file.name+'"...',"info");
  try{
    var wb=await readWorkbook(file);
    _excelGuids=extractGuids(wb);
    setStat("s-excel",fmtNum(_excelGuids.length));
    if(_excelGuids.length>0){document.getElementById("applyBtn").disabled=false;log('✓ Sẵn sàng! Nhấn "Áp màu" để chạy DEBUG TEST.',"ok");}
    else log("⚠ Không tìm thấy GUID nào.","warn");
  }catch(err){log("✗ "+(err&&err.message?err.message:String(err)),"err");_excelGuids=[];document.getElementById("applyBtn").disabled=true;}
});
(function(){
  var z=document.getElementById("uploadZone");
  z.addEventListener("dragover",function(e){e.preventDefault();z.classList.add("over");});
  z.addEventListener("dragleave",function(){z.classList.remove("over");});
  z.addEventListener("drop",function(e){e.preventDefault();z.classList.remove("over");var f=e.dataTransfer&&e.dataTransfer.files&&e.dataTransfer.files[0];if(f){document.getElementById("fileInput").files=e.dataTransfer.files;document.getElementById("fileInput").dispatchEvent(new Event("change"));}});
})();
document.getElementById("applyBtn").addEventListener("click",applyColors);
document.getElementById("resetBtn").addEventListener("click",resetViewer);
document.getElementById("saveBtn").addEventListener("click",saveView);
