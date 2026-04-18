/**
 * Paint Approval Tool v6.0 — FINAL FIX
 * ─────────────────────────────────────
 * BUG ĐÃ TÌM RA:
 *   convertToObjectRuntimeIds trả về runtimeId "cha" (assembly),
 *   nhưng getObjects trả về runtimeId "con" (mỗi part/plate/bolt).
 *   → greenSet chứa IDs cha, grayIds chứa IDs con
 *   → filter không loại trừ được → xám đè xanh
 *
 * FIX: Không dùng getObjects để lọc gray nữa.
 *   Bước 1: Reset màu
 *   Bước 2: Tô XANH cho matched GUIDs
 *   Bước 3: Tô XÁM cho TẤT CẢ model (setObjectState undefined = all)
 *           nhưng TRƯỚC khi tô xám, ta cần cách khác
 *
 * CHIẾN LƯỢC MỚI:
 *   Dùng 2 lệnh RIÊNG BIỆT, KHÔNG CHỒNG CHÉO:
 *   1. Reset hết
 *   2. Tô XÁM cho TOÀN BỘ model (1 lệnh duy nhất, undefined selector)
 *   3. Đợi render
 *   4. Tô XANH cho matched GUIDs (sẽ đè lên xám)
 *
 *   Lý do: tô xanh SAU sẽ override xám → xanh thắng!
 * ─────────────────────────────────────
 */

var COLOR_GREEN    = "#00FF00";
var COLOR_GRAY     = "#888888";
var RETRY_MAX      = 7;
var RETRY_DELAY_MS = 2000;
var BATCH_CONVERT  = 500;
var BATCH_COLOR    = 300;
var PAINT_DELAY    = 200;

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

/* ═══ Model groups — CHỈ lấy modelId, không cần runtimeIds chi tiết ═══ */
async function getModelIds() {
  var api = await getAPI();
  for (var a=1; a<=RETRY_MAX; a++) {
    var raw;
    try { raw = await api.viewer.getObjects(); } catch(e) {
      log("getObjects err ("+a+"): "+(e&&e.message?e.message:String(e)),"warn");
      if(a<RETRY_MAX){await sleep(RETRY_DELAY_MS);continue;} throw e;
    }
    if (!Array.isArray(raw) || !raw.length) {
      log("Chưa có object ("+a+"/"+RETRY_MAX+")...","warn");
      if(a<RETRY_MAX){await sleep(RETRY_DELAY_MS);continue;}
      throw new Error("Viewer trống. Đợi model load xong.");
    }

    // Lấy tổng objects để hiển thị stats
    var totalObjects = 0;
    var modelIds = [];
    raw.forEach(function(g) {
      if (!g || !g.modelId) return;
      if (modelIds.indexOf(g.modelId) === -1) modelIds.push(g.modelId);
      // Đếm objects từ mọi format
      if (Array.isArray(g.objects)) totalObjects += g.objects.length;
      else if (Array.isArray(g.objectRuntimeIds)) totalObjects += g.objectRuntimeIds.length;
      else if (Array.isArray(g.ids)) totalObjects += g.ids.length;
    });

    if (!modelIds.length) {
      if(a<RETRY_MAX){await sleep(RETRY_DELAY_MS);continue;}
      throw new Error("Không tìm thấy modelId.");
    }

    modelIds.forEach(function(id){ log("  model "+id.substr(0,12)+"...","info"); });
    return { modelIds: modelIds, totalObjects: totalObjects };
  }
}

/* ═══ Convert GUIDs ═══ */
function flat(v){if(v===null||v===undefined)return[];if(typeof v==="number")return[v];if(Array.isArray(v)){var o=[];v.forEach(function(x){if(typeof x==="number")o.push(x);else if(Array.isArray(x))x.forEach(function(y){if(typeof y==="number")o.push(y);});});return o;}return[];}

async function batchConvert(api,modelId,guids){
  var out=[];
  for(var i=0;i<guids.length;i+=BATCH_CONVERT){
    var c=guids.slice(i,i+BATCH_CONVERT);
    var r;
    try{r=await api.viewer.convertToObjectRuntimeIds(modelId,c);}catch(e){for(var k=0;k<c.length;k++)out.push(null);continue;}
    if(!Array.isArray(r)){for(var k=0;k<c.length;k++)out.push(null);continue;}
    out=out.concat(r);
  }
  return out;
}

/**
 * Trả về Map<modelId, number[]> — danh sách runtimeIds đã match
 * Thử cả UUID và IFC format
 */
async function smartConvert(api, modelIds, guids) {
  var matchByModel = new Map();

  var uuids=[], ifcs=[], others=[];
  guids.forEach(function(g){ var f=detectFmt(g); if(f==="uuid"||f==="uuid-nd") uuids.push(g); else if(f==="ifc") ifcs.push(g); else others.push(g); });
  var u2i = uuids.map(uuid2ifc).filter(Boolean);
  var i2u = ifcs.map(ifc2uuid).filter(Boolean);
  log("  UUID="+uuids.length+", IFC="+ifcs.length+", U→I="+u2i.length,"info");

  for (var mi=0; mi<modelIds.length; mi++) {
    var modelId = modelIds[mi];
    var allIds = [];  // tất cả runtimeIds match được (bao gồm cả sub-parts)

    async function tryList(list, label) {
      if (!list.length) return;
      var conv = await batchConvert(api, modelId, list);
      var hit=0;
      for (var i=0; i<list.length; i++) {
        var ids = flat(conv[i]);
        if (ids.length) { hit++; allIds = allIds.concat(ids); }
      }
      log("  ["+label+"] "+modelId.substr(0,12)+"...: "+hit+"/"+list.length+" GUIDs → "+allIds.length+" runtimeIds", hit>0?"ok":"warn");
    }

    await tryList(uuids,  "UUID");
    await tryList(ifcs,   "IFC ");
    await tryList(u2i,    "U→I ");
    await tryList(i2u,    "I→U ");
    await tryList(others, "RAW ");

    if (allIds.length) {
      // Deduplicate
      var unique = {};
      allIds.forEach(function(id){ unique[id] = true; });
      matchByModel.set(modelId, Object.keys(unique).map(Number));
    }
  }
  return matchByModel;
}

/* ═══ Paint ═══ */
async function paint(api, modelId, ids, color, label) {
  for (var i=0; i<ids.length; i+=BATCH_COLOR) {
    var chunk = ids.slice(i, i+BATCH_COLOR);
    try {
      await api.viewer.setObjectState(
        { modelObjectIds: [{ modelId: modelId, objectRuntimeIds: chunk }] },
        { color: color }
      );
    } catch(e) {
      log("  ✗ "+label+" batch lỗi: "+(e&&e.message?e.message:String(e)),"err");
    }
    if (i+BATCH_COLOR < ids.length) await sleep(PAINT_DELAY);
  }
}

/* ═══════════════════════════════════════
   MAIN v6 — CHIẾN LƯỢC MỚI:
   
   1. Reset
   2. Tô XÁM toàn bộ model (dùng selector undefined = ALL objects)
   3. Đợi 2 giây
   4. Tô XANH cho matched (sẽ ĐÈ lên xám → xanh thắng)
   
   Đây là thứ tự NGƯỢC so với v5.
   Xám trước = nền, Xanh sau = đè lên nền.
   Không cần lọc gray/green vì xanh LUÔN là lệnh cuối.
═══════════════════════════════════════ */
async function applyColors() {
  lockUI(true);
  clearLog();
  setProgress(5);

  try {
    if (!_excelGuids.length) throw new Error("Chưa có GUID.");
    var api = await getAPI();

    // 1. RESET
    log("Bước 1/4: Reset màu...", "info");
    try { await api.viewer.setObjectState(undefined, { color: "reset" }); } catch(e) {}
    await sleep(500);
    setProgress(10);

    // 2. Get model info
    log("Bước 2/4: Lấy model info...", "info");
    var modelInfo = await getModelIds();
    setStat("s-total", fmtN(modelInfo.totalObjects));
    setStat("s-excel", fmtN(_excelGuids.length));
    setProgress(20);

    // 3. Convert GUIDs → runtimeIds
    log("Bước 3/4: Map GUID → runtimeId...", "info");
    var matchMap = await smartConvert(api, modelInfo.modelIds, _excelGuids);
    var greenTotal = 0;
    matchMap.forEach(function(ids){ greenTotal += ids.length; });
    var grayTotal = modelInfo.totalObjects - greenTotal;
    setStat("s-green", fmtN(greenTotal));
    setStat("s-gray", fmtN(grayTotal));

    if (greenTotal === 0) {
      log("✗ Không match object nào!", "err");
      setProgress(0); lockUI(false); return;
    }

    log("Matched: " + fmtN(greenTotal) + " runtimeIds từ " + _excelGuids.length + " GUIDs", "ok");
    setProgress(35);

    // ═══════════════════════════════════
    // 4A. TÔ XÁM TOÀN BỘ MODEL
    // Dùng undefined selector = tất cả objects trong viewer
    // ═══════════════════════════════════
    log("━━━ TÔ XÁM toàn bộ model ━━━", "info");
    try {
      await api.viewer.setObjectState(undefined, { color: COLOR_GRAY });
      log("  ✓ Toàn bộ model → xám", "ok");
    } catch(e) {
      log("  ⚠ setObjectState(undefined) không hỗ trợ, tô theo từng model...", "warn");
      // Fallback: tô từng model
      for (var mi=0; mi<modelInfo.modelIds.length; mi++) {
        try {
          await api.viewer.setObjectState(
            { modelObjectIds: [{ modelId: modelInfo.modelIds[mi] }] },
            { color: COLOR_GRAY }
          );
        } catch(e2) {
          log("  ✗ Fallback xám model lỗi: "+(e2&&e2.message?e2.message:String(e2)), "err");
        }
      }
    }
    setProgress(55);

    // ĐỢI 2 GIÂY để viewer render xám hoàn toàn
    log("Đợi viewer render xám...", "info");
    await sleep(2000);

    // ═══════════════════════════════════
    // 4B. TÔ XANH CHO MATCHED (ĐÈ LÊN XÁM)
    // Đây là lệnh CUỐI CÙNG → luôn thắng
    // ═══════════════════════════════════
    log("━━━ TÔ XANH " + fmtN(greenTotal) + " objects (đè lên xám) ━━━", "info");
    for (var mi=0; mi<modelInfo.modelIds.length; mi++) {
      var modelId = modelInfo.modelIds[mi];
      var greenIds = matchMap.get(modelId);
      if (!greenIds || !greenIds.length) continue;
      await paint(api, modelId, greenIds, COLOR_GREEN, "Xanh");
      log("  ▪ Tô xanh " + fmtN(greenIds.length) + " runtimeIds [" + modelId.substr(0,12) + "...]", "ok");
    }
    setProgress(95);

    // Đợi render xanh
    await sleep(500);

    setProgress(100);
    log("", "info");
    log("✓ HOÀN TẤT!", "ok");
    log("  Xanh: " + fmtN(greenTotal) + " (cấu kiện trong Excel)", "ok");
    log("  Xám:  " + fmtN(grayTotal)  + " (cấu kiện còn lại)", "ok");
    setTimeout(function(){setProgress(0);}, 2000);

  } catch(err) {
    log("✗ " + (err&&err.message?err.message:String(err)), "err");
    setProgress(0);
  } finally {
    lockUI(false);
    if (!_excelGuids.length) document.getElementById("applyBtn").disabled = true;
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
  var c=await api.view.createView({name:name,description:"Paint Approval Tool v6.0 | Le Van Thao"});
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
