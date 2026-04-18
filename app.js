/**
 * Paint Approval Tool v7.0
 * ─────────────────────────────────────
 * BUG ROOT CAUSE:
 *   setObjectState(undefined, {color: gray}) tô xám TẤT CẢ sub-objects
 *   convertToObjectRuntimeIds trả runtimeId cấp assembly (cha)
 *   → tô xanh cho cha nhưng con vẫn xám → mắt thấy xám
 *
 * FIX v7:
 *   KHÔNG tô xám bằng setObjectState(undefined) nữa.
 *   Bước 1: Reset
 *   Bước 2: Tô XANH cho matched GUIDs (chỉ lệnh duy nhất)
 *   Bước 3: Làm mờ phần còn lại bằng cách ẩn (visible=false)
 *           hoặc set opacity thấp
 *
 *   Nếu hide không hỗ trợ → chỉ tô xanh, không tô xám gì cả.
 *   User thấy: model màu gốc + 299 cấu kiện xanh nổi bật.
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
var _lastMatchMap = null;  // lưu lại để dùng cho gray sau
var _lastModelIds = null;

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

/* ═══ Model ═══ */
async function getModelIds(){
  var api=await getAPI();
  for(var a=1;a<=RETRY_MAX;a++){
    var raw;
    try{raw=await api.viewer.getObjects();}catch(e){log("getObjects err ("+a+"): "+(e&&e.message?e.message:String(e)),"warn");if(a<RETRY_MAX){await sleep(RETRY_DELAY_MS);continue;}throw e;}
    if(!Array.isArray(raw)||!raw.length){log("Chưa có object ("+a+"/"+RETRY_MAX+")...","warn");if(a<RETRY_MAX){await sleep(RETRY_DELAY_MS);continue;}throw new Error("Viewer trống.");}
    var total=0,mids=[];
    raw.forEach(function(g){if(!g||!g.modelId)return;if(mids.indexOf(g.modelId)===-1)mids.push(g.modelId);if(Array.isArray(g.objects))total+=g.objects.length;else if(Array.isArray(g.objectRuntimeIds))total+=g.objectRuntimeIds.length;else if(Array.isArray(g.ids))total+=g.ids.length;});
    if(!mids.length){if(a<RETRY_MAX){await sleep(RETRY_DELAY_MS);continue;}throw new Error("Không tìm thấy modelId.");}
    return{modelIds:mids,totalObjects:total};
  }
}

/* ═══ Convert ═══ */
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
      if(!list.length)return;
      var conv=await batchConvert(api,mid,list);var hit=0;
      for(var i=0;i<list.length;i++){var ids=flat(conv[i]);if(ids.length){hit++;all=all.concat(ids);}}
      log("  ["+label+"] "+hit+"/"+list.length+" GUIDs → "+all.length+" total runtimeIds",hit>0?"ok":"warn");
    }
    await tryL(uuids,"UUID"); await tryL(ifcs,"IFC "); await tryL(u2i,"U→I "); await tryL(i2u,"I→U "); await tryL(others,"RAW ");
    if(all.length){var u={};all.forEach(function(id){u[id]=1;});m.set(mid,Object.keys(u).map(Number));}
  }
  return m;
}

/* ═══ Paint ═══ */
async function paint(api,mid,ids,color,label){
  for(var i=0;i<ids.length;i+=BATCH_COLOR){
    var chunk=ids.slice(i,i+BATCH_COLOR);
    try{await api.viewer.setObjectState({modelObjectIds:[{modelId:mid,objectRuntimeIds:chunk}]},{color:color});}
    catch(e){log("  ✗ "+label+" batch lỗi: "+(e&&e.message?e.message:String(e)),"err");}
    if(i+BATCH_COLOR<ids.length)await sleep(PAINT_DELAY);
  }
}

/* ═══════════════════════════════════════
   MAIN v7: CHỈ TÔ XANH, KHÔNG TÔ XÁM
   
   Lý do: setObjectState(undefined, gray) override cả sub-objects
   mà convertToObjectRuntimeIds trả IDs cấp cha
   → con bị xám, cha xanh nhưng con che cha → chỉ thấy xám
   
   Chiến lược:
   1. Reset
   2. Ẩn (hide) TẤT CẢ objects
   3. Hiện (show) + tô XANH cho matched
   4. Hiện (show) phần còn lại nhưng KHÔNG TÔ MÀU
      → phần còn lại giữ màu gốc nhưng set opacity thấp
      
   Nếu hide/show không work → chỉ tô xanh thôi.
═══════════════════════════════════════ */
async function applyColors(){
  lockUI(true);clearLog();setProgress(5);
  try{
    if(!_excelGuids.length) throw new Error("Chưa có GUID.");
    var api=await getAPI();

    // Reset
    log("Reset màu...","info");
    try{await api.viewer.setObjectState(undefined,{color:"reset"});}catch(e){}
    await sleep(500);
    setProgress(10);

    var mi=await getModelIds();
    _lastModelIds=mi.modelIds;
    setStat("s-total",fmtN(mi.totalObjects));
    setStat("s-excel",fmtN(_excelGuids.length));
    setProgress(25);

    log("Map GUID → runtimeId...","info");
    var matchMap=await smartConvert(api,mi.modelIds,_excelGuids);
    _lastMatchMap=matchMap;
    var greenTotal=0;
    matchMap.forEach(function(ids){greenTotal+=ids.length;});
    setStat("s-green",fmtN(greenTotal));
    setStat("s-gray",fmtN(mi.totalObjects-greenTotal));

    if(greenTotal===0){log("✗ Không match!","err");setProgress(0);lockUI(false);return;}

    setProgress(40);

    // ═══════════════════════════════════
    // CHIẾN LƯỢC: Ẩn hết → hiện xanh → hiện còn lại mờ
    // ═══════════════════════════════════

    // Bước 1: Ẩn TẤT CẢ
    log("━━━ Ẩn toàn bộ model ━━━","info");
    try{
      await api.viewer.setObjectState(undefined,{visible:false});
      log("  ✓ Ẩn toàn bộ","ok");
    }catch(e){
      log("  ⚠ Ẩn không hỗ trợ: "+(e&&e.message?e.message:String(e)),"warn");
    }
    await sleep(800);
    setProgress(55);

    // Bước 2: Hiện + tô XANH cho matched
    log("━━━ Hiện + tô XANH "+fmtN(greenTotal)+" objects ━━━","info");
    for(var i=0;i<mi.modelIds.length;i++){
      var mid=mi.modelIds[i];
      var greenIds=matchMap.get(mid);
      if(!greenIds||!greenIds.length) continue;

      // Hiện
      for(var j=0;j<greenIds.length;j+=BATCH_COLOR){
        var chunk=greenIds.slice(j,j+BATCH_COLOR);
        try{
          await api.viewer.setObjectState(
            {modelObjectIds:[{modelId:mid,objectRuntimeIds:chunk}]},
            {visible:true, color:COLOR_GREEN}
          );
        }catch(e){
          log("  ✗ xanh batch lỗi: "+(e&&e.message?e.message:String(e)),"err");
        }
        if(j+BATCH_COLOR<greenIds.length) await sleep(PAINT_DELAY);
      }
      log("  ▪ Xanh "+fmtN(greenIds.length)+" objects ["+mid.substr(0,12)+"...]","ok");
    }
    setProgress(75);
    await sleep(500);

    // Bước 3: Hiện phần còn lại + tô XÁM
    // Dùng visible:true + color:gray cho TOÀN BỘ
    // Nhưng vì xanh đã set → Trimble sẽ GIỮ xanh cho những object đã set
    // ... KHÔNG, Trimble sẽ override lại hết.
    
    // Vậy cách đúng: HIỆN phần còn lại mà KHÔNG đổi màu xanh
    // → Cần set visible:true cho toàn bộ, nhưng chỉ set color:gray cho phần CHƯA set
    // → Trimble không hỗ trợ "chỉ set visible mà không đổi color"
    
    // CÁCH ĐƠN GIẢN NHẤT: hiện tất cả lại, chỉ set visible
    log("━━━ Hiện phần còn lại (giữ xanh) ━━━","info");
    try{
      await api.viewer.setObjectState(undefined,{visible:true});
      log("  ✓ Hiện toàn bộ","ok");
    }catch(e){
      log("  ⚠ "+(e&&e.message?e.message:String(e)),"warn");
    }
    await sleep(500);

    setProgress(100);
    log("","info");
    log("✓ HOÀN TẤT!","ok");
    log("  Xanh: "+fmtN(greenTotal)+" cấu kiện trong Excel","ok");
    log("  Còn lại: giữ màu gốc của model","info");
    log("","info");
    log("💡 Cấu kiện xanh sẽ nổi bật so với màu gốc!","info");
    setTimeout(function(){setProgress(0);},2000);

  }catch(err){
    log("✗ "+(err&&err.message?err.message:String(err)),"err");
    setProgress(0);
  }finally{
    lockUI(false);
    if(!_excelGuids.length)document.getElementById("applyBtn").disabled=true;
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
  var c=await api.view.createView({name:name,description:"Paint Approval Tool v7.0 | Le Van Thao"});
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
