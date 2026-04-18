/**
 * Paint Approval Tool v1.0
 * ─────────────────────────────────────
 * GUID có trong Excel  →  xanh #34d874
 * Tất cả còn lại       →  xám  #8b95a8
 * ─────────────────────────────────────
 * Developed by Le Van Thao
 */

/* ═══════════════════════════════════════
   CONSTANTS
═══════════════════════════════════════ */
var COLOR_GREEN    = "#34d874";
var COLOR_GRAY     = "#8b95a8";
var RETRY_MAX      = 7;      // số lần thử getObjects
var RETRY_DELAY_MS = 2000;   // ms chờ giữa các lần thử
var BATCH_CONVERT  = 500;    // batch convertToObjectRuntimeIds
var BATCH_COLOR    = 800;    // batch setObjectState

/* ═══════════════════════════════════════
   STATE
═══════════════════════════════════════ */
var _api        = null;
var _excelGuids = [];   // string[]

/* ═══════════════════════════════════════
   UI
═══════════════════════════════════════ */
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
function clearLog() {
  var el = document.getElementById("log");
  if (el) el.innerHTML = "";
}
function setStat(id, v) {
  var el = document.getElementById(id);
  if (el) el.textContent = (v !== null && v !== undefined) ? v : "—";
}
function resetStats() {
  ["s-total","s-excel","s-green","s-gray"].forEach(function(id){ setStat(id, "—"); });
}
function setProgress(pct) {
  var wrap = document.getElementById("progWrap");
  var bar  = document.getElementById("progBar");
  if (!wrap || !bar) return;
  if (pct <= 0) { wrap.classList.remove("on"); bar.style.width = "0%"; return; }
  wrap.classList.add("on");
  bar.style.width = Math.min(pct, 100) + "%";
}
function lockUI(yes) {
  ["applyBtn","resetBtn","saveBtn"].forEach(function(id){
    var el = document.getElementById(id);
    if (el) el.disabled = yes;
  });
}
function sleep(ms) { return new Promise(function(r){ setTimeout(r, ms); }); }
function pad2(n)   { return String(n).padStart(2,"0"); }
function fmtNum(n) { return (typeof n === "number") ? n.toLocaleString() : String(n); }

/* ═══════════════════════════════════════
   TRIMBLE API
═══════════════════════════════════════ */
async function getAPI() {
  if (_api) return _api;
  _api = await TrimbleConnectWorkspace.connect(window.parent, function(ev, data){
    console.log("[Trimble]", ev, data);
  });
  log("Đã kết nối Trimble Workspace API.", "ok");
  return _api;
}

/* ═══════════════════════════════════════
   EXCEL
═══════════════════════════════════════ */
function readWorkbook(file) {
  return new Promise(function(resolve, reject) {
    var reader = new FileReader();
    reader.onload = function(e) {
      try { resolve(XLSX.read(e.target.result, { type: "array" })); }
      catch(err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function extractGuids(wb) {
  if (!wb || !wb.SheetNames || !wb.SheetNames.length)
    throw new Error("File Excel không có sheet.");
  var sheetName = wb.SheetNames[0];
  var rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: "" });
  if (!rows.length) throw new Error("Sheet đầu tiên không có dữ liệu.");

  // Tìm cột GUID (case-insensitive)
  var keys    = Object.keys(rows[0]);
  var guidKey = keys.find(function(k){ return k.trim().toUpperCase() === "GUID"; });
  if (!guidKey) {
    guidKey = keys[0];
    log('⚠ Không thấy cột "GUID", dùng cột đầu: "' + guidKey + '"', "warn");
  }

  var seen = {};
  var out  = [];
  rows.forEach(function(r) {
    var g = String(r[guidKey] || "").trim();
    if (g && !seen[g]) { seen[g] = true; out.push(g); }
  });

  log('Sheet: "' + sheetName + '" | ' + rows.length + ' dòng | ' + out.length + ' GUID duy nhất', "info");
  return out;
}

/* ═══════════════════════════════════════
   GET MODEL GROUPS — robust, multi-format
═══════════════════════════════════════ */

/** Đọc runtimeIds từ 1 group object, hỗ trợ tất cả format Trimble đã biết */
function parseIds(group) {
  var ids = {};
  if (!group) return [];

  // Format A: group.objects[].id  (phổ biến nhất)
  if (Array.isArray(group.objects)) {
    group.objects.forEach(function(o) {
      var v = (o && o.id !== undefined) ? o.id : (o && o.runtimeId);
      if      (typeof v === "number")                    ids[v] = true;
      else if (typeof v === "string" && v && !isNaN(+v)) ids[+v] = true;
    });
  }
  // Format B: group.objectRuntimeIds[]
  if (Array.isArray(group.objectRuntimeIds)) {
    group.objectRuntimeIds.forEach(function(v) {
      if      (typeof v === "number")                    ids[v] = true;
      else if (typeof v === "string" && v && !isNaN(+v)) ids[+v] = true;
    });
  }
  // Format C: group.ids[]
  if (Array.isArray(group.ids)) {
    group.ids.forEach(function(v) {
      if      (typeof v === "number")                    ids[v] = true;
      else if (typeof v === "string" && v && !isNaN(+v)) ids[+v] = true;
    });
  }
  return Object.keys(ids).map(Number);
}

async function getModelGroups() {
  var api = await getAPI();

  for (var attempt = 1; attempt <= RETRY_MAX; attempt++) {
    var raw;
    try { raw = await api.viewer.getObjects(); }
    catch(err) {
      log("getObjects error (attempt " + attempt + "): " + (err && err.message ? err.message : String(err)), "warn");
      if (attempt < RETRY_MAX) { await sleep(RETRY_DELAY_MS); continue; }
      throw err;
    }

    if (!Array.isArray(raw) || !raw.length) {
      log("Viewer chưa có object (attempt " + attempt + "/" + RETRY_MAX + "), đợi " + (RETRY_DELAY_MS/1000) + "s...", "warn");
      if (attempt < RETRY_MAX) { await sleep(RETRY_DELAY_MS); continue; }
      throw new Error("Viewer không trả về object sau " + RETRY_MAX + " lần thử.\nHãy đảm bảo model đã load xong rồi thử lại.");
    }

    // Debug log
    log("[getObjects] " + raw.length + " group(s) | keys: " + Object.keys(raw[0]||{}).join(", "), "info");

    var groups = raw
      .map(function(g){ return { modelId: g && g.modelId, runtimeIds: parseIds(g) }; })
      .filter(function(g){ return g.modelId && g.runtimeIds.length > 0; });

    if (!groups.length) {
      log("Không parse được runtimeId (attempt " + attempt + "/" + RETRY_MAX + ")...", "warn");
      if (attempt < RETRY_MAX) { await sleep(RETRY_DELAY_MS); continue; }
      throw new Error("Không đọc được runtimeId.\nXem log '[getObjects] keys' để tìm format đúng.");
    }

    groups.forEach(function(g){
      log("  ✓ model " + g.modelId + ": " + fmtNum(g.runtimeIds.length) + " objects", "info");
    });
    return groups;
  }
}

/* ═══════════════════════════════════════
   CONVERT GUIDs → runtimeIds
═══════════════════════════════════════ */
function flattenConvertResult(val) {
  if (val === null || val === undefined) return [];
  if (typeof val === "number") return [val];
  if (Array.isArray(val)) {
    var out = [];
    val.forEach(function(v){
      if (typeof v === "number") out.push(v);
      else if (Array.isArray(v)) v.forEach(function(vv){ if(typeof vv==="number") out.push(vv); });
    });
    return out;
  }
  return [];
}

async function convertGuids(api, modelGroups, guids) {
  // returns Map<modelId, Set<runtimeId>>
  var result = new Map();

  for (var gi = 0; gi < modelGroups.length; gi++) {
    var group   = modelGroups[gi];
    var chunks  = [];
    for (var i = 0; i < guids.length; i += BATCH_CONVERT)
      chunks.push(guids.slice(i, i + BATCH_CONVERT));

    var converted = [];
    for (var ci = 0; ci < chunks.length; ci++) {
      var res;
      try { res = await api.viewer.convertToObjectRuntimeIds(group.modelId, chunks[ci]); }
      catch(err) {
        log("Convert error model " + group.modelId + " batch " + (ci+1) + ": " + (err&&err.message?err.message:String(err)), "warn");
        for (var k=0; k<chunks[ci].length; k++) converted.push(null);
        continue;
      }
      if (!Array.isArray(res)) {
        for (var k=0; k<chunks[ci].length; k++) converted.push(null);
        continue;
      }
      converted = converted.concat(res);
    }

    var hit = 0;
    for (var i = 0; i < guids.length; i++) {
      var ids = flattenConvertResult(converted[i]);
      if (!ids.length) continue;
      if (!result.has(group.modelId)) result.set(group.modelId, new Set());
      ids.forEach(function(id){ result.get(group.modelId).add(id); });
      hit++;
    }
    log("  Model " + group.modelId + ": " + hit + "/" + guids.length + " GUID matched", hit > 0 ? "ok" : "warn");
  }
  return result;
}

/* ═══════════════════════════════════════
   SET COLOR BATCH
═══════════════════════════════════════ */
async function paintIds(api, modelId, runtimeIds, color) {
  for (var i = 0; i < runtimeIds.length; i += BATCH_COLOR) {
    var chunk = runtimeIds.slice(i, i + BATCH_COLOR);
    await api.viewer.setObjectState(
      { modelObjectIds: [{ modelId: modelId, objectRuntimeIds: chunk }] },
      { color: color }
    );
  }
}

/* ═══════════════════════════════════════
   MAIN — APPLY COLORS
═══════════════════════════════════════ */
async function applyColors() {
  lockUI(true);
  clearLog();
  setProgress(5);

  try {
    if (!_excelGuids.length) throw new Error("Chưa có GUID. Hãy chọn file Excel trước.");

    var api = await getAPI();

    log("Đang lấy danh sách object từ viewer...", "info");
    setProgress(10);
    var modelGroups  = await getModelGroups();
    var totalObjects = modelGroups.reduce(function(s,g){ return s + g.runtimeIds.length; }, 0);
    setStat("s-total", fmtNum(totalObjects));
    setStat("s-excel", fmtNum(_excelGuids.length));

    setProgress(28);
    log("Đang map " + _excelGuids.length + " GUID → runtimeId...", "info");
    var matchMap = await convertGuids(api, modelGroups, _excelGuids);

    var greenTotal = 0;
    matchMap.forEach(function(s){ greenTotal += s.size; });
    var grayTotal  = totalObjects - greenTotal;
    setStat("s-green", fmtNum(greenTotal));
    setStat("s-gray",  fmtNum(grayTotal));

    if (greenTotal === 0) {
      log("⚠ Không match được object nào!", "warn");
      log("  GUID[0]: " + (_excelGuids[0]||"N/A"), "warn");
      log("  Kiểm tra lại tên cột trong Excel (phải là GUID)", "warn");
    }

    setProgress(42);
    log("Đang tô màu...", "info");

    var totalSteps = modelGroups.length * 2;
    var done = 0;

    for (var gi = 0; gi < modelGroups.length; gi++) {
      var group    = modelGroups[gi];
      var greenSet = matchMap.get(group.modelId) || new Set();
      var greenIds = Array.from(greenSet);
      var grayIds  = group.runtimeIds.filter(function(id){ return !greenSet.has(id); });

      // Tô xám trước
      if (grayIds.length) {
        await paintIds(api, group.modelId, grayIds, COLOR_GRAY);
        log("  Tô xám " + fmtNum(grayIds.length) + " objects (model " + group.modelId + ")", "ok");
      }
      done++;
      setProgress(42 + (done/totalSteps)*53);

      // Tô xanh đè
      if (greenIds.length) {
        await paintIds(api, group.modelId, greenIds, COLOR_GREEN);
        log("  Tô xanh " + fmtNum(greenIds.length) + " objects (model " + group.modelId + ")", "ok");
      }
      done++;
      setProgress(42 + (done/totalSteps)*53);
    }

    setProgress(100);
    log("✓ Hoàn tất! Xanh: " + fmtNum(greenTotal) + " | Xám: " + fmtNum(grayTotal), "ok");
    setTimeout(function(){ setProgress(0); }, 1800);

  } catch(err) {
    log("✗ " + (err && err.message ? err.message : String(err)), "err");
    setProgress(0);
  } finally {
    lockUI(false);
    if (!_excelGuids.length) document.getElementById("applyBtn").disabled = true;
  }
}

/* ═══════════════════════════════════════
   RESET
═══════════════════════════════════════ */
async function resetViewer() {
  lockUI(true);
  clearLog();
  setProgress(10);
  try {
    var api = await getAPI();
    try { await api.viewer.setObjectState(undefined, { color: "reset", visible: "reset" }); }
    catch(e) { log("color reset fallback: " + (e&&e.message?e.message:String(e)), "warn"); }
    await api.viewer.reset();
    resetStats();
    setProgress(100);
    log("✓ Đã reset viewer.", "ok");
    setTimeout(function(){ setProgress(0); }, 1000);
  } catch(err) {
    log("✗ " + (err&&err.message?err.message:String(err)), "err");
    setProgress(0);
  } finally {
    lockUI(false);
    if (!_excelGuids.length) document.getElementById("applyBtn").disabled = true;
  }
}

/* ═══════════════════════════════════════
   SAVE VIEW
═══════════════════════════════════════ */
async function saveView() {
  try {
    var api  = await getAPI();
    var inp  = document.getElementById("viewName");
    var name = inp ? inp.value.trim() : "";
    if (!name) {
      var now = new Date();
      name = "Approval " + now.getFullYear() + "-" + pad2(now.getMonth()+1) + "-" + pad2(now.getDate())
           + " " + pad2(now.getHours()) + ":" + pad2(now.getMinutes());
      if (inp) inp.value = name;
    }
    var created = await api.view.createView({ name: name, description: "Paint Approval Tool v1.0 | Le Van Thao" });
    if (!created || !created.id) throw new Error("No view ID returned.");
    await api.view.updateView({ id: created.id });
    await api.view.selectView(created.id);
    log('✓ Đã lưu view: "' + name + '"', "ok");
  } catch(err) {
    log("✗ Save view: " + (err&&err.message?err.message:String(err)), "err");
  }
}

/* ═══════════════════════════════════════
   EVENT LISTENERS
═══════════════════════════════════════ */
document.getElementById("fileInput").addEventListener("change", async function(e) {
  var file = e.target.files && e.target.files[0];
  if (!file) return;
  document.getElementById("fileName").textContent = file.name;
  clearLog();
  log('Đang đọc "' + file.name + '"...', "info");
  try {
    var wb = await readWorkbook(file);
    _excelGuids = extractGuids(wb);
    setStat("s-excel", fmtNum(_excelGuids.length));
    if (_excelGuids.length > 0) {
      document.getElementById("applyBtn").disabled = false;
      log('✓ Sẵn sàng! Nhấn "Áp màu" để bắt đầu.', "ok");
    } else {
      log("⚠ Không tìm thấy GUID nào.", "warn");
    }
  } catch(err) {
    log("✗ " + (err&&err.message?err.message:String(err)), "err");
    _excelGuids = [];
    document.getElementById("applyBtn").disabled = true;
  }
});

// Drag & drop
(function() {
  var zone = document.getElementById("uploadZone");
  zone.addEventListener("dragover",  function(e){ e.preventDefault(); zone.classList.add("over"); });
  zone.addEventListener("dragleave", function()  { zone.classList.remove("over"); });
  zone.addEventListener("drop", function(e) {
    e.preventDefault();
    zone.classList.remove("over");
    var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) {
      document.getElementById("fileInput").files = e.dataTransfer.files;
      document.getElementById("fileInput").dispatchEvent(new Event("change"));
    }
  });
})();

document.getElementById("applyBtn").addEventListener("click", applyColors);
document.getElementById("resetBtn").addEventListener("click", resetViewer);
document.getElementById("saveBtn").addEventListener("click",  saveView);
