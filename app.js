/**
 * Paint Approval Tool v2.0
 * ─────────────────────────────────────
 * GUID có trong Excel  →  xanh #34d874
 * Tất cả còn lại       →  xám  #8b95a8
 * ─────────────────────────────────────
 * FIX v2.0:
 *  - Auto convert UUID (36 chars, with dashes) ↔ IFC GUID (22 chars, base64)
 *  - Thử cả 2 format để match model bất kể IFC export kiểu nào
 *  - Chỉ tô xám các object KHÔNG xanh (không đè màu)
 *  - Verify match thật sự bằng cách check runtimeIds có nằm trong model không
 * Developed by Le Van Thao
 */

/* ═══════════════════════════════════════
   CONSTANTS
═══════════════════════════════════════ */
var COLOR_GREEN    = "#34d874";
var COLOR_GRAY     = "#8b95a8";
var RETRY_MAX      = 7;
var RETRY_DELAY_MS = 2000;
var BATCH_CONVERT  = 500;
var BATCH_COLOR    = 800;

/* ═══════════════════════════════════════
   STATE
═══════════════════════════════════════ */
var _api        = null;
var _excelGuids = [];

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
   ★★★ UUID ↔ IFC GUID CONVERSION ★★★
   
   UUID:     66192c05-bf8d-4916-a7fe-c86fcb140bfd  (36 chars)
   IFC GUID: 1ZQh_2aby78AfT_Lk02pHl                (22 chars, base64-like)
   
   IFC GUIDs use a compressed base64 encoding of the UUID.
═══════════════════════════════════════ */

// IFC GUID uses a 64-char alphabet (different from standard base64)
var IFC_BASE64_CHARS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_$";

/** Convert 4-char string from IFC base64 to 3 bytes */
function _cvTo64(num, nDigits) {
  var digits = [];
  for (var i = 0; i < nDigits; i++) {
    digits.push(IFC_BASE64_CHARS.charAt(num % 64));
    num = Math.floor(num / 64);
  }
  return digits.reverse().join("");
}
function _cvFrom64(str) {
  var result = 0;
  for (var i = 0; i < str.length; i++) {
    var idx = IFC_BASE64_CHARS.indexOf(str.charAt(i));
    if (idx < 0) return -1;
    result = result * 64 + idx;
  }
  return result;
}

/** UUID (36 chars with dashes) → IFC GUID (22 chars) */
function uuidToIfcGuid(uuid) {
  if (!uuid) return null;
  var hex = String(uuid).replace(/-/g, "").toLowerCase();
  if (hex.length !== 32) return null;
  if (!/^[0-9a-f]{32}$/.test(hex)) return null;

  // Parse 32 hex chars → 16 bytes → 6 groups
  // First group: 2 chars  → base64 2 chars
  // Next 5 groups: 6 chars → base64 4 chars each
  var num = [];
  num.push(parseInt(hex.substr(0, 2), 16));
  for (var i = 0; i < 5; i++) {
    num.push(parseInt(hex.substr(2 + i * 6, 6), 16));
  }

  var result = _cvTo64(num[0], 2);
  for (var i = 1; i < 6; i++) {
    result += _cvTo64(num[i], 4);
  }
  return result;
}

/** IFC GUID (22 chars) → UUID (36 chars with dashes) */
function ifcGuidToUuid(ifc) {
  if (!ifc || ifc.length !== 22) return null;
  var parts = [];
  parts.push(_cvFrom64(ifc.substr(0, 2)));
  for (var i = 0; i < 5; i++) {
    parts.push(_cvFrom64(ifc.substr(2 + i * 4, 4)));
  }
  if (parts.some(function(p){ return p < 0; })) return null;

  var hex = parts[0].toString(16).padStart(2, "0");
  for (var i = 1; i < 6; i++) {
    hex += parts[i].toString(16).padStart(6, "0");
  }
  return hex.substr(0,8) + "-" + hex.substr(8,4) + "-" + hex.substr(12,4) + "-" + hex.substr(16,4) + "-" + hex.substr(20,12);
}

/** Detect format of a GUID string */
function detectGuidFormat(g) {
  if (!g) return "unknown";
  var s = String(g).trim();
  if (s.length === 36 && /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(s)) return "uuid";
  if (s.length === 32 && /^[0-9a-f]{32}$/i.test(s)) return "uuid-nodash";
  if (s.length === 22) return "ifc";
  return "unknown";
}

/** Generate BOTH formats for each GUID — match bất kể model dùng format nào */
function expandGuidsToBothFormats(guids) {
  var out = [];     // array of { original, variants: [] }
  var seen = {};

  guids.forEach(function(g) {
    if (!g) return;
    var s = String(g).trim();
    var fmt = detectGuidFormat(s);
    var variants = [s];

    if (fmt === "uuid") {
      var ifc = uuidToIfcGuid(s);
      if (ifc) variants.push(ifc);
    } else if (fmt === "uuid-nodash") {
      // Thêm dashes
      var withDash = s.substr(0,8)+"-"+s.substr(8,4)+"-"+s.substr(12,4)+"-"+s.substr(16,4)+"-"+s.substr(20,12);
      variants.push(withDash);
      var ifc = uuidToIfcGuid(withDash);
      if (ifc) variants.push(ifc);
    } else if (fmt === "ifc") {
      var uuid = ifcGuidToUuid(s);
      if (uuid) variants.push(uuid);
    }

    // Dedupe variants per item
    variants = variants.filter(function(v){
      var key = "v:" + v;
      if (seen[key]) return false;
      seen[key] = true;
      return true;
    });

    if (variants.length) out.push({ original: s, variants: variants });
  });

  return out;
}

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

  // Log sample + format
  var fmt = out.length ? detectGuidFormat(out[0]) : "none";
  log('Sheet: "' + sheetName + '" | ' + rows.length + ' dòng | ' + out.length + ' GUID duy nhất', "info");
  log('GUID[0]: ' + (out[0]||"N/A") + ' (' + fmt + ')', "info");

  return out;
}

/* ═══════════════════════════════════════
   GET MODEL GROUPS — robust, multi-format
═══════════════════════════════════════ */
function parseIds(group) {
  var ids = {};
  if (!group) return [];
  if (Array.isArray(group.objects)) {
    group.objects.forEach(function(o) {
      var v = (o && o.id !== undefined) ? o.id : (o && o.runtimeId);
      if      (typeof v === "number")                    ids[v] = true;
      else if (typeof v === "string" && v && !isNaN(+v)) ids[+v] = true;
    });
  }
  if (Array.isArray(group.objectRuntimeIds)) {
    group.objectRuntimeIds.forEach(function(v) {
      if      (typeof v === "number")                    ids[v] = true;
      else if (typeof v === "string" && v && !isNaN(+v)) ids[+v] = true;
    });
  }
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
      throw new Error("Viewer không trả về object. Đợi model load xong rồi thử lại.");
    }

    var groups = raw
      .map(function(g){ return { modelId: g && g.modelId, runtimeIds: parseIds(g) }; })
      .filter(function(g){ return g.modelId && g.runtimeIds.length > 0; });

    if (!groups.length) {
      log("Không parse được runtimeId (attempt " + attempt + "/" + RETRY_MAX + ")...", "warn");
      if (attempt < RETRY_MAX) { await sleep(RETRY_DELAY_MS); continue; }
      throw new Error("Không đọc được runtimeId.");
    }

    groups.forEach(function(g){
      log("  ✓ model " + g.modelId + ": " + fmtNum(g.runtimeIds.length) + " objects", "info");
    });
    return groups;
  }
}

/* ═══════════════════════════════════════
   CONVERT GUIDs → runtimeIds
   Thử cả UUID và IFC format, dùng format nào match nhiều hơn
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

/**
 * Convert function: thử convert với 1 list GUID (cùng format)
 * Returns Map<originalIndex, number[]>
 */
async function tryConvert(api, modelId, guids) {
  var converted = [];
  for (var i = 0; i < guids.length; i += BATCH_CONVERT) {
    var chunk = guids.slice(i, i + BATCH_CONVERT);
    var res;
    try { res = await api.viewer.convertToObjectRuntimeIds(modelId, chunk); }
    catch(err) {
      log("  convert batch error: " + (err&&err.message?err.message:String(err)), "warn");
      for (var k=0; k<chunk.length; k++) converted.push(null);
      continue;
    }
    if (!Array.isArray(res)) {
      for (var k=0; k<chunk.length; k++) converted.push(null);
      continue;
    }
    converted = converted.concat(res);
  }
  return converted;
}

/**
 * Smart convert: thử cả UUID và IFC format, chọn format match tốt nhất
 * Returns Map<modelId, Set<runtimeId>> + match stats
 */
async function smartConvertGuids(api, modelGroups, originalGuids) {
  var matchByModel = new Map();
  var totalMatched = 0;

  // Phân loại original GUIDs theo format
  var uuidList = []; var ifcList = []; var others = [];
  originalGuids.forEach(function(g) {
    var fmt = detectGuidFormat(g);
    if (fmt === "uuid" || fmt === "uuid-nodash") uuidList.push(g);
    else if (fmt === "ifc") ifcList.push(g);
    else others.push(g);
  });

  log("  Format distribution: UUID=" + uuidList.length + ", IFC=" + ifcList.length + ", other=" + others.length, "info");

  // Tạo list IFC guids từ UUID (để dự phòng nếu model dùng IFC)
  var uuidToIfc = uuidList.map(function(u){ return uuidToIfcGuid(u); }).filter(function(v){ return !!v; });
  // Tạo list UUID từ IFC
  var ifcToUuid = ifcList.map(function(i){ return ifcGuidToUuid(i); }).filter(function(v){ return !!v; });

  for (var gi = 0; gi < modelGroups.length; gi++) {
    var group = modelGroups[gi];
    var modelRuntimeIds = new Set(group.runtimeIds);  // để verify
    var modelMatches = new Set();

    // Helper: thử convert 1 list, chỉ giữ những runtimeId có thật trong model
    async function tryAndVerify(guidsList, labelFmt) {
      if (!guidsList.length) return 0;
      var converted = await tryConvert(api, group.modelId, guidsList);
      var hit = 0;
      var realHit = 0;
      for (var i = 0; i < guidsList.length; i++) {
        var ids = flattenConvertResult(converted[i]);
        if (!ids.length) continue;
        hit++;
        // VERIFY: chỉ giữ runtimeId có thật trong model
        ids.forEach(function(id) {
          if (modelRuntimeIds.has(id)) {
            modelMatches.add(id);
            realHit++;
          }
        });
      }
      log("  [" + labelFmt + "] model " + group.modelId + ": convert=" + hit + "/" + guidsList.length + ", verified=" + realHit, realHit > 0 ? "ok" : "warn");
      return realHit;
    }

    // Thử theo thứ tự
    if (uuidList.length)  await tryAndVerify(uuidList,  "UUID");
    if (ifcList.length)   await tryAndVerify(ifcList,   "IFC ");
    if (uuidToIfc.length) await tryAndVerify(uuidToIfc, "U→I ");  // UUID-converted-to-IFC
    if (ifcToUuid.length) await tryAndVerify(ifcToUuid, "I→U ");  // IFC-converted-to-UUID
    if (others.length)    await tryAndVerify(others,    "RAW ");

    if (modelMatches.size > 0) {
      matchByModel.set(group.modelId, modelMatches);
      totalMatched += modelMatches.size;
    }
  }

  return matchByModel;
}

/* ═══════════════════════════════════════
   PAINT
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
   MAIN
═══════════════════════════════════════ */
async function applyColors() {
  lockUI(true);
  clearLog();
  setProgress(5);

  try {
    if (!_excelGuids.length) throw new Error("Chưa có GUID. Chọn file Excel trước.");

    var api = await getAPI();

    log("Đang lấy danh sách object từ viewer...", "info");
    setProgress(10);
    var modelGroups  = await getModelGroups();
    var totalObjects = modelGroups.reduce(function(s,g){ return s + g.runtimeIds.length; }, 0);
    setStat("s-total", fmtNum(totalObjects));
    setStat("s-excel", fmtNum(_excelGuids.length));

    setProgress(28);
    log("Đang map " + _excelGuids.length + " GUID → runtimeId (thử nhiều format)...", "info");
    var matchMap = await smartConvertGuids(api, modelGroups, _excelGuids);

    var greenTotal = 0;
    matchMap.forEach(function(s){ greenTotal += s.size; });
    var grayTotal  = totalObjects - greenTotal;
    setStat("s-green", fmtNum(greenTotal));
    setStat("s-gray",  fmtNum(grayTotal));

    if (greenTotal === 0) {
      log("✗ VẪN không match được object nào!", "err");
      log("  Có thể model dùng format GUID khác hoàn toàn.", "err");
      log("  Hãy gửi log này cho Thảo để kiểm tra.", "err");
      setProgress(0);
      lockUI(false);
      return;
    }

    setProgress(45);
    log("Đang tô màu... (xanh: " + fmtNum(greenTotal) + ", xám: " + fmtNum(grayTotal) + ")", "info");

    var totalSteps = modelGroups.length * 2;
    var done = 0;

    // ★ QUAN TRỌNG: Tô xám TRƯỚC (chỉ các ID không phải xanh), xong rồi tô xanh
    // Cách cũ bị lỗi vì tô xám HẾT rồi xanh đè — lệnh cuối có thể bị Trimble batch ngược
    for (var gi = 0; gi < modelGroups.length; gi++) {
      var group    = modelGroups[gi];
      var greenSet = matchMap.get(group.modelId) || new Set();
      var greenIds = Array.from(greenSet);
      var grayIds  = group.runtimeIds.filter(function(id){ return !greenSet.has(id); });

      if (grayIds.length) {
        await paintIds(api, group.modelId, grayIds, COLOR_GRAY);
        log("  ▫ Tô xám " + fmtNum(grayIds.length) + " objects (" + group.modelId.substr(0,10) + "...)", "ok");
      }
      done++;
      setProgress(45 + (done/totalSteps)*50);
    }

    // Tô xanh SAU KHI xám xong hoàn toàn
    for (var gi = 0; gi < modelGroups.length; gi++) {
      var group    = modelGroups[gi];
      var greenSet = matchMap.get(group.modelId) || new Set();
      var greenIds = Array.from(greenSet);

      if (greenIds.length) {
        await paintIds(api, group.modelId, greenIds, COLOR_GREEN);
        log("  ▪ Tô xanh " + fmtNum(greenIds.length) + " objects (" + group.modelId.substr(0,10) + "...)", "ok");
      }
      done++;
      setProgress(45 + (done/totalSteps)*50);
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
    var created = await api.view.createView({ name: name, description: "Paint Approval Tool v2.0 | Le Van Thao" });
    if (!created || !created.id) throw new Error("No view ID returned.");
    await api.view.updateView({ id: created.id });
    await api.view.selectView(created.id);
    log('✓ Đã lưu view: "' + name + '"', "ok");
  } catch(err) {
    log("✗ Save view: " + (err&&err.message?err.message:String(err)), "err");
  }
}

/* ═══════════════════════════════════════
   EVENTS
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
