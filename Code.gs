/**
 * 數辦設備借用清單 — 後端 API(Google Apps Script)
 *
 * 速度設計:
 *  1. 全部資料一次批次讀取(getDataRange().getValues()),不逐格存取。
 *  2. CacheService 快取 30 秒:多人同時開頁面時,只有第一個人真的讀試算表。
 *  3. 寫入用 LockService 防止同時借用造成超借,寫完立即失效快取並回傳最新資料,
 *     前端不必再發第二次請求。
 *
 * 部署步驟見「部署說明.md」。第一次使用請先執行 setup()。
 */

var SHEET_ITEMS = "設備";
var SHEET_LOANS = "紀錄";
var CACHE_KEY = "all_v1";
var CACHE_SEC = 30;
var HISTORY_LIMIT = 100; // 已歸還紀錄最多回傳筆數,避免資料越長越慢

/* ---------- HTTP 入口 ---------- */

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || "all";
  if (action === "all") return jsonOut(getAll());
  return jsonOut({ ok: false, error: "未知的 action:" + action });
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ ok: false, error: "請求格式錯誤" });
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return jsonOut({ ok: false, error: "系統忙碌中,請再試一次" });
  }

  try {
    if (req.action === "borrow") return jsonOut(doBorrow(req));
    if (req.action === "return") return jsonOut(doReturn(req));
    return jsonOut({ ok: false, error: "未知的 action:" + req.action });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 讀取(含快取) ---------- */

function getAll() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(CACHE_KEY);
  if (hit) return JSON.parse(hit);

  var data = readAllFresh();
  try {
    cache.put(CACHE_KEY, JSON.stringify(data), CACHE_SEC);
  } catch (err) {
    // 資料超過快取上限(100KB)時直接略過快取,不影響功能
  }
  return data;
}

function readAllFresh() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var itemRows = ss.getSheetByName(SHEET_ITEMS).getDataRange().getValues();
  var equipment = [];
  for (var i = 1; i < itemRows.length; i++) {
    var r = itemRows[i];
    if (!r[0]) continue;
    equipment.push({
      id: String(r[0]),
      name: String(r[1]),
      cat: String(r[2]),
      total: Number(r[3]) || 0,
      note: String(r[4] || "")
    });
  }

  var loanRows = ss.getSheetByName(SHEET_LOANS).getDataRange().getValues();
  var active = [], done = [];
  for (var j = loanRows.length - 1; j >= 1; j--) { // 由新到舊
    var l = loanRows[j];
    if (!l[0]) continue;
    var rec = {
      id: String(l[0]),
      itemId: String(l[1]),
      itemName: String(l[2]),
      qty: Number(l[3]) || 0,
      borrower: String(l[4]),
      unit: String(l[5]),
      date: dateStr(l[6]),
      slot: String(l[7] || ""),
      due: dateStr(l[8]),
      status: String(l[9]),
      returned: dateStr(l[10]),
      note: String(l[11] || "")
    };
    if (rec.status === "借用中") active.push(rec);
    else if (done.length < HISTORY_LIMIT) done.push(rec);
  }

  return { ok: true, equipment: equipment, loans: active.concat(done) };
}

function dateStr(v) {
  if (!v) return "";
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(v);
}

/* ---------- 寫入 ---------- */

function doBorrow(req) {
  var qty = Math.floor(Number(req.qty));
  if (!req.itemId || !qty || qty < 1) throw new Error("借用資料不完整");
  if (!req.borrower || !req.unit) throw new Error("請填寫借用人與班級/處室");

  var data = readAllFresh();
  var item = null;
  for (var i = 0; i < data.equipment.length; i++) {
    if (data.equipment[i].id === String(req.itemId)) { item = data.equipment[i]; break; }
  }
  if (!item) throw new Error("找不到設備:" + req.itemId);

  var out = 0;
  for (var j = 0; j < data.loans.length; j++) {
    var l = data.loans[j];
    if (l.status === "借用中" && l.itemId === item.id) out += l.qty;
  }
  var avail = item.total - out;
  if (qty > avail) throw new Error("可借數量不足(剩 " + avail + " 件)");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOANS);
  var id = "L" + Utilities.getUuid().slice(0, 8);
  sheet.appendRow([
    id, item.id, item.name, qty,
    String(req.borrower), String(req.unit),
    String(req.date || ""), String(req.slot || ""), String(req.due || ""),
    "借用中", "", String(req.note || ""),
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
  ]);

  return invalidateAndReturn();
}

function doReturn(req) {
  if (!req.loanId) throw new Error("缺少紀錄編號");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOANS);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(req.loanId)) {
      if (String(rows[i][9]) === "已歸還") throw new Error("這筆紀錄已經歸還過了");
      var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
      sheet.getRange(i + 1, 10, 1, 2).setValues([["已歸還", today]]);
      return invalidateAndReturn();
    }
  }
  throw new Error("找不到這筆借用紀錄,請重新整理頁面");
}

function invalidateAndReturn() {
  CacheService.getScriptCache().remove(CACHE_KEY);
  var data = readAllFresh();
  try {
    CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(data), CACHE_SEC);
  } catch (err) {}
  return data;
}

/* ---------- 初始化:建立工作表與範例設備 ----------
 * 在 Apps Script 編輯器選擇 setup 後按「執行」,授權後即完成。
 * 已存在的工作表不會被覆蓋,可放心重複執行。
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(SHEET_ITEMS)) {
    var s1 = ss.insertSheet(SHEET_ITEMS);
    var items = [
      ["編號", "設備名稱", "分類", "總數", "備註"],
      ["T01", "iPad 平板(充電車 A)", "學習載具", 30, "含充電車,借用請整車推走"],
      ["T02", "Chromebook 筆電", "學習載具", 20, ""],
      ["P01", "單槍投影機", "影音設備", 3, "附 HDMI 線"],
      ["P02", "實物投影機", "影音設備", 4, ""],
      ["A01", "行動擴音音箱", "影音設備", 5, "含頭戴麥克風"],
      ["C01", "網路攝影機", "影音設備", 6, "視訊、直播用"],
      ["M01", "無線麥克風組", "影音設備", 3, "一組兩支"],
      ["V01", "VR 眼鏡", "新興科技", 10, "使用後請酒精擦拭"],
      ["R01", "藍牙簡報筆", "週邊配件", 8, ""],
      ["H01", "HDMI 訊號線(5m)", "週邊配件", 10, ""],
      ["E01", "動力延長線", "週邊配件", 12, ""],
      ["S01", "相機三腳架", "週邊配件", 4, ""]
    ];
    s1.getRange(1, 1, items.length, 5).setValues(items);
    s1.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#E1F1EB");
    s1.setFrozenRows(1);
    s1.setColumnWidth(2, 220);
    s1.setColumnWidth(5, 260);
  }

  if (!ss.getSheetByName(SHEET_LOANS)) {
    var s2 = ss.insertSheet(SHEET_LOANS);
    var head = [["紀錄ID", "設備編號", "設備名稱", "數量", "借用人", "班級/處室",
                 "借用日期", "使用時段", "預計歸還日", "狀態", "歸還日期", "備註", "登記時間"]];
    s2.getRange(1, 1, 1, 13).setValues(head).setFontWeight("bold").setBackground("#E1F1EB");
    s2.setFrozenRows(1);
    s2.setColumnWidth(3, 200);
    // 借用/歸還日期欄以純文字保存,避免時區換算誤差
    s2.getRange("G:I").setNumberFormat("@");
    s2.getRange("K:K").setNumberFormat("@");
  }

  // 刪掉預設的空白工作表(如果還在而且不只一張)
  var def = ss.getSheetByName("工作表1") || ss.getSheetByName("Sheet1");
  if (def && ss.getSheets().length > 2) ss.deleteSheet(def);
}
