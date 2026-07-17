/**
 * 數辦設備借用清單 — 後端 API(Google Apps Script)v2:序號管理
 *
 * 速度設計:
 *  1. 全部資料一次批次讀取(getDataRange().getValues()),不逐格存取。
 *  2. CacheService 快取 30 秒:多人同時開頁面時,只有第一個人真的讀試算表。
 *  3. 寫入用 LockService 防止同時借用造成超借,寫完立即失效快取並回傳最新資料。
 *
 * v2 新增:
 *  - 每件設備有獨立序號。「設備」工作表 F 欄「序號清單」可填正式財產編號
 *    (逗號分隔);留空則自動以「編號-01、編號-02…」產生,數量依 D 欄總數。
 *  - 借用登記記錄實際借出的序號;歸還可只勾選部分序號(拆單歸還)。
 *
 * v3 新增:
 *  - 盤點機制:「盤點」工作表逐序號記錄盤點結果(正常/異常/未尋獲/長期借用),
 *    非「正常」的序號自動停借;每項設備回傳「上次盤點日」。
 *  - 序號改用逗號/頓號/分號分隔(不再用空白切割),編號可含空格。
 *  - importEquipment():一次性匯入「設備借用動態總表」的正式設備與編號。
 *
 * 第一次使用執行 setup();從 v1 升級執行 upgrade()。
 */

var SHEET_ITEMS = "設備";
var SHEET_LOANS = "紀錄";
var SHEET_INV = "盤點";
var CACHE_KEY = "all_v3";
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
    if (req.action === "inventory") return jsonOut(doInventory(req));
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
      note: String(r[4] || ""),
      serials: String(r[5] || "")   // F 欄:自訂序號清單,留空 = 自動產生
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
      note: String(l[11] || ""),
      serials: String(l[13] || "")  // N 欄:此筆借出的序號
    };
    if (rec.status === "借用中") active.push(rec);
    else if (done.length < HISTORY_LIMIT) done.push(rec);
  }

  // 盤點:取每個序號最新一筆結果,設備附上「上次盤點日」與停借清單
  var invSheet = ss.getSheetByName(SHEET_INV);
  var latest = {};
  if (invSheet) {
    var invRows = invSheet.getDataRange().getValues();
    for (var k = 1; k < invRows.length; k++) {
      var v = invRows[k];
      if (!v[2] || !v[4]) continue;
      latest[String(v[2]) + "|" + String(v[4])] = {
        d: dateStr(v[1]), r: String(v[5]), note: String(v[6] || "")
      };
    }
  }
  for (var m = 0; m < equipment.length; m++) {
    var it = equipment[m];
    var list = serialsOfItem(it);
    var last = "", bad = [];
    for (var n = 0; n < list.length; n++) {
      var rec2 = latest[it.id + "|" + list[n]];
      if (!rec2) continue;
      if (rec2.d > last) last = rec2.d;
      if (rec2.r && rec2.r !== "正常") bad.push({ s: list[n], r: rec2.r, note: rec2.note });
    }
    it.lastCheck = last;
    it.bad = bad;
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

/* ---------- 序號工具 ---------- */

function parseList(s) {
  // 只用逗號/頓號/分號/換行切割,序號本身可以含空格(例:DJI NEO-01)
  return String(s || "").split(/[,、;\n]+/)
    .map(function (x) { return x.trim(); })
    .filter(function (x) { return x; });
}

// 一件設備的全部序號:F 欄有填就用它,否則依總數自動產生 編號-01…
function serialsOfItem(item) {
  var custom = parseList(item.serials);
  if (custom.length) return custom;
  var n = item.total, pad = n >= 100 ? 3 : 2, out = [];
  for (var i = 1; i <= n; i++) {
    out.push(item.id + "-" + ("000" + i).slice(-pad));
  }
  return out;
}

/* ---------- 寫入 ---------- */

function doBorrow(req) {
  var serials = (req.serials || []).map(String);
  if (!req.itemId || !serials.length) throw new Error("請至少選擇一個序號");
  if (!req.borrower || !req.unit) throw new Error("請填寫借用人與班級/處室");

  var data = readAllFresh();
  var item = null;
  for (var i = 0; i < data.equipment.length; i++) {
    if (data.equipment[i].id === String(req.itemId)) { item = data.equipment[i]; break; }
  }
  if (!item) throw new Error("找不到設備:" + req.itemId);

  // 收集已借出的序號,逐一檢查要借的序號還在架上
  var outSet = {};
  for (var j = 0; j < data.loans.length; j++) {
    var l = data.loans[j];
    if (l.status !== "借用中" || l.itemId !== item.id) continue;
    parseList(l.serials).forEach(function (s) { outSet[s] = true; });
  }
  var badSet = {};
  (item.bad || []).forEach(function (b) { badSet[b.s] = b.r; });

  var all = serialsOfItem(item);
  for (var k = 0; k < serials.length; k++) {
    if (all.indexOf(serials[k]) === -1) throw new Error("序號不存在:" + serials[k]);
    if (outSet[serials[k]]) throw new Error("序號 " + serials[k] + " 剛被借走了,請重新選擇");
    if (badSet[serials[k]]) throw new Error("序號 " + serials[k] + " 停借中(" + badSet[serials[k]] + ")");
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOANS);
  var id = "L" + Utilities.getUuid().slice(0, 8);
  sheet.appendRow([
    id, item.id, item.name, serials.length,
    String(req.borrower), String(req.unit),
    String(req.date || ""), String(req.slot || ""), String(req.due || ""),
    "借用中", "", String(req.note || ""),
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
    serials.join(", ")
  ]);

  return invalidateAndReturn();
}

function doReturn(req) {
  if (!req.loanId) throw new Error("缺少紀錄編號");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOANS);
  var rows = sheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(req.loanId)) continue;
    if (String(rows[i][9]) === "已歸還") throw new Error("這筆紀錄已經歸還過了");

    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var rowSerials = parseList(rows[i][13]);
    var toReturn = (req.serials || []).map(String);

    // 部分歸還:勾選的序號拆成一筆「已歸還」,原紀錄留下未還的序號
    if (toReturn.length && rowSerials.length && toReturn.length < rowSerials.length) {
      var remain = rowSerials.filter(function (s) { return toReturn.indexOf(s) === -1; });
      if (remain.length + toReturn.length !== rowSerials.length) {
        throw new Error("勾選的序號與紀錄不符,請重新整理頁面");
      }
      sheet.getRange(i + 1, 4).setValue(remain.length);              // 數量
      sheet.getRange(i + 1, 14).setValue(remain.join(", "));         // 剩餘序號
      var r = rows[i];
      sheet.appendRow([
        "L" + Utilities.getUuid().slice(0, 8), r[1], r[2], toReturn.length,
        r[4], r[5], dateStr(r[6]), r[7], dateStr(r[8]),
        "已歸還", today, r[11],
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
        toReturn.join(", ")
      ]);
    } else {
      // 全部歸還
      sheet.getRange(i + 1, 10, 1, 2).setValues([["已歸還", today]]);
    }
    return invalidateAndReturn();
  }
  throw new Error("找不到這筆借用紀錄,請重新整理頁面");
}

/* ---------- 盤點 ---------- */

var INV_RESULTS = { "正常": 1, "異常": 1, "未尋獲": 1, "長期借用": 1 };

function doInventory(req) {
  var results = req.results || [];
  if (!req.itemId || !results.length) throw new Error("沒有盤點資料");
  if (!req.checker) throw new Error("請填寫盤點人");

  var data = readAllFresh();
  var item = null;
  for (var i = 0; i < data.equipment.length; i++) {
    if (data.equipment[i].id === String(req.itemId)) { item = data.equipment[i]; break; }
  }
  if (!item) throw new Error("找不到設備:" + req.itemId);

  var all = serialsOfItem(item);
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var rows = [];
  for (var j = 0; j < results.length; j++) {
    var s = String(results[j].s || ""), r = String(results[j].r || "");
    if (all.indexOf(s) === -1) throw new Error("序號不存在:" + s);
    if (!INV_RESULTS[r]) throw new Error("盤點結果不正確:" + r);
    rows.push([
      "I" + Utilities.getUuid().slice(0, 8), today, item.id, item.name,
      s, r, String(req.note || ""), String(req.checker), ts
    ]);
  }

  var sheet = ensureInvSheet();
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
  return invalidateAndReturn();
}

function ensureInvSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SHEET_INV);
  if (!s) {
    s = ss.insertSheet(SHEET_INV);
    s.getRange(1, 1, 1, 9).setValues([[
      "盤點ID", "盤點日期", "設備編號", "設備名稱", "序號", "結果", "備註", "盤點人", "登記時間"
    ]]).setFontWeight("bold").setBackground("#E1F1EB");
    s.setFrozenRows(1);
    s.setColumnWidth(4, 200);
    s.setColumnWidth(5, 180);
    s.getRange("B:B").setNumberFormat("@");
    s.getRange("E:E").setNumberFormat("@");
  }
  return s;
}

function invalidateAndReturn() {
  CacheService.getScriptCache().remove(CACHE_KEY);
  var data = readAllFresh();
  try {
    CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(data), CACHE_SEC);
  } catch (err) {}
  return data;
}

/* ---------- 一次性匯入:同仁「設備借用動態總表」的正式設備與編號 ----------
 * 在 Apps Script 編輯器選擇 importEquipment 後按「執行」。做四件事:
 *  1. 「設備」工作表整份換成 46 項正式設備(共 270 個序號)。
 *  2. 「紀錄」工作表清空重建,寫入總表上目前已借出的 29 件(22 筆紀錄)。
 *  3. 建立「盤點」工作表,故障/長期借用的序號登記為停借,
 *     並保留總表既有的 18 筆盤點日(2026-04-23)。
 * 注意:會清除原本的示範設備與測試紀錄;重複執行會重設回總表狀態。
 */
function importEquipment() {
  function seq(prefix, from, to, pad) {
    var out = [];
    for (var i = from; i <= to; i++) out.push(prefix + ("000" + i).slice(-pad));
    return out.join(", ");
  }

  // [編號, 名稱, 分類, 備註, 序號清單]
  var items = [
    ["IPAD", "iPad", "學習載具", "", seq("IPAD-", 1, 60, 2)],
    ["IPDCRC", "平板充電車(準備教室A)", "學習載具", "", "IPDCRC-01"],
    ["MM", "Mac Mini", "電腦設備", "",
      "MM01(114後埔國小採購), MM02(114後埔國小採購), MM03(114後埔國小採購), MM04(114後埔國小採購), MM05(114後埔國小採購), " + seq("數辦ydlo-", 1, 5, 2)],
    ["NB", "筆電", "電腦設備", "",
      seq("數辦", 1, 7, 2) + ", dynabook-A50-K(114後埔國小採購)01, dynabook-A50-K(114後埔國小採購)02, 教卓4, 教卓7, 教卓8, 教卓11, 教卓12, 教卓15, 教卓18, 教卓19"],
    ["MEGA", "手持擴音喇叭(大聲公)", "影音設備", "", "MISC-01-MEGA, MISC-02-MEGA"],
    ["AT", "Apple TV", "影音設備", "", "AT-01"],
    ["PJECR", "投影機", "影音設備", "114後埔國小採購", "PJECR-01(114後埔國小採購)"],
    ["RODE", "RODE 無線麥克風", "影音設備", "", "RODE-01-MIC, RODE-02-MIC, RODE-03-MIC, RODE-04-MIC"],
    ["CAM", "錄影機(DV)", "影音設備", "", "CAM-01"],
    ["CAMTRI", "錄影機DV腳架", "影音設備", "", "CAM-TRI-01"],
    ["OSMO", "手機雲台(DJI OSMO MOBILE 6)", "影音設備", "", "DJI OSMO MOBILE 6-01, DJI OSMO MOBILE 6-02"],
    ["INSTA", "Insta 360 x3 全景相機", "影音設備", "", "Insta 360 x3-01"],
    ["MIRTV", "奇蹟大螢幕", "影音設備", "", "奇蹟大螢幕"],
    ["ADPLR", "廣告機", "影音設備", "", "廣告機01, 廣告機02"],
    ["VR", "VR頭盔", "新興科技", "", "VR-01"],
    ["LRLE", "AI吉他", "新興科技", "114後埔國小採購",
      "LRLE-01(114後埔國小採購), LRLE-02(114後埔國小採購), LRLE-03(114後埔國小採購), LRLE-04(114後埔國小採購), LRLE-05(114後埔國小採購), LRLE-06(114後埔國小採購)"],
    ["AIGR", "AI圍棋機器人", "新興科技", "", "AIGR-01, AIGR-02"],
    ["DJINEO", "空拍機(DJI NEO)", "新興科技", "", seq("DJI NEO-", 1, 5, 2)],
    ["DRONSC", "足球無人機", "新興科技", "", seq("DRONSC-", 1, 20, 2)],
    ["KEBBI", "凱比機器人", "新興科技", "", "Kebbi-37, Kebbi-39, Kebbi-40"],
    ["CZL", "小栗方AI學習機", "新興科技", "", "創造栗-23, 創造栗-24, 創造栗-31, 創造栗-32, 創造栗-37, 創造栗-43, 創造栗-49, 創造栗-56"],
    ["MOUSE", "滑鼠", "週邊配件", "", "MOUSE-01"],
    ["SDMC", "SD記憶卡", "週邊配件", "", "SDMC-01"],
    ["IBMC", "標籤機", "週邊配件", "", "IBMC-01"],
    ["APEPL", "Apple Pencil", "週邊配件", "", seq("APEPL-", 1, 5, 2)],
    ["LOGIKMC", "羅技無線鍵鼠組", "週邊配件", "", seq("LOGIKMC-", 1, 32, 2)],
    ["PRES", "簡報筆", "週邊配件", "", "PRES-01"],
    ["TABSAD", "平板架", "週邊配件", "", seq("TABSAD-", 1, 8, 2)],
    ["YLCFM", "AP(無線基地台)", "網路設備", "", "YLCFM-19, YLCFM-22, YLCFM-59, YLCFM-65, YLCFM-72, YLCFM-75, YLCFM-76"],
    ["FML", "無線基地台充電線+充電頭", "網路設備", "", "FML-54, FML-56, FML-61, FML-77"],
    ["CHG", "充電線", "線材電源", "", "CABLE-CHG-01"],
    ["ACERCHG", "充電線(ACER筆電)", "線材電源", "", "教卓8, 教卓11, 教卓15, 教卓17, 教卓18, 教卓19"],
    ["APTEC", "Apple 20W 充電器(Type C)", "線材電源", "", seq("APTEC20W-", 1, 5, 2)],
    ["APLHG", "Apple Lightning線", "線材電源", "", seq("APLHG-", 1, 4, 2)],
    ["DTC", "雙Type C線", "線材電源", "", seq("DTC-", 1, 5, 2)],
    ["TCH", "觸控螢幕線", "線材電源", "", "CABLE-TCH-01"],
    ["TCHJX", "觸控線(家宣)", "線材電源", "", "家宣觸控線-1, 家宣觸控線-2, 家宣觸控線-3, 家宣觸控線-4"],
    ["HDMI", "HDMI線", "線材電源", "", "家宣HDMI-1, 家宣HDMI-2, 家宣HDMI-3, 家宣HDMI-4, CABLE-HDMI-01, CABLE-HDMI-10"],
    ["ADAPR", "音源轉接頭", "線材電源", "", "ADAPR-01"],
    ["3PTO2P", "3孔轉2孔電源接頭", "線材電源", "", seq("3PTO2P-", 1, 5, 2)],
    ["P6", "6孔延長線(黑色)", "線材電源", "", "CABLE-P6-01, CABLE-P6-02, CABLE-P6-03"],
    ["EXT", "輪座四孔延長線", "線材電源", "", seq("CABLE-EXT-", 1, 5, 2)],
    ["BAT3", "3號電池", "其他", "", "BAT3-01, BAT3-02, BAT3-03"],
    ["BAT4", "4號電池", "其他", "", seq("BAT4-", 1, 4, 2)],
    ["CART", "推車", "其他", "", "大黃【大板車】, 小黑1號, 小黑2號, 小綠, 小白"],
    ["UMB", "雨傘", "其他", "", "UMB-01, UMB-02, UMB-03"]
  ];

  // 總表上目前已借出:[設備編號, 借用人, 借用日, 預計歸還日, 序號…]
  var loans = [
    ["IPAD", "呂其樺", "2026-07-17", "2026-07-20", "IPAD-04, IPAD-05, IPAD-06, IPAD-09, IPAD-10, IPAD-12"],
    ["IPAD", "曾欒閔", "2026-06-08", "2040-12-31", "IPAD-17"],
    ["IPAD", "黃銘麒", "2026-03-06", "2027-01-31", "IPAD-26"],
    ["MM", "呂舒婷", "2026-04-10", "2040-12-31", "MM05(114後埔國小採購)"],
    ["MM", "康芳鈞", "2026-07-09", "2029-07-31", "數辦ydlo-02"],
    ["NB", "戴君櫟", "2026-07-09", "2030-12-31", "數辦01"],
    ["NB", "陳奕瑄", "2026-07-09", "2029-12-31", "數辦05"],
    ["NB", "康芳鈞", "2026-07-09", "2029-07-31", "數辦06"],
    ["NB", "曾欒閔", "2026-02-02", "2063-02-02", "數辦07"],
    ["NB", "劉宏龍", "2026-05-18", "2040-12-31", "dynabook-A50-K(114後埔國小採購)01"],
    ["NB", "康芳鈞", "2026-05-18", "2027-07-31", "dynabook-A50-K(114後埔國小採購)02"],
    ["NB", "陳奕瑄", "2026-02-23", "2031-12-31", "教卓8"],
    ["NB", "黃銘麒", "2026-03-25", "2027-02-01", "教卓18"],
    ["NB", "康芳鈞", "2026-03-26", "2027-01-31", "教卓11"],
    ["NB", "曾欒閔", "2026-03-27", "2099-03-27", "教卓7"],
    ["LRLE", "李學人", "2026-01-28", "2027-03-28", "LRLE-04(114後埔國小採購), LRLE-05(114後埔國小採購), LRLE-06(114後埔國小採購)"],
    ["LOGIKMC", "蒲信宏", "2026-02-09", "2030-01-09", "LOGIKMC-05"],
    ["LOGIKMC", "曾欒閔", "2026-02-05", "2059-02-05", "LOGIKMC-22"],
    ["HDMI", "陳奕瑄", "2026-02-09", "2047-02-09", "CABLE-HDMI-10"],
    ["ACERCHG", "陳奕瑄", "2026-02-23", "2031-12-31", "教卓8"],
    ["ACERCHG", "康芳鈞", "2026-03-26", "2027-01-31", "教卓11"],
    ["ACERCHG", "黃銘麒", "2026-03-25", "2027-02-01", "教卓18"]
  ];

  // 停借序號:[設備編號, 序號, 結果, 備註]
  var badSerials = [
    ["IPAD", "IPAD-01", "異常", "故障"],
    ["IPAD", "IPAD-16", "異常", "故障"],
    ["IPAD", "IPAD-48", "異常", "故障"],
    ["NB", "教卓4", "異常", "故障"],
    ["NB", "教卓19", "異常", "故障"],
    ["DRONSC", "DRONSC-04", "異常", "故障"],
    ["DRONSC", "DRONSC-12", "異常", "故障"],
    ["DRONSC", "DRONSC-16", "異常", "故障"],
    ["MM", "數辦ydlo-03", "長期借用", ""],
    ["MM", "數辦ydlo-05", "長期借用", ""],
    ["LRLE", "LRLE-02(114後埔國小採購)", "長期借用", ""],
    ["ACERCHG", "教卓17", "長期借用", ""]
  ];

  // 總表既有的盤點紀錄(2026-04-23 盤過的 iPad)
  var checked0423 = parseList(
    "IPAD-02, IPAD-03, IPAD-04, IPAD-05, IPAD-06, IPAD-07, IPAD-08, IPAD-09, IPAD-10, " +
    "IPAD-11, IPAD-12, IPAD-13, IPAD-14, IPAD-15, IPAD-18, IPAD-19, IPAD-21, IPAD-40");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nameOf = {};
  items.forEach(function (it) { nameOf[it[0]] = it[1]; });
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  // 1) 設備
  var s1 = ss.getSheetByName(SHEET_ITEMS);
  s1.clearContents();
  s1.getRange("F:F").setNumberFormat("@");
  var rows1 = [["編號", "設備名稱", "分類", "總數", "備註", "序號清單"]];
  items.forEach(function (it) {
    rows1.push([it[0], it[1], it[2], parseList(it[4]).length, it[3], it[4]]);
  });
  s1.getRange(1, 1, rows1.length, 6).setValues(rows1);
  s1.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#E1F1EB");

  // 2) 紀錄(清掉測試資料,寫入目前借出中的)
  var s2 = ss.getSheetByName(SHEET_LOANS);
  s2.clearContents();
  var rows2 = [["紀錄ID", "設備編號", "設備名稱", "數量", "借用人", "班級/處室",
                "借用日期", "使用時段", "預計歸還日", "狀態", "歸還日期", "備註", "登記時間", "序號"]];
  loans.forEach(function (l) {
    var n = parseList(l[4]).length;
    rows2.push(["L" + Utilities.getUuid().slice(0, 8), l[0], nameOf[l[0]], n,
                l[1], "—", l[2], "", l[3], "借用中", "", "自動態總表匯入", now, l[4]]);
  });
  s2.getRange(1, 1, rows2.length, 14).setValues(rows2);
  s2.getRange(1, 1, 1, 14).setFontWeight("bold").setBackground("#E1F1EB");

  // 3) 盤點(既有盤點日 + 停借序號)
  var s3 = ensureInvSheet();
  if (s3.getLastRow() > 1) {
    s3.getRange(2, 1, s3.getLastRow() - 1, 9).clearContent();
  }
  var rows3 = [];
  checked0423.forEach(function (s) {
    rows3.push(["I" + Utilities.getUuid().slice(0, 8), "2026-04-23", "IPAD", "iPad",
                s, "正常", "", "匯入", now]);
  });
  badSerials.forEach(function (b) {
    rows3.push(["I" + Utilities.getUuid().slice(0, 8), today, b[0], nameOf[b[0]],
                b[1], b[2], b[3], "匯入", now]);
  });
  s3.getRange(2, 1, rows3.length, 9).setValues(rows3);

  CacheService.getScriptCache().remove(CACHE_KEY);
}

/* ---------- v1 → v2 升級:替既有工作表補上序號欄 ---------- */
function upgrade() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var s1 = ss.getSheetByName(SHEET_ITEMS);
  if (s1 && String(s1.getRange("F1").getValue()) !== "序號清單") {
    s1.getRange("F1").setValue("序號清單")
      .setFontWeight("bold").setBackground("#E1F1EB");
    s1.setColumnWidth(6, 300);
    s1.getRange("F:F").setNumberFormat("@");
  }

  var s2 = ss.getSheetByName(SHEET_LOANS);
  if (s2 && String(s2.getRange("N1").getValue()) !== "序號") {
    s2.getRange("N1").setValue("序號")
      .setFontWeight("bold").setBackground("#E1F1EB");
    s2.setColumnWidth(14, 220);
    s2.getRange("N:N").setNumberFormat("@");
  }

  ensureInvSheet();
  CacheService.getScriptCache().remove(CACHE_KEY);
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
      ["編號", "設備名稱", "分類", "總數", "備註", "序號清單"],
      ["T01", "iPad 平板(充電車 A)", "學習載具", 30, "含充電車,借用請整車推走", ""],
      ["T02", "Chromebook 筆電", "學習載具", 20, "", ""],
      ["P01", "單槍投影機", "影音設備", 3, "附 HDMI 線", ""],
      ["P02", "實物投影機", "影音設備", 4, "", ""],
      ["A01", "行動擴音音箱", "影音設備", 5, "含頭戴麥克風", ""],
      ["C01", "網路攝影機", "影音設備", 6, "視訊、直播用", ""],
      ["M01", "無線麥克風組", "影音設備", 3, "一組兩支", ""],
      ["V01", "VR 眼鏡", "新興科技", 10, "使用後請酒精擦拭", ""],
      ["R01", "藍牙簡報筆", "週邊配件", 8, "", ""],
      ["H01", "HDMI 訊號線(5m)", "週邊配件", 10, "", ""],
      ["E01", "動力延長線", "週邊配件", 12, "", ""],
      ["S01", "相機三腳架", "週邊配件", 4, "", ""]
    ];
    s1.getRange(1, 1, items.length, 6).setValues(items);
    s1.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#E1F1EB");
    s1.setFrozenRows(1);
    s1.setColumnWidth(2, 220);
    s1.setColumnWidth(5, 260);
    s1.setColumnWidth(6, 300);
    s1.getRange("F:F").setNumberFormat("@");
  }

  if (!ss.getSheetByName(SHEET_LOANS)) {
    var s2 = ss.insertSheet(SHEET_LOANS);
    var head = [["紀錄ID", "設備編號", "設備名稱", "數量", "借用人", "班級/處室",
                 "借用日期", "使用時段", "預計歸還日", "狀態", "歸還日期", "備註", "登記時間", "序號"]];
    s2.getRange(1, 1, 1, 14).setValues(head).setFontWeight("bold").setBackground("#E1F1EB");
    s2.setFrozenRows(1);
    s2.setColumnWidth(3, 200);
    s2.setColumnWidth(14, 220);
    // 日期與序號欄以純文字保存,避免時區換算與格式誤差
    s2.getRange("G:I").setNumberFormat("@");
    s2.getRange("K:K").setNumberFormat("@");
    s2.getRange("N:N").setNumberFormat("@");
  }

  // 刪掉預設的空白工作表(如果還在而且不只一張)
  var def = ss.getSheetByName("工作表1") || ss.getSheetByName("Sheet1");
  if (def && ss.getSheets().length > 2) ss.deleteSheet(def);

  upgrade();
}
