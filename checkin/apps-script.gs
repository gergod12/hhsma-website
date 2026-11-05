/**
 * 全人健康自主管理推廣協會 線上簽到 後端
 * 貼到「線上簽到單」試算表的 Apps Script（擴充功能 → Apps Script），
 * 部署為「網頁應用程式」：執行身分＝我、存取權＝所有人。
 * 簽到資料會寫進同一份試算表的「簽到紀錄」工作表。
 */
var TZ = "Asia/Taipei";
var SHEET_NAME = "簽到紀錄";
// 「線上簽到單」試算表的 ID（網址 /spreadsheets/d/ 後面那一串）。
// 有填時可以用獨立的 Apps Script 專案（script.google.com 新增專案）部署，不必從試算表開啟。
var SPREADSHEET_ID = "1k0IFCW2EssIMf5wKi69E9iKjmob08ElawxYoVDRiHog";
var HEADER = ["簽到時間", "會議", "姓名", "職稱", "來源"];

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    if (p.action === "checkin") {
      out = checkin_(p);
    } else {
      out = { ok: true, rows: list_(p.meeting || "") };
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  var json = JSON.stringify(out);
  if (p.callback && /^[\w$.]+$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  var ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight("bold");
  }
  return sh;
}

function list_(meeting) {
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, HEADER.length).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!r[2]) continue;
    if (meeting && String(r[1]) !== String(meeting)) continue;
    rows.push({
      time: fmt_(r[0]),
      meeting: String(r[1]),
      name: String(r[2]),
      role: String(r[3] || ""),
      source: String(r[4] || "")
    });
  }
  return rows;
}

function checkin_(p) {
  var meeting = String(p.meeting || "").trim();
  var name = String(p.name || "").trim();
  var role = String(p.role || "").trim();
  var source = String(p.source || "").trim().slice(0, 40);
  if (!meeting || !name) return { ok: false, error: "缺少會議或姓名" };
  if (name.length > 20 || role.length > 20) return { ok: false, error: "姓名或職稱過長" };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var existing = list_(meeting);
    for (var i = 0; i < existing.length; i++) {
      if (existing[i].name === name) {
        return { ok: true, already: true, time: existing[i].time, name: name, role: existing[i].role };
      }
    }
    var now = new Date();
    sheet_().appendRow([now, meeting, name, role, source]);
    return { ok: true, already: false, time: fmt_(now), name: name, role: role };
  } finally {
    lock.releaseLock();
  }
}

function fmt_(d) {
  if (!(d instanceof Date)) return String(d || "");
  return Utilities.formatDate(d, TZ, "yyyy-MM-dd HH:mm");
}
