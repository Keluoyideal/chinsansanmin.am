/**
 * 門市早中班檢查表：Google 試算表後端
 *
 * 「資料」工作表：網頁讀寫用，一天一列（請勿手動修改）
 * 「盤點紀錄」工作表：給人看的，一天一列，每個品項一欄
 */

const JUICE = ["水梨","椪柑","甘蔗","蘋果","柳丁","紅柚","檸檬","水蜜桃"];
const COLD  = ["鮮奶","多多","鮮奶油","荔枝","冬露","燕麥奶","糖桶"];

function doGet(e) {
  const date = String((e.parameter || {}).date || "");
  const rows = dataSheet_().getDataRange().getValues().slice(1);
  let day = null, prev = null;
  rows.forEach(r => {
    const d = String(r[0]);
    if (d === date) day = parse_(r);
    else if (d < date && (!prev || d > prev.date)) prev = parse_(r);
  });
  return json_({ day, prev });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const date = String(body.date || "");
    if (!/^\d{6}$/.test(date)) return json_({ ok: false, error: "日期格式錯誤" });
    const now = new Date();
    const row = [date, body.filler || "", JSON.stringify(body.checks || {}),
                 JSON.stringify(body.juice || {}), JSON.stringify(body.cold || {}), now,
                 JSON.stringify(body.adj || {}), body.helpers || "", body.memo || ""];
    upsert_(dataSheet_(), date, row);
    upsert_(reportSheet_(), date, reportRow_(body, now));
    // 每日檢核：一天一列，每一項寫「✓ 填表人 時間」，沒勾留空
    if (Array.isArray(body.checkLog)) {
      const labels = body.checkLog.map(x => x[0]);
      const sh = sheet_("每日檢核", ["日期","主要填表人","共同填表人","完成"].concat(labels));
      sh.getRange(1, 1, 1, 4 + labels.length).setValues([["日期","主要填表人","共同填表人","完成"].concat(labels)]);
      const done = body.checkLog.filter(x => x[1]).length;
      upsert_(sh, date, [date, body.filler || "", body.helpers || "", done + "/" + labels.length].concat(body.checkLog.map(x => x[1])));
    }
    // 每次修改都留一筆，填錯時可以查回原本的數字
    logSheet_().appendRow([now, date, body.filler || "", row[3], row[4]]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function parse_(r) {
  const p = s => { try { return JSON.parse(s || "{}"); } catch (e) { return {}; } };
  return { date: String(r[0]), filler: r[1], checks: p(r[2]), juice: p(r[3]), cold: p(r[4]), adj: p(r[6]), helpers: r[7] || "", memo: r[8] || "" };
}

function upsert_(sheet, date, row) {
  const n = sheet.getLastRow();
  if (n > 1) {
    const dates = sheet.getRange(2, 1, n - 1, 1).getDisplayValues().map(r => r[0]);
    const i = dates.indexOf(date);
    if (i >= 0) { sheet.getRange(i + 2, 1, 1, row.length).setValues([row]); return; }
  }
  sheet.appendRow(row);
}

function reportRow_(b, now) {
  const j = b.juice || {}, c = b.cold || {};
  const done = Object.values(b.checks || {}).filter(Boolean).length;
  return [b.date, [b.filler, b.helpers].filter(String).join("、"), done]
    .concat(JUICE.map(n => (j[n] || []).map(v => v || 0).concat([0,0,0]).slice(0,3).join("/")))
    .concat(COLD.map(n => (c[n] || [])[0] || 0))
    .concat([now, b.memo || ""]);
}

function dataSheet_() {
  return sheet_("資料", ["日期","填表人","檢查","原汁","冷藏常溫","更新時間","手動修正","共同填表人","備註"]);
}
function reportSheet_() {
  return sheet_("盤點紀錄", ["日期","填表人","已勾項目"]
    .concat(JUICE.map(n => n + "（庫存/退冰/進貨）"))
    .concat(COLD).concat(["更新時間","備註"]));
}
function logSheet_() {
  return sheet_("修改紀錄", ["修改時間","盤點日期","填表人","原汁","冷藏常溫"], true);
}
// 從試算表開的 Apps Script 會直接用那份試算表；
// 如果是在 script.google.com 單獨建立的，第一次執行時會自動新建一份試算表並記住它
function ss_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("SHEET_ID");
  if (id) return SpreadsheetApp.openById(id);
  const created = SpreadsheetApp.create("早班中班Checklist 資料");
  props.setProperty("SHEET_ID", created.getId());
  return created;
}

function sheet_(name, header, isLog) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(isLog ? "B:B" : "A:A").setNumberFormat("@"); // 日期當文字存，避免 261011 變成數字
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** 第一次設定時在編輯器執行一次：建立工作表並存入 261011 的盤點 */
function setup() {
  Logger.log("資料存在這份試算表：" + ss_().getUrl());
  doPost({ postData: { contents: JSON.stringify({
    date: "261011", filler: "小唐", checks: {},
    juice: { "水梨":["65","9","0"], "椪柑":["0","1","0"], "甘蔗":["0","0","0"], "蘋果":["18","2","0"],
             "柳丁":["0","0","0"], "紅柚":["4","1","0"], "檸檬":["2","0","0"], "水蜜桃":["19","2","0"] },
    cold:  { "鮮奶":["69"], "多多":["51"], "鮮奶油":["6"], "荔枝":["4"], "冬露":["6"], "燕麥奶":["3"], "糖桶":["4"] }
  }) } });
}
