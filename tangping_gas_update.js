// ══════════════════════════════════════════════════════════════════
// 躺平 KL Livehouse — 点歌系统 Google Apps Script (并发安全版)
// 替换你现有的 Apps Script 代码，然后重新部署为 Web App
// ══════════════════════════════════════════════════════════════════

const SHEET_NAME = 'Votes';   // 投票记录 sheet 名称（可改）

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 如果 sheet 不存在则创建
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Timestamp', 'Date', 'Song Title', 'Singers', 'Session TS']);
    }

    const now = new Date().toISOString();
    const date = data.date || '';
    const ts = data.ts || '';

    // 每一首投票歌曲 append 一行（append-row 是原子操作，天然并发安全）
    (data.votes || []).forEach(vote => {
      sheet.appendRow([now, date, vote.title, vote.singer || '', ts]);
    });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// GET 请求：返回今天的投票统计
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return ContentService.createTextOutput('{}').setMimeType(ContentService.MimeType.JSON);

    const today = new Date().toISOString().slice(0, 10);
    const rows = sheet.getDataRange().getValues();
    const counts = {};

    rows.slice(1).forEach(row => {
      const rowDate = String(row[1]).slice(0, 10);
      if (rowDate === today) {
        const title = row[2];
        counts[title] = (counts[title] || 0) + 1;
      }
    });

    return ContentService
      .createTextOutput(JSON.stringify(counts))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
