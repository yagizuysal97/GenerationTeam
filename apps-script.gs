/**
 * Humanising Energy — Jüri Puanlama Backend
 * Bu kodu Google Sheets içinde Uzantılar > Apps Script'e yapıştırın.
 * Kurulum adımları için README.md dosyasına bakın.
 */

const SHEET_NAME = 'Yanitlar';

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const sheet = getOrCreateSheet_();

  sheet.appendRow([
    new Date(),          // sunucu zaman damgası
    data.timestamp || '',
    data.judge || '',
    data.team || '',
    data.A || 0,
    data.B || 0,
    data.C || 0,
    data.D || 0,
    data.E || 0,
    data.total || 0,
    data.note || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const sheet = getOrCreateSheet_();
  const rows = sheet.getDataRange().getValues();
  rows.shift(); // başlık satırını at

  const data = rows
    .filter(r => r[3]) // takım adı boş olan satırları atla
    .map(r => ({
      serverTime: r[0],
      timestamp: r[1],
      judge: r[2],
      team: r[3],
      A: Number(r[4]) || 0,
      B: Number(r[5]) || 0,
      C: Number(r[6]) || 0,
      D: Number(r[7]) || 0,
      E: Number(r[8]) || 0,
      total: Number(r[9]) || 0,
      note: r[10] || ''
    }));

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'Sunucu Zamanı', 'Gönderim Zamanı', 'Jüri', 'Takım',
      'A - Problem&HE (15)', 'B - Teknik (25)', 'C - Uygulanabilirlik&Risk (20)',
      'D - AI Agent (25)', 'E - Sunum&Q&A (15)', 'Toplam (100)', 'Not'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
