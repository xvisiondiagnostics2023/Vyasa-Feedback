// ─────────────────────────────────────────────────────────────
// VYASA DIAGNOSTICS — Patient Feedback → Google Sheets
// Paste this entire script into Google Apps Script (script.google.com)
// Then click Deploy → New deployment → Web App
// ─────────────────────────────────────────────────────────────

const SHEET_NAME = 'Feedback';

const HEADERS = [
  'Timestamp (IST)',
  'Visit Date',
  'Scan Type',
  'Referred By',
  '★ Overall',
  '★ Reception',
  '★ Wait Time',
  '★ Staff',
  '★ Comfort',
  '★ Doctor',
  '★ Report Speed',
  '★ Cleanliness',
  'Avg Rating',
  'Explained Procedure?',
  'Privacy Maintained?',
  'Would Recommend?',
  'What We Did Well',
  'What To Improve',
  'Patient Name',
  'Contact',
  'Premium Service Interest'
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Create sheet and headers if first time
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      const headerRow = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRow.setValues([HEADERS]);
      headerRow.setFontWeight('bold');
      headerRow.setBackground('#1B2A4A');
      headerRow.setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 160);
      sheet.setColumnWidth(2, 100);
      sheet.setColumnWidth(3, 180);
      sheet.setColumnWidth(4, 160);
      for (let i = 5; i <= 13; i++) sheet.setColumnWidth(i, 90);
      sheet.setColumnWidth(14, 160);
      sheet.setColumnWidth(15, 140);
      sheet.setColumnWidth(16, 150);
      sheet.setColumnWidth(17, 240);
      sheet.setColumnWidth(18, 240);
      sheet.setColumnWidth(19, 140);
      sheet.setColumnWidth(20, 160);
      sheet.setColumnWidth(21, 180);
    }

    const ratings = [
      data.overall, data.reception, data.wait,
      data.staff, data.comfort, data.doctor,
      data.report, data.clean
    ].filter(v => v !== '' && v != null).map(Number);

    const avg = ratings.length > 0
      ? (ratings.reduce((a, b) => a + b, 0) / ratings.length).toFixed(1)
      : '';

    const row = [
      data.timestamp,
      data.visitDate,
      data.scanType,
      data.refBy,
      data.overall,
      data.reception,
      data.wait,
      data.staff,
      data.comfort,
      data.doctor,
      data.report,
      data.clean,
      avg,
      data.explained,
      data.privacy,
      data.recommend,
      data.good,
      data.improve,
      data.patientName,
      data.contact,
      data.premium
    ];

    sheet.appendRow(row);

    // Colour-code the avg rating cell
    const lastRow = sheet.getLastRow();
    const avgCell = sheet.getRange(lastRow, 13);
    const avgVal = parseFloat(avg);
    if (avgVal >= 4.5) avgCell.setBackground('#D6F5E3');
    else if (avgVal >= 3.5) avgCell.setBackground('#FFF9E6');
    else if (avgVal > 0) avgCell.setBackground('#FDDEDE');

    // Colour-code recommend column
    const recCell = sheet.getRange(lastRow, 16);
    if (data.recommend === 'Yes, definitely') recCell.setBackground('#D6F5E3');
    else if (data.recommend === 'No') recCell.setBackground('#FDDEDE');

    return ContentService
      .createTextOutput(JSON.stringify({status: 'success'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Test this by running doGet — lets you verify the script is deployed correctly
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({status: 'Vyasa feedback script is live'}))
    .setMimeType(ContentService.MimeType.JSON);
}
