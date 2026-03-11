const RECORD_FOLDER_ID = '1RzXbhxbVJDVxZ6l4FugJgHLFA9H5rs9i';
const CERT_FOLDER_ID   = '1hnlN418k0pyVuiMCVkLxx4kV65z2_ARC';
const CERT_TEMPLATE_ID = '1gm9_69LakHRXo-PZpxW_Yn_EnZK1jQK4xOFbqXENvDs';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Record & Certificate Generator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ================= TEMPLATE DATA ================= */

function getTemplatesData() {
  const ss = SpreadsheetApp.openById('1f0Ml-FNPsxJFjuLDwxYzEJbCK2R0SNT9ke3RoOptRk0');
  const sh = ss.getSheetByName('Templates');

  if (!sh) {
    throw new Error('Sheet "Templates" not found');
  }

  // 👇 CRITICAL FIX
  const v = sh.getDataRange().getDisplayValues();

  const out = [];

  for (let i = 1; i < v.length; i++) {
    if (!v[i][0] || !v[i][1]) continue;

    out.push({
      id: v[i][0],
      templateId: v[i][1],
      Standard: v[i][2],
      RecordSheet: v[i][12],
      Error: v[i][13],
      Uncertainity:v[i][14],
      UncSheetLink:v[i][31],

      Master1: v[i][4],
      Master2: v[i][5],
      Master3: v[i][6],
      Master4: v[i][7],

      Master1ID: v[i][8],
      Master2ID: v[i][9],
      Master3ID: v[i][10],
      Master4ID: v[i][11],

      Make1: v[i][15],
      Make2: v[i][16],
      Make3: v[i][17],
      Make4: v[i][18],

      ULRIID: v[i][19],
      ULR2ID: v[i][20],
      ULR3ID: v[i][21],
      ULR4ID: v[i][22],

      DueDateI: v[i][23],
      DueDate2: v[i][24],
      DueDate3: v[i][25],
      DueDate4: v[i][26],

      CalibI: v[i][27],
      Calib2: v[i][28],
      Calib3: v[i][29],
      Calib4: v[i][30]
    });
  }

  return out;
}


/* ================= FMS ================= */

const SPREADSHEET_ID = '1f0Ml-FNPsxJFjuLDwxYzEJbCK2R0SNT9ke3RoOptRk0';

function getSRFList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('FMS');

  if (!sh) {
    throw new Error('Sheet "FMS" not found in the spreadsheet.');
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  return sh
    .getRange(2, 2, lastRow - 1, 1)   // FMS!ColB
    .getDisplayValues()               // formula-safe
    .flat()
    .map(v => v.trim())
    .filter(v => v !== "")
    .reverse();
}


function fmt(d) {
  return d instanceof Date
    ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy')
    : d;
}



function getFMSBySRF(srf) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('FMS');
  const v = sh.getDataRange().getValues();

  for (let i = 1; i < v.length; i++) {
    if (String(v[i][1]).trim() === srf) {
      return {
        username: v[i][2],
        receipt: fmt(v[i][4]),
        recommendDue: fmt(v[i][22])
      };
    }
  }
  throw new Error('SRF not found in FMS');
}

/* ================= DOCUMENT CREATOR ================= */

function createDoc(templateId, name, folderId, replacements) {
  const file = DriveApp.getFileById(templateId)
    .makeCopy(name, DriveApp.getFolderById(folderId));

  const doc = DocumentApp.openById(file.getId());
  replaceAll(doc.getBody(), replacements);
  if (doc.getHeader()) replaceAll(doc.getHeader(), replacements);
  if (doc.getFooter()) replaceAll(doc.getFooter(), replacements);
  doc.saveAndClose();

  return {id: file.getId(), name: file.getName(), url: file.getUrl() };
}

function replaceAll(el, reps) {
  for (let k in reps) {
    while (el.findText(`{{${k}}}`)) {
      el.replaceText(`{{${k}}}`, reps[k] || '');
    }
  }
}

/* ================= MAIN ================= */

function generateAll(payload) {
  const fms = getFMSBySRF(payload.srf);
  const t = payload.template;

  const base = {
    'SRF No': payload.srf,
    'ID':payload.ID,
    'ULR No': payload.ulr,
    'Date of Calibration': payload.date,
    'issuereport': payload.date,
    'RecommendDueDate': fms.recommendDue,
    'username&address': fms.username,
    'receipt': fms.receipt,
    'Make': payload.make,
    'Range/Size': payload.range,
    'Description': payload.desc,
    'Start Temp': payload.startTemp,
    'End Temp': payload.endTemp,
    'RH Start': payload.rhStart,
    'RH End': payload.rhEnd,
    'Calibrated Person': payload.calibratedPerson,
    'Calibrated desg': payload.calibratedDesg,
    'Standard': t.Standard,
    'RecordSheet': t.RecordSheet,
    'Master1': t.Master1,
    'Master2': t.Master2,
    'Master3': t.Master3,
    'Master4': t.Master4,
    'Master1ID': t.Master1ID,
    'Master2ID': t.Master2ID,
    'Master3ID': t.Master3ID,
    'Master4ID': t.Master4ID,
    'Make1': t.Make1,
    'Make2': t.Make2,
    'Make3': t.Make3,
    'Make4': t.Make4,
    'ULRIID': t.ULRIID,
    'ULR2ID': t.ULR2ID,
    'ULR3ID': t.ULR3ID,
    'ULR4ID': t.ULR4ID,
    'DueDateI': t.DueDateI,
    'DueDate2': t.DueDate2,
    'DueDate3': t.DueDate3,
    'DueDate4': t.DueDate4,
    'CalibI': t.CalibI,
    'Calib2': t.Calib2,
    'Calib3': t.Calib3,
    'Calib4': t.Calib4,
    'Error' : t.Error,
    'Uncertainity' : t.Uncertainity,
    'UncSheetLink' : t.UncSheetLink

  };

  const record = createDoc(
    t.templateId,
    `RecordSheet_${payload.srf}`,
    RECORD_FOLDER_ID,
    base
  );

  const cert = createDoc(
    CERT_TEMPLATE_ID,
    `Certificate - ${payload.ulr}`,
    CERT_FOLDER_ID,
    base
  );
  generateCalibrationTables({
    recordDocId: record.id,
    certificateDocId: cert.id,
    calibrationResults: payload.calibrationResults,
    repeatabilityResults: payload.repeatabilityResults
  });
  // 🔥 ADD THIS LINE
  appendToCertificateDatabase(payload, record, cert);
  return { record, cert };
}

// ==============For Certificate Database Pop up Table ================
function getCertificateDatabase() {
  const ss = SpreadsheetApp.openById("11duatGxbZbCf8QkH7RUoiuuoVb1s696LCRZCrcl0094")
  const sheet = ss.getSheetByName('Certificate database 2026');

  if (!sheet) throw new Error('Sheet not found');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const ulrCol = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getDisplayValues()
    .flat()
    .filter(v => v.trim() !== '');

  const rowCount = ulrCol.length;
  if (rowCount === 0) return [];

  const values = sheet.getRange(2, 1, rowCount, 10).getDisplayValues();
  const links  = sheet.getRange(2, 1, rowCount, 10).getRichTextValues();

  return values.map((r, i) => ({
    ulr: r[0],
    srf: r[1],
    issue: r[3],
    due: r[5],
    record: links[i][6]?.getLinkUrl() || '',
    pdf: links[i][7]?.getLinkUrl() || '',
    word: links[i][9]?.getLinkUrl() || ''
  })).reverse();
}

// ============get the link direct to the sheet
function appendToCertificateDatabase(payload, record, cert) {
  const ss = SpreadsheetApp.openById("11duatGxbZbCf8QkH7RUoiuuoVb1s696LCRZCrcl0094")
  const sh = ss.getSheetByName('Certificate database 2026');

  if (!sh) throw new Error('Certificate database 2026 sheet not found');

  // Find next row
  const lastDataRow = sh
    .getRange(2, 1, sh.getLastRow())
    .getValues()
    .flat()
    .filter(String)
    .length;

  const nextRow = lastDataRow + 2;

  // ---- BASIC DATA ----
  sh.getRange(nextRow, 1).setValue(payload.ulr);   // Column A
  sh.getRange(nextRow, 2).setValue(payload.srf);   // Column B

  // ✅ ADD THIS LINE
  sh.getRange(nextRow, 4).setValue(payload.date);  // Column D → Date of Calibration

  // ---- LINKS ----
  sh.getRange(nextRow, 7).setRichTextValue(
    SpreadsheetApp.newRichTextValue()
      .setText('Open Record')
      .setLinkUrl(record.url)
      .build()
  );

  sh.getRange(nextRow, 10).setRichTextValue(
    SpreadsheetApp.newRichTextValue()
      .setText('Open Certificate')
      .setLinkUrl(cert.url)
      .build()
  );
}

// =======================
function generateCalibrationTables(payload) {
  const recordDoc = DocumentApp.openById(payload.recordDocId);
  const certDoc = DocumentApp.openById(payload.certificateDocId);

  insertCalibrationTable(
    recordDoc.getBody(),
    payload.calibrationResults,
    'record'
  );

  insertCalibrationTable(
    certDoc.getBody(),
    payload.calibrationResults,
    'certificate'
  );

  // 🔥 ADD THIS
  insertRepeatabilityTable(
    recordDoc.getBody(),
    payload.repeatabilityResults
  );

  recordDoc.saveAndClose();
  certDoc.saveAndClose();
}

function insertCalibrationTable(body, data, type) {
  const found = body.findText('{{CALIBRATION_RESULTS}}');
  if (!found) throw new Error('Placeholder {{CALIBRATION_RESULTS}} not found');

  const el = found.getElement().asText();
  const start = found.getStartOffset();
  const end = found.getEndOffsetInclusive();
  el.deleteText(start, end);

  const parent = el.getParent();
  const index = body.getChildIndex(parent);

  // Remove placeholder paragraph completely
  body.removeChild(parent);

  const table = body.insertTable(index);

  if (type === 'record') {
    buildRecordTable(table, data);
  } else {
    buildCertificateTable(table, data);
  }

  styleTable(table, type);
}



function buildRecordTable(table, data) {
  const headers = [
    'S.No',
    'Accuracy Check',
    'Nominal Value',
    'Measured I',
    'Measured II',
    'Measured III',
    'Avg. Measured Value',
    'Error'
  ];

  const headerRow = table.appendTableRow();
  headers.forEach(h => headerRow.appendTableCell(h).setBold(true));

  data.forEach((r, i) => {
    const row = table.appendTableRow();
    row.appendTableCell(String(i + 1));
    row.appendTableCell(r.accuracy || '');
    row.appendTableCell(r.nominal);
    row.appendTableCell(r.m1);
    row.appendTableCell(r.m2);
    row.appendTableCell(r.m3);
    row.appendTableCell(r.avg);
    row.appendTableCell(r.error);
  });
  // ✅ Set Column Widths
  setRecordColumnWidths(table);

  // ✅ Apply Style
  styleTable(table);
}


function buildCertificateTable(table, data) {
  const headers = [
    'S.No',
    'Accuracy Check',
    'Nominal Value',
    'Avg. Measured Value',
    'Error'
  ];

  const headerRow = table.appendTableRow();
  headers.forEach(h => headerRow.appendTableCell(h).setBold(true));

  data.forEach((r, i) => {
    const row = table.appendTableRow();
    row.appendTableCell(String(i + 1));
    row.appendTableCell(r.accuracy);   // Accuracy Check
    row.appendTableCell(r.nominal);
    row.appendTableCell(r.avg);
    row.appendTableCell(r.error);
  });
  setCertificateColumnWidths(table);
  styleTable(table);
}

// Column Width Function
function setRecordColumnWidths(table) {

  const widths = [40, 80, 80, 65, 65, 65, 80, 50];

  for (let i = 0; i < widths.length; i++) {
    table.setColumnWidth(i, widths[i]);
  }
}

function setCertificateColumnWidths(table) {

  const widths = [70, 150, 150, 100];

  for (let i = 0; i < widths.length; i++) {
    table.setColumnWidth(i, widths[i]);
  }
}


function styleTable(table, type) {

  table.setBorderWidth(1);

  const rows = table.getNumRows();

  for (let r = 0; r < rows; r++) {
    const row = table.getRow(r);
    const cols = row.getNumCells();

    for (let c = 0; c < cols; c++) {

      const cell = row.getCell(c);

      // Remove any previous bold formatting
      const text = cell.editAsText();
      text.setBold(false);

      const para = cell.getChild(0).asParagraph();

      para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      para.setFontSize(10);

      cell.setPaddingTop(3);
      cell.setPaddingBottom(3);
      cell.setPaddingLeft(4);
      cell.setPaddingRight(4);

      // ONLY HEADER ROW BOLD
      if (r === 0) {
        text.setBold(true);
      }
    }
  }
}

// 
function insertRepeatabilityTable(body, data) {

  const found = body.findText('{{Measured_RESULTS}}');
  if (!found) return;

  const textElement = found.getElement().asText();
  const start = found.getStartOffset();
  const end = found.getEndOffsetInclusive();

  // Remove only the placeholder text
  textElement.deleteText(start, end);

  const paragraph = textElement.getParent();
  const parent = paragraph.getParent();
  const index = parent.getChildIndex(paragraph);

  // Insert table AFTER the placeholder paragraph
  const table = parent.insertTable(index + 1);

  const header = table.appendTableRow();
  header.appendTableCell('S.No').setBold(true);
  header.appendTableCell('Measured value').setBold(true);

  data.forEach(r => {
    const row = table.appendTableRow();
    row.appendTableCell(String(r.sno));
    row.appendTableCell(r.value);
  });

}
