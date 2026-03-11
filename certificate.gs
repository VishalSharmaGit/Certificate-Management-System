function ULRCertificateNumber() {
  const SPREADSHEET_ID = '11duatGxbZbCf8QkH7RUoiuuoVb1s696LCRZCrcl0094';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Certificate database 2026");

  if (!sheet) {
    throw new Error('Sheet "Certificate database 2026" not found');
  }

  const lastRow = sheet.getLastRow();
  const values = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    : [];

  const count = values.flat().filter(String).length + 1;

  const prefix = "CC247226";
  const numberLength = 9;

  const padded = String(count).padStart(numberLength, "0");
  console.log(prefix + padded + "F");
  return prefix + padded + "F";
}



// function generateAndSaveCertificateNumber() {
//   const lock = LockService.getScriptLock();
//   lock.waitLock(30000); // wait up to 30 seconds

//   try {
//     const sheet = SpreadsheetApp
//       .getActiveSpreadsheet()
//       .getSheetByName("certifi");

//     const lastRow = sheet.getLastRow();
//     const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

//     const count = values.flat().filter(Boolean).length + 1;

//     const prefix = "CC247226";
//     const numberLength = 9;
//     const padded = String(count).padStart(numberLength, "0");
//     const certificateNumber = prefix + padded;

//     // Save the number permanently
//     sheet.appendRow([certificateNumber]);

//     return certificateNumber;

//   } finally {
//     lock.releaseLock();
//   }
// }
