function generateGapFundDocs() {
  // ---- CONFIGURATION ----
  const SPREADSHEET_ID = '1_RLjZRsfH0Dwfkfbuk7rZlNE48DleuvjlsNwmtjwiFg';  // GAP Fund responses sheet
  const SHEET_NAME = 'Form Responses 1'; 
  const INVOICE_TEMPLATE_ID = '1MPpTsfvCEz_UMDUx-v5-MyXIOu44RiQ9V0y-YDb0Ah8'; // GAP Fund Invoice template
  const AWARD_TEMPLATE_ID = '1d_1Rw2l_PrdVceSl8-N6MwktJ6agL6g6cuTxpob1sBs';   // GAP Fund Award Letter template
  const DESTINATION_FOLDER_ID = '1cvoDUPuf8aRAhfqMT2WQ3toepoM0-CKJ';          // GAP Fund Generated Docs folder
  
  // ---- FETCH DATA ----
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  // ---- Ensure "Processed" column exists ----
  let header = data[0];
  let processedColIndex = header.indexOf("Processed");
  if (processedColIndex === -1) {
    processedColIndex = header.length;
    sheet.getRange(1, processedColIndex + 1).setValue("Processed");
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Skip if already processed
    if (row[processedColIndex] && row[processedColIndex].toString().toLowerCase() === "yes") {
      continue;
    }

    // ---- SAFELY EXTRACT RESPONSES ----
    const legalName = row[2] ? row[2].toString() : "";
    const ventureName = row[3] ? row[3].toString() : "Unknown Venture";
    const mailingAddress = row[5] ? row[5].toString() : "";
    const poNumber = row[6] ? row[6].toString() : "";
    const invoiceNumber = row[7] ? row[7].toString() : "";
    const awardAmount = row[8] ? row[8].toString() : "";
    const awardMonth = row[9] ? row[9].toString() : "Month";   // ✅ Award Cycle Month
    const awardYear = row[10] ? row[10].toString() : "Year";  // ✅ Award Cycle Year
    const createdBy = row[11] ? row[11].toString() : "";

    const firstName = legalName ? legalName.trim().split(" ")[0] : "";
    const monthYear = `${awardMonth} ${awardYear}`;  // ✅ Month + Year string

    // ---- CREATE FOLDER ----
    const parentFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    const folderName = `${ventureName} - ${monthYear}`;
    const newFolder = parentFolder.createFolder(folderName);

    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    // ---- GENERATE INVOICE DOC ----
    const invoiceCopy = DriveApp.getFileById(INVOICE_TEMPLATE_ID)
      .makeCopy(`${today}_Gap_${ventureName}_Invoice`, newFolder);
    const invoiceDoc = DocumentApp.openById(invoiceCopy.getId());
    const invoiceBody = invoiceDoc.getBody();

    invoiceBody.replaceText('{{DATE}}', today);
    invoiceBody.replaceText('{{COMPANY_NAME}}', ventureName);
    invoiceBody.replaceText('{{RECIPIENT_NAME}}', legalName);
    invoiceBody.replaceText('{{MAILING_ADDRESS}}', mailingAddress);
    invoiceBody.replaceText('{{PURCHASE_ORDER}}', poNumber);
    invoiceBody.replaceText('{{INVOICE_NUMBER}}', invoiceNumber);
    invoiceBody.replaceText('{{MONTH_YEAR}}', monthYear);
    invoiceBody.replaceText('{{AWARD_AMOUNT}}', awardAmount);
    invoiceBody.replaceText('{{CREATED_BY}}', createdBy);

    invoiceDoc.saveAndClose();

    // ---- EXPORT INVOICE AS PDF ----
    const invoicePdf = DriveApp.getFileById(invoiceCopy.getId())
      .getAs(MimeType.PDF)
      .setName(`${today}_Gap_${ventureName}_Invoice.pdf`);
    newFolder.createFile(invoicePdf);

    // ---- GENERATE AWARD LETTER DOC ----
    const awardCopy = DriveApp.getFileById(AWARD_TEMPLATE_ID)
      .makeCopy(`${today}_Gap_${ventureName}`, newFolder);
    const awardDoc = DocumentApp.openById(awardCopy.getId());
    const awardBody = awardDoc.getBody();

    awardBody.replaceText('{{DATE}}', today);
    awardBody.replaceText('{{COMPANY_NAME}}', ventureName);
    awardBody.replaceText('{{RECIPIENT_NAME}}', legalName);
    awardBody.replaceText('{{MAILING_ADDRESS}}', mailingAddress);
    awardBody.replaceText('{{FIRST_NAME}}', firstName);
    awardBody.replaceText('{{AWARD_AMOUNT}}', awardAmount);
    awardBody.replaceText('{{MONTH_YEAR}}', monthYear);
    awardDoc.saveAndClose();

    // ---- EXPORT AWARD LETTER AS PDF ----
    const awardPdf = DriveApp.getFileById(awardCopy.getId())
      .getAs(MimeType.PDF)
      .setName(`${today}_Gap_${ventureName}.pdf`);
    newFolder.createFile(awardPdf);

    // ---- MARK AS PROCESSED ----
    sheet.getRange(i + 1, processedColIndex + 1).setValue("Yes");
  }

  Logger.log("Processed all new GAP Fund responses and updated sheet.");
}
