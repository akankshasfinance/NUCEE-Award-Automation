function generateAlphaFundDocs() {
  // ---- CONFIGURATION ----
  const SPREADSHEET_ID = '1Zu2yqUdz_us_qP6RTcxxLMJupcq1jVIpS7dIn_TvaSs';
  const SHEET_NAME = 'Form Responses 1'; 
  const INVOICE_TEMPLATE_ID = '121zecm2OEqYo3QDHJmcwRQpzRLimvd_vQLIYCsxc2So';
  const AWARD_TEMPLATE_ID = '1JMj54sCDtSk10EXDzje_VVQHsgI_ji3aIpZBTz7IOtE';
  const DESTINATION_FOLDER_ID = '1lsHMtLLP219mRAAaeE8E08RIoYHejxdw';
  const SCRIPT_PROP_KEY = 'lastProcessedRow';
  
  // ---- FETCH DATA ----
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  const props = PropertiesService.getScriptProperties();
  let lastProcessedRow = parseInt(props.getProperty(SCRIPT_PROP_KEY)) || 1;
  if (lastProcessedRow >= data.length) {
    Logger.log('No new responses to process.');
    return;
  }

  for (let i = lastProcessedRow; i < data.length; i++) {
    const row = data[i];

    // ---- SAFELY EXTRACT RESPONSES ----
    const legalName = (row[2] && typeof row[2] === "string") ? row[2] : "";
    const ventureName = row[3] ? row[3].toString() : "Unknown Venture";
    const mailingAddress = row[5] || "";
    const poNumber = row[6] || "";
    const invoiceNumber = row[7] || "";
    const awardAmount = row[8] || "";
    const awardTerm = row[9] || "Term";
    const awardYear = row[10] || "Year";
    const createdBy = row[11] || "";  //
    const firstName = legalName ? legalName.trim().split(" ")[0] : "";
    const semesterYear = `${awardTerm} ${awardYear}`;

    // ---- CREATE FOLDER ----
    const parentFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
    const folderName = `${ventureName} - ${awardTerm}, ${awardYear}`;
    const newFolder = parentFolder.createFolder(folderName);

    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    // ---- GENERATE INVOICE DOC ----
    const invoiceCopy = DriveApp.getFileById(INVOICE_TEMPLATE_ID)
      .makeCopy(`${today}_Alpha_${ventureName}_Invoice`, newFolder);
    const invoiceDoc = DocumentApp.openById(invoiceCopy.getId());
    const invoiceBody = invoiceDoc.getBody();

    invoiceBody.replaceText('{{DATE}}', today);
    invoiceBody.replaceText('{{COMPANY_NAME}}', ventureName);
    invoiceBody.replaceText('{{RECIPIENT_NAME}}', legalName);
    invoiceBody.replaceText('{{MAILING_ADDRESS}}', mailingAddress);
    invoiceBody.replaceText('{{PURCHASE_ORDER}}', poNumber);
    invoiceBody.replaceText('{{INVOICE_NUMBER}}', invoiceNumber);
    invoiceBody.replaceText('{{SEMESTER_YEAR}}', semesterYear);
    invoiceBody.replaceText('{{AWARD_AMOUNT}}', awardAmount);
    invoiceBody.replaceText('{{CREATED_BY}}', createdBy);

    invoiceDoc.saveAndClose();

    // ---- EXPORT INVOICE AS PDF ----
    const invoicePdf = DriveApp.getFileById(invoiceCopy.getId())
      .getAs(MimeType.PDF)
      .setName(`${today}_Alpha_${ventureName}_Invoice.pdf`);
    newFolder.createFile(invoicePdf);

    // ---- GENERATE AWARD LETTER DOC ----
    const awardCopy = DriveApp.getFileById(AWARD_TEMPLATE_ID)
      .makeCopy(`${today}_Alpha_${ventureName}`, newFolder);
    const awardDoc = DocumentApp.openById(awardCopy.getId());
    const awardBody = awardDoc.getBody();

    awardBody.replaceText('{{DATE}}', today);
    awardBody.replaceText('{{COMPANY_NAME}}', ventureName);
    awardBody.replaceText('{{RECIPIENT_NAME}}', legalName);
    awardBody.replaceText('{{MAILING_ADDRESS}}', mailingAddress);
    awardBody.replaceText('{{FIRST_NAME}}', firstName);
    awardBody.replaceText('{{AWARD_AMOUNT}}', awardAmount);
    awardBody.replaceText('{{SEMESTER_YEAR}}', semesterYear);
    awardDoc.saveAndClose();

    // ---- EXPORT AWARD LETTER AS PDF ----
    const awardPdf = DriveApp.getFileById(awardCopy.getId())
      .getAs(MimeType.PDF)
      .setName(`${today}_Alpha_${ventureName}.pdf`);
    newFolder.createFile(awardPdf);
  }

  props.setProperty(SCRIPT_PROP_KEY, data.length);
  Logger.log(`Processed ${data.length - lastProcessedRow} new response(s).`);
}
