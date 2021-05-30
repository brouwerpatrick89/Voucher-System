function onOpen() {
  // create custom menu
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    { name: 'New Voucher', functionName: 'newVoucher' },
    { name: 'Send Voucher', functionName: 'sendVoucher' },
    { name: 'Reset Counter', functionName: 'resetCounter' }
  ];
  spreadsheet.addMenu('My Menu', menuItems);
}

function resetCounter() {
  // reset counter used in voucher id generation
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('VOUCHER_ID', '0');

  // show alert when finished
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Counter is reset');
}

function newVoucher() {
  // duplicate template sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Template');
  var newSheet = sheet.copyTo(ss);

  // generate voucher id
  var formattedDate = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd");
  var scriptProperties = PropertiesService.getScriptProperties();
  var voucherID = scriptProperties.getProperty('VOUCHER_ID');
  var voucherID = +voucherID; //convert string to number
  var voucherID = voucherID + 1; //increment id by 1
  scriptProperties.setProperty('VOUCHER_ID', voucherID); //store new id in property

  if (voucherID < 10) {
    voucherID.toString;
    var newVoucherID = '00' + voucherID;
  } else if (voucherID < 100) {
    voucherID.toString;
    var newVoucherID = '0' + voucherID;
  } else {
    voucherID.toString;
    var newVoucherID = voucherID;
  }

  var finalVoucherID = formattedDate + newVoucherID;

  // rename duplicated sheet with voucher id & set as active sheet
  SpreadsheetApp.flush();
  newSheet.setName(finalVoucherID);
  ss.setActiveSheet(newSheet);

  // insert voucher ID into cell I30
  var voucherIDinputRange = ss.getActiveSheet().getRange("I30");
  voucherIDinputRange.setValue("Green_Camp_Voucher" + finalVoucherID);

}

function sendVoucher() {
  // insert info into reminder sheet section
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  // get details from sheet
  var contactName = sheet.getRange("D18").getValues().toString();
  var contactEmail = sheet.getRange("D19").getValues().toString();
  var dateIssued = Utilities.formatDate(sheet.getRange("D21").getValue(), "GMT+8", "MM/dd/yyyy");
  var expireDate = Utilities.formatDate(sheet.getRange("D22").getValue(), "GMT+8", "dd MMMM yyyy");
  var voucherID = sheet.getSheetName().toString();
  var inputObject = [contactName, contactEmail, voucherID, dateIssued];

  // get ranges & create formula object
  var reminderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('EmailData');
  var lastRow = reminderSheet.getLastRow() + 1;
  var inputRange = reminderSheet.getRange("A" + lastRow + ":D" + lastRow);
  var checkboxRange = reminderSheet.getRange("E" + lastRow);
  var formulaRange = reminderSheet.getRange("F" + lastRow + ":J" + lastRow);
  var formulaObject = ["=EDATE($D" + lastRow + ", 12)", "=EDATE($D" + lastRow + ", 18)", "=EDATE($D" + lastRow + ", 21)", "=EDATE($D" + lastRow + ", 23)", "=EDATE($D" + lastRow + ", 24)"];

  // insert details & formulas in reminder sheet
  inputRange.setValues([inputObject]);
  checkboxRange.insertCheckboxes();
  formulaRange.setFormulas([formulaObject]);

  // create pdf section
  var gid = sheet.getSheetId().toString();
  var url = 'https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/export?';

  var exportOptions =
    'exportFormat=pdf&format=pdf' +
    '&size=A4' +
    '&portrait=false' +
    '&fitw=true' + // fit to page width
    '&sheetnames=false&printtitle=false' + // hide optional headers and footers
    '&pagenumbers=false&gridlines=false' + // hide page numbers and gridlines
    '&fzr=false' + // do not repeat row headers (frozen rows) on each page
    '&gid=' + gid; // sheet ID

  var params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  
  // generate the pdf
  var response = UrlFetchApp.fetch(url + exportOptions, params).getBlob();

  // send the pdf as attachement 
  var subject = "Your voucher is here!";
  var htmlTemplate = HtmlService.createTemplateFromFile("Template.html"); // Get email template
  htmlTemplate.contactName = contactName;
  htmlTemplate.body = 1;
  htmlTemplate.expireDate = expireDate;
  var htmlBody = htmlTemplate.evaluate().getContent();

  MailApp.sendEmail({
    to: contactEmail,
    subject: subject,
    htmlBody: htmlBody,
    replyTo: "YOUREMAIL@EMAIL.COM",
    name: "YOUR NAME",
    attachments: [{
            fileName: "Voucher" + voucherID + ".pdf",
            content: response.getBytes(),
            mimeType: "application/pdf"
        }]
  });

  // save the pdf to Drive
  var nameFile = "Voucher" + sheet.getSheetName().toString() + ".pdf"
  var folder = DriveApp.getFolderById('YOUR_FOLDER_ID');
  folder.createFile(response.setName(nameFile));

  // delete voucher sheet
  SpreadsheetApp.flush();
  ss.deleteActiveSheet();
  ss.setActiveSheet(ss.getSheetByName('Template'));

  // show alert when finished
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Voucher is send');

}
