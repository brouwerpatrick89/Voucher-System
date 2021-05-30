function sendReminder() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("EmailData");
  var columnC = sheet.getRange("C2:C").getValues();
  var numRows = columnC.filter(String).length;
  var startRow = 2;
  var data = sheet.getRange(startRow, 1, numRows, 10).getValues();
  var emailSent = "Email Sent";
  var dateToday = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy");

  for (var i = 0; i < data.length; ++i) { // Get data from sheet per row
    var column = data[i];
    var contactName = column[0];
    var emailAddress = column[1];
    var redeemed = column[3];
    if (column[4] != emailSent) { // If cell does not contain 'Email Sent' set variable to date
      var oneYear = Utilities.formatDate(column[4], "GMT+8", "dd/MM/yyyy");
    }
    if (column[4] == emailSent) { // If cell does contain 'Email Sent' set variable to 'Email Sent'
      var oneYear = column[4];
    }
    if (column[5] != emailSent) { // If cell does not contain 'Email Sent' set variable to date
      var sixMonths = Utilities.formatDate(column[5], "GMT+8", "dd/MM/yyyy");
    }
    if (column[5] == emailSent) { // If cell does contain 'Email Sent' set variable to 'Email Sent'
      var sixMonths = column[5];
    }
    if (column[6] != emailSent) { // If cell does not contain 'Email Sent' set variable to date
      var threeMonths = Utilities.formatDate(column[6], "GMT+8", "dd/MM/yyyy");
    }
    if (column[6] == emailSent) { // If cell does contain 'Email Sent' set variable to 'Email Sent'
      var threeMonths = column[6];
    }
    if (column[7] != emailSent) { // If cell does not contain 'Email Sent' set variable to date
      var oneMonth = Utilities.formatDate(column[7], "GMT+8", "dd/MM/yyyy");
    }
    if (column[7] == emailSent) { // If cell does contain 'Email Sent' set variable to 'Email Sent'
      var oneMonth = column[7];
    }
    if (column[8] != emailSent) { // If cell does not contain 'Email Sent' set variable to date
      var expired = Utilities.formatDate(column[8], "GMT+8", "dd/MM/yyyy");
      var expireDate = Utilities.formatDate(column[8], "GMT+8", "dd MMMM yyyy");
    }
    if (column[8] == emailSent) { // If cell does contain 'Email Sent' set variable to 'Email Sent'
      var expired = column[8];
    }
 
    var htmlTemplate = HtmlService.createTemplateFromFile("Template.html"); // Get email template
    htmlTemplate.contactName = contactName;
    htmlTemplate.body = 2;
    var subject = "Your Voucher Validity"; // Set email Subject
    var timeFrame = "";
    
    if (oneYear == dateToday && oneYear !== emailSent && redeemed !== "TRUE") { // Check if 1 year date equals today's date and if email has been sent or voucher has been redeemed
      htmlTemplate.timeFrame = "1 year"; // Set timeFrame variable to be used in the HTML template
      htmlTemplate.expireDate = expireDate; // Set expireDate variable to be used in the HTML template
      sheet.getRange(startRow + i, 5).setValue(emailSent); // Input 'Email Sent' into cell
    }
    else if (sixMonths == dateToday && sixMonths !== emailSent && redeemed !== "TRUE") { // Check if 6 months date equals today's date and if email has been sent or voucher has been redeemed
      htmlTemplate.timeFrame = "6 months"; // Set timeFrame variable to be used in the HTML template
      htmlTemplate.expireDate = expireDate; // Set expireDate variable to be used in the HTML template
      sheet.getRange(startRow + i, 6).setValue(emailSent); // Input 'Email Sent' into cell
    }
    else if (threeMonths == dateToday && threeMonths !== emailSent && redeemed !== "TRUE") { // Check if 3 months date equals today's date and if email has been sent or voucher has been redeemed
      htmlTemplate.timeFrame = "3 months"; // Set timeFrame variable to be used in the HTML template
      htmlTemplate.expireDate = expireDate; // Set expireDate variable to be used in the HTML template
      sheet.getRange(startRow + i, 7).setValue(emailSent); // Input 'Email Sent' into cell
    }
    else if (oneMonth == dateToday && oneMonth !== emailSent && redeemed !== "TRUE") { // Check if 1 month date equals today's date and if email has been sent or voucher has been redeemed
      htmlTemplate.timeFrame = "1 month"; // Set timeFrame variable to be used in the HTML template
      htmlTemplate.expireDate = expireDate; // Set expireDate variable to be used in the HTML template
      sheet.getRange(startRow + i, 8).setValue(emailSent); // Input 'Email Sent' into cell
    }
    else if (expired == dateToday && expired !== emailSent && redeemed !== "TRUE") { // Check if expire date equals today's date and if email has been sent or voucher has been redeemed
      htmlTemplate.timeFrame = "expired"; // Set timeFrame variable to be used in the HTML template
      sheet.getRange(startRow + i, 9).setValue(emailSent); // Input 'Email Sent' into cell
    }
    else { // if nothing matches previous conditions, skip row
      continue
    }
    
    var htmlBody = htmlTemplate.evaluate().getContent();
    // Send Email
    MailApp.sendEmail({
    to: emailAddress,
    subject: subject,
    htmlBody: htmlBody,
    replyTo: "YOUREMAIL@EMAIL.COM",
    name: "YOUR NAME"
  });
    SpreadsheetApp.flush();
  }
}
