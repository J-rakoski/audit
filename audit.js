function auditProducts() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("your-sheet-name");
  
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  var emailAddress = 'your-email goes here'; // Update with your email address
  var subject = 'email-subject';
  
  var products = [];
  
  for (var i = 1; i < values.length; i++) {
    var productName = values[i][0]; // Column A
    var productNumber = values[i][9]; // Column J
    
    if (values[i][2] === 'key-word-in-excel') { // Column C
      var message = 'Product Name: ' + productName + '\nProduct Number: ' + productNumber;
      products.push(message);
    }
  }
  
  if (products.length > 0) {
    var body = products.join('\n\n');

    MailApp.sendEmail({
    to: emailAddress,
    subject: subject,
    body: body
  });
  }
}
