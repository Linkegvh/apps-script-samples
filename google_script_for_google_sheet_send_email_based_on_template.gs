// This constant is written in column C for rows for which an email
// has been sent successfully.
var EMAIL_SENT = 'EMAIL_SENT';
var space = ' ';
var file = DriveApp.getFileById('1D_ncXCJX4xvkiHZl15QnpfKkgvZl9f8eWl_BtLlgMLE');

/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function sendEmails2() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 3; // First row of data to process
  var numRows = 1000; // Number of rows to process
  // getRange(row, column, numRows, numColumns)
  var dataRange = sheet.getRange(startRow, 1, numRows, 8);
  // Fetch values for each row in the Range.
  // Returns the rectangular grid of values for this range. Returns a two-dimensional array of values, indexed by row, then by column. 
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0]; // First column
    var message = 'Dear participant of our COVID-19, What kind of Resident are you? Survey, the result of the quiz that you have just performed is in the attachment file below.';
    var emailSent = row[7]; // 8th Column
    if (emailSent !== 'EMAIL_SENT' && emailAddress) { // Prevents sending duplicates
      // we will make a copy of the template and save it in a sub-folder
      var folder = DriveApp.getFolderById('13riewWgcAsIQDZtDgGY03__Bb8wBCgPb');
      var copy = file.makeCopy(row[1], folder); 
      
      // start manipulating the google doc
      var doc = DocumentApp.openById(copy.getId());
      var body = doc.getBody();
      // we will replace some part of the document
      body.replaceText('{Openness paragraph}', row[2]);
      body.replaceText('{Agreeableness paragraph}', row[3]);
      body.replaceText('{Conscientiousness paragraph}', row[4]);
      body.replaceText('{Extraversion paragraph}', row[5]);
      body.replaceText('{Neuroticism paragraph}', row[6]);
      
      doc.saveAndClose(); 
      
      var subject = 'Sending emails from a Spreadsheet';
      //var message = 'somehow this does not work';
      MailApp.sendEmail(emailAddress, subject, message,{
        name: 'Automatic Emailer Script', attachments: [doc]});
      sheet.getRange(startRow + i, 8).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}
