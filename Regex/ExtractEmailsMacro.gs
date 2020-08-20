function ExtractEmail(cell){
  
//First lets define our variables

//  Email Regex
var regExp = new RegExp( "([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)", "gi"); // "g" is for global,
                                                                                     //"i" is for case insensitive
// Extract Emails
var Email = regExp.exec(cell)[0];

return Email;
  
}

// The macro
function ExtractEmails() {
  
  // Define where your extracted values will be stored
  var spreadsheet = SpreadsheetApp.getActive();
  
  // First Calculated value
  spreadsheet.getRange('C2').activate();
  spreadsheet.getCurrentCell().setFormula('=ExtractEmail(B2)');
  
  // Apply accross your defined range
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C2:C51'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  
  // Name Column
  spreadsheet.getRange('C1').activate();
  spreadsheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('Email')
  .setTextStyle(0, 5, SpreadsheetApp.newTextStyle()
  .setBold(true)
  .build())
  .build());
  
  //Center title
  spreadsheet.getRange('C1').setHorizontalAlignment('center');
  spreadsheet.getRange('C1:C51').setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  //Autofit text
  //need to first call SpreadsheetApp.flush()
  SpreadsheetApp.flush();
  spreadsheet.getActiveSheet().autoResizeColumns(3, 1);
};

