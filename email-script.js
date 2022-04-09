function onOpen() 
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation')
    .addItem('send PDF Form', 'sendPDFForm')
    .addItem('send to all', 'sendFormToAll')
    .addToUi();
}

function sendPDFForm()
{
  var row = SpreadsheetApp.getActiveSheet().getActiveCell().getRow();
  sendEmailWithAttachment(row);
}

function sendFormToAll()
{
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var last_row = sheet.getDataRange().getLastRow();
  
  for(var row = 2; row <= last_row; row++)
  {
    sendEmailWithAttachment(row);
  }
}

function sendEmailWithAttachment(row)
{
  var client = getClientInfo(row);
  var sheet = SpreadsheetApp.getActiveSheet();

  var filename = client.name;
  var file = DriveApp.getFilesByName(filename);
  if (!file.hasNext()) 
  {
    console.error("Could not open file " + filename);
    sheet.getRange(row,3).setValue("failed (no file)");
    return;
  }
  
  var template = HtmlService.createTemplateFromFile('LoveLettersTemplate');
  template.client = client;
  var message = template.evaluate().getContent();
  
  try {
    MailApp.sendEmail({
      to: client.email,
      subject: "Yellow Shirts Love Letters 2022",
      htmlBody: message,
      attachments: [file.next().getAs(MimeType.PDF)]
    });
    sheet.getRange(row,3).setValue("email sent");
  } catch {
    sheet.getRange(row,3).setValue("failed (email failed to send)");
  }
}

function getClientInfo(row)
{
  var sheet = SpreadsheetApp.getActive().getSheetByName('Sheet1');
   
  var values = sheet.getRange(row,1,row,3).getValues();
  var rec = values[0];
  
  var client = 
  {
    name: rec[0],
    email: rec[1]
  };
  return client;
}
