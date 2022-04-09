function autoFillGoogleDocFromForm(e) {
  // Yellow Shirts Love Letters Automation Script
  // Written by Brandon Winsley 2022-01-03

  const templateName = "Love Letters Template";

  //e.values is an array of form values
  const timestamp = e.values[0];
  const dsName = e.values[1];

  // Calculate column number
  let i = 0;
  switch(dsName) {
    case "Corlettos":
      i = 1;
      break;
    case "Dabees":
      i = 2;
      break;
    case "Durians (OT)":
      i = 3;
      break;
    case "Eevios":
      i = 4;
      break;
    case "Hypers":
      i = 5;
      break;
    case "Kavocs":
      i = 6;
      break;
    case "Niternals":
      i = 7;
      break;
    case "Primacots":
      i = 8;
      break;
    case "Rockems":
      i = 9;
      break;
    case "Saikis":
      i = 10;
      break;
    case "Sylvines":
      i = 11;
      break;
  }
  console.log("i is " + i);
  
  const recipient = e.values[(i - 1) * 3 + 2];
  const message = e.values[(i - 1) * 3 + 3];
  const sender = e.values[(i - 1) * 3 + 4];

  // get folder of spreadsheet
  console.log("Getting spreadsheet folder...");
  let spreadsheetId =  SpreadsheetApp.getActiveSpreadsheet().getId();
  let spreadsheetFile =  DriveApp.getFileById(spreadsheetId);
  let rootFolderId = spreadsheetFile.getParents().next().getId();
  let rootFolder = DriveApp.getFolderById(rootFolderId);
  console.log("Got spreadsheet folder!")

  // create DS folder if it does not exist
  if (!rootFolder.getFoldersByName(dsName).hasNext()) {
    console.log("Folder with name: <" + dsName + "> does not exist. Creating it now...");
    rootFolder.createFolder(dsName);
    console.log(dsName + " folder created!");
  }

  // get DS folder
  console.log("Getting folder of DS...");
  let folders = rootFolder.getFoldersByName(dsName);
  dsFolder = folders.next();
  console.log("Got folder of DS!");

  // Create new document if squaddie does not have one yet
  if (!dsFolder.getFilesByName(recipient).hasNext()) {
    console.log("Squaddie does not have a document yet. Making it now. If program breaks after this, check that the template exists in the root folder");

    let template = DriveApp.getFileById(rootFolder.getFilesByName(templateName).next().getId());
    console.log(template.getId());
    let copy = template.makeCopy(recipient, dsFolder);

    let doc = DocumentApp.openById(copy.getId());
    let body = doc.getBody();
    body.replaceText('{{Recipient}}', recipient);
    body.replaceText('{{DS Name}}', dsName);

    doc.saveAndClose();
    console.log("Document made successfully!");
  }

  console.log("Getting squaddie's document... If it fails on the next line, then the file does not exist and was not created successfully :(");
  let file = dsFolder.getFilesByName(recipient).next();
  let doc = DocumentApp.openById(file.getId());
  
  console.log("Got document! Adding message...");
  let body = doc.getBody();
  let para = body.appendParagraph(message);
  para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  if (sender != '') {
    let para = body.appendParagraph(sender);
    para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }
  para = body.appendParagraph("");
  para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  doc.saveAndClose();
  console.log("Added message! Script is done.");
}
