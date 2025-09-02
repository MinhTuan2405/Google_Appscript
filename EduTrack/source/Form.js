function mainFormBuilder () {
  var html = HtmlService.createHtmlOutputFromFile('sidebar_ui') // ref: sidebar_ui.html
    .setTitle (' ')
    .setWidth (300)
  SpreadsheetApp.getUi().showSidebar(html);
}

// build the form from 'form_' sheet
function buildForm(subject, classCode, deadline, notes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("form_");
  const data = sheet.getDataRange().getValues();
  data.shift(); // remove header row

  // --- Create Form ---
  const formTitle = subject + " - " + classCode;
  const form = FormApp.create(formTitle)
    .setDescription(notes + "\nDeadline: " + deadline);

  // --- Add questions ---
  data.forEach(row => {
    const question = row[0];
    const type = row[1];

    let item;
    switch (type.toLowerCase()) {
      case "text":
        item = form.addTextItem().setTitle(question);
        break;
      case "paragraph":
        item = form.addParagraphTextItem().setTitle(question);
        break;
      case "multiple choice":
        item = form.addMultipleChoiceItem()
          .setTitle(question)
          .setChoices([
            form.addMultipleChoiceItem().createChoice("Option 1"),
            form.addMultipleChoiceItem().createChoice("Option 2"),
          ]);
        break;
      default:
        item = form.addTextItem().setTitle(question);
        break;
    }
    item.setRequired(true);
  });

  const editUrl = form.getEditUrl();
  const publishUrl = form.getPublishedUrl();
  const formId = form.getId();

  // --- Get parent folder of the app spreadsheet ---
  const appFile = DriveApp.getFileById(ss.getId());
  const parents = appFile.getParents();
  let parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  // --- Find or create "temp" folder ---
  let tempFolder;
  const folders = parentFolder.getFoldersByName("temp");
  tempFolder = folders.hasNext() ? folders.next() : parentFolder.createFolder("temp");

  // --- Create a subfolder for this form ---
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
  const formFolder = tempFolder.createFolder(formId + "_" + timestamp);

  // --- Move form into the subfolder ---
  const formFile = DriveApp.getFileById(formId);
  formFile.moveTo(formFolder);

  // --- Create response sheet (R) ---
  const responseSs = SpreadsheetApp.create("Responses - " + formTitle);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, responseSs.getId());

  // Move response sheet into the same subfolder
  const responseFile = DriveApp.getFileById(responseSs.getId());
  responseFile.moveTo(formFolder);

  // --- Create trigger for syncing ---
  ScriptApp.newTrigger("syncToClassList")
    .forSpreadsheet(responseSs)
    .onFormSubmit()
    .create();

  // --- Write log ---
  let logSheet = ss.getSheetByName("logs");
  if (!logSheet) {
    logSheet = ss.insertSheet("logs");
    logSheet.appendRow(["Timestamp", "Subject", "Class", "Publish URL", "Edit URL", "Folder URL", "Response Sheet"]);
    logSheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#d9ead3");
  }
  logSheet.appendRow([new Date(), subject, classCode, publishUrl, editUrl, formFolder.getUrl(), responseSs.getUrl()]);

  return publishUrl; // return form link
}


function syncToClassList(e) {
  const appSs = SpreadsheetApp.getActiveSpreadsheet(); // root spreadsheet
  let targetSheet = appSs.getSheetByName("classList");
  if (!targetSheet) {
    targetSheet = appSs.insertSheet("classList");
  }

  // --- Copy header if empty ---
  if (targetSheet.getLastRow() === 0) {
    // e.values chưa có header, nên lấy từ form responses sheet đầu tiên
    const formResponsesId = e.source.getId();
    const responseSheet = SpreadsheetApp.openById(formResponsesId).getSheets()[0];
    const headers = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0];
    targetSheet.appendRow(headers);
  }

  // --- Append new data ---
  targetSheet.appendRow(e.values);
}



//////////////
syncData
///////////////
function syncData () {
  
}



///////////////////
downloalClassList
///////////////////
function downloalClassList () {
 
}
