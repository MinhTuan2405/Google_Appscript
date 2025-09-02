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

  // Create new form
  const formTitle = subject + " - " + classCode;
  const form = FormApp.create(formTitle)
    .setDescription(notes + "\nDeadline: " + deadline);

  // Add questions
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

  // Get current spreadsheet file
  const appFile = DriveApp.getFileById(ss.getId());
  const parents = appFile.getParents();

  let parentFolder;
  if (parents.hasNext()) {
    parentFolder = parents.next();
  } else {
    // without parent (file in root My Drive)
    parentFolder = DriveApp.getRootFolder();
  }

  // find or create "temp" folder
  let tempFolder;
  const folders = parentFolder.getFoldersByName("temp");
  if (folders.hasNext()) {
    tempFolder = folders.next();
  } else {
    tempFolder = parentFolder.createFolder("temp");
  }

  // move form into temp folder
  const formFile = DriveApp.getFileById(form.getId());
  formFile.moveTo(tempFolder);

  // ---- Write log into "logs" sheet ----
  let logSheet = ss.getSheetByName("logs");
  if (!logSheet) {
    logSheet = ss.insertSheet("logs");
    logSheet.appendRow(["Timestamp", "Subject", "Class", "Publish URL", "Edit URL"]);
  }
  logSheet.appendRow([new Date(), subject, classCode, publishUrl, editUrl]);

  return publishUrl; // return form link
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
