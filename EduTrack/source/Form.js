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

//////////////
// syncData
///////////////
function manualSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- Ensure classList sheet exists ---
  let targetSheet = ss.getSheetByName("classList");
  if (!targetSheet) {
    targetSheet = ss.insertSheet("classList");
  } else {
    targetSheet.clear(); // optional: clear previous content
  }

  // --- Find temp folder ---
  const appFile = DriveApp.getFileById(ss.getId());
  const parentFolder = appFile.getParents().hasNext() ? appFile.getParents().next() : DriveApp.getRootFolder();
  const tempFolders = parentFolder.getFoldersByName("temp");
  if (!tempFolders.hasNext()) {
    SpreadsheetApp.getUi().alert("Folder 'temp' not found");
    return;
  }
  const tempFolder = tempFolders.next();

  // --- Loop through all subfolders (forms) ---
  const subfolders = tempFolder.getFolders();
  let headersSet = false;

  while (subfolders.hasNext()) {
    const formFolder = subfolders.next();
    const files = formFolder.getFilesByType(MimeType.GOOGLE_SHEETS);

    while (files.hasNext()) {
      const responseFile = files.next();
      const responseSs = SpreadsheetApp.openById(responseFile.getId());
      const responseSheet = responseSs.getSheets()[0];
      const data = responseSheet.getDataRange().getValues();

      if (data.length > 1) {
        // --- Set header if not yet ---
        if (!headersSet) {
          const formHeaders = data[0]; // header from response sheet
          targetSheet.appendRow(["Class Name", ...formHeaders]);
          targetSheet.getRange(1, 1, 1, 15).setFontWeight("bold").setBackground("#d9ead3");
          headersSet = true;
        }

        // --- Append data rows ---
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          // Get class name from form title (response sheet name: "Responses - {ClassName}")
          let className = responseSs.getName();
          const clsNames = className.split ('-')
          className = clsNames[clsNames.length-1]
          targetSheet.appendRow([className, ...row]);
        }
      }
    }
  }

  SpreadsheetApp.getUi().alert("Sync complete!");
}







///////////////////
downloalClassList
///////////////////
function downloalClassList () {
 
}
