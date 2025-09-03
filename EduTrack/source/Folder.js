function showFolderCreatorSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('folderCreationUI') // ref: folderCreationUI
    .setTitle('Folder Creator');
  SpreadsheetApp.getUi().showSidebar(html);
}


function getAllClasses() {
  const sheet = SpreadsheetApp
                .getActiveSpreadsheet()
                .getSheetByName('classList');

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; 

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  const cleanedData = data.map(row => row[0]).filter(val => val);

  const distinctSet = new Set(cleanedData);

  const res = Array.from(distinctSet);

  res.forEach(item => Logger.log(item));

  return res;
}


function createFolderStructure (classname, jsontr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet ()
  let folderSheet = ss.getSheetByName ('Folder Structure')

  if (!folderSheet) {
    ss.insertSheet ('Folder Structure')
    folderSheet = ss.getSheetByName ('Folder Structure')


  }
}


/////////////////////////////////////////////
function showChangeFolderStructureSidebar () {
  const html = HtmlService.createHtmlOutputFromFile ('changeFolderStructureUI') // ref: changeFolderStructureUI.html
      .setTitle ('Change Folder Structure')

  SpreadsheetApp.getUi ().showSidebar (html);
}

