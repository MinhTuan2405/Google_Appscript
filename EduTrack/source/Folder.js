function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('folderCreationUI') // ref: folderCreationUI
    .setTitle('Folder Pattern Creator');
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

