function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('folderCreationUI') // ref: folderCreationUI.html
    .setTitle('Folder Creator');
  SpreadsheetApp.getUi().showSidebar(html);
}

// --- Create folder with subfolders ---
function createFolderStructure(parentFolderName, subfolders) {
  const parentFolder = DriveApp.createFolder(parentFolderName);

  subfolders.forEach(name => {
    if (name) parentFolder.createFolder(name);
  });

  return `Folder "${parentFolderName}" created with ${subfolders.length} subfolders!`;
}
