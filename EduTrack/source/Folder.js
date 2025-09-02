function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('folderCreationUI') // ref: folderCreationUI
    .setTitle('Folder Pattern Creator');
  SpreadsheetApp.getUi().showSidebar(html);
}

// --- Get list of root folders (children of a parent folder) ---
function getRootFolders(parentFolderId) {
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const folders = [];
  const iterator = parentFolder.getFolders();
  while (iterator.hasNext()) {
    const f = iterator.next();
    folders.push({id: f.getId(), name: f.getName()});
  }
  return folders;
}

// --- Create subfolders under a root folder ---
function createSubfolders(rootFolderId, subfolderNames) {
  const rootFolder = DriveApp.getFolderById(rootFolderId);
  subfolderNames.forEach(name => {
    if (name) rootFolder.createFolder(name);
  });
  return `Created ${subfolderNames.length} subfolders under "${rootFolder.getName()}"`;
}