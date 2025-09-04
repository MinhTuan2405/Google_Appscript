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

function getAllGroupOfClass(classname = 'COMP1314') {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('classList');

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; 

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  const cleanedData = data
    .filter(row => row[0] === classname) 
    .map(row => row[1])                 
    .filter(val => val);                 
    
  const distinctSet = new Set(cleanedData);
  const res = Array.from(distinctSet);
  return res;
}


function createClassDrive(inputClassname, obj) {
  // const classname = "classA"; 
  // const groups = ["group01", "group02"]; // initial groups
  // const data = [
  //   {
  //     "name": "lab01",
  //     "children": [
  //       {
  //         "name": "demo1",
  //         "children": [
  //           {
  //             "name": "ass1",
  //             "children": []
  //           }
  //         ]
  //       }
  //     ]
  //   },
  //   {
  //     "name": "lab02",
  //     "children": [
  //       {
  //         "name": "demo2",
  //         "children": []
  //       }
  //     ]
  //   }
  // ];

  const classname = inputClassname;
  const data = obj

  const groups = getAllGroupOfClass (classname);

  const parentFolder = getSpreadsheetParent();

  // === Step 1: temp/classA/_template ===
  let tempFolder = getOrCreateFolder(parentFolder, "temp");
  let classFolder = getOrCreateFolder(tempFolder, classname);
  let templateFolder = getOrCreateFolder(classFolder, "_template");

  // Build template once if empty
  if (!templateFolder.getFolders().hasNext() && !templateFolder.getFiles().hasNext()) {
    createFolders(data, templateFolder);
  }

  // === Step 2: userprofile/groups ===
  groups.forEach(group => {
    let ids = addGroupToClass(classname, group);
    Logger.log(JSON.stringify(ids, null, 2)); // log ID tree of group
  });
}

/**
 * Add a new group into an existing class
 * - classname: e.g. "classA"
 * - groupName: e.g. "group05"
 */
function addGroupToClass(classname='classA', groupName='demo') {
  const parentFolder = getSpreadsheetParent();

  // === Find template: temp/classA/_template ===
  let tempFolder = getOrCreateFolder(parentFolder, "temp");
  let classFolder = getOrCreateFolder(tempFolder, classname);
  let templateFolders = classFolder.getFoldersByName("_template");
  if (!templateFolders.hasNext()) {
    throw new Error("No _template folder found for " + classname);
  }
  let templateFolder = templateFolders.next();

  // === userprofile/groupName ===
  let userProfileFolder = getOrCreateFolder(parentFolder, "userprofile");
  let classProfileFolder = getOrCreateFolder (userProfileFolder, classname)
  let groupFolder = getOrCreateFolder(classProfileFolder, groupName);

  // If empty, copy template into it
  if (!groupFolder.getFolders().hasNext() && !groupFolder.getFiles().hasNext()) {
    copyContents(templateFolder, groupFolder);
    Logger.log("Created new group: " + groupName);
  } else {
    Logger.log("Group " + groupName + " already exists, skipped copy");
  }

  // Return ID tree
  const ids = collectFolderIds(groupFolder)
  Logger.log (ids)
  return ids;
}

/**
 * Recursively collect IDs of a folder and its subfolders
 * returns { name, id, children: [] }
 */
function collectFolderIds(folder) {
  let children = [];
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const sub = subfolders.next();
    children.push(collectFolderIds(sub));
  }
  return { name: folder.getName(), id: folder.getId(), children: children };
}


/**
 * Utility: get parent folder of spreadsheet (or root)
 */
function getSpreadsheetParent() {
  const ssFile = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const parents = ssFile.getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

/**
 * Utility: get or create subfolder
 */
function getOrCreateFolder(parent, name) {
  let folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

/**
 * Create folder structure from JSON
 */
function createFolders(nodes, parent) {
  nodes.forEach(node => {
    let folder = parent.createFolder(node.name);
    if (node.children && node.children.length > 0) {
      createFolders(node.children, folder);
    }
  });
}

/**
 * Copy folder contents recursively
 */
function copyContents(source, target) {
  // Copy files
  const files = source.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    file.makeCopy(file.getName(), target);
  }

  // Copy subfolders
  const subfolders = source.getFolders();
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    const newSub = target.createFolder(subfolder.getName());
    copyContents(subfolder, newSub);
  }
}

/////////////////////////////////////////////
function showChangeFolderStructureSidebar () {
  const html = HtmlService.createHtmlOutputFromFile ('changeFolderStructureUI') // ref: changeFolderStructureUI.html
      .setTitle ('Change Folder Structure')

  SpreadsheetApp.getUi ().showSidebar (html);
}

