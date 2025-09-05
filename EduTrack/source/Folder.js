function showFolderCreatorSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('folderCreationUI') // ref: folderCreationUI
    .setTitle('Folder Creator');
  SpreadsheetApp.getUi().showSidebar(html);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////
function writeClassTree(classname, obj) {
  if (!obj || obj.length === 0) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Create Folder');
  if (!sheet) {
    sheet = ss.insertSheet('Create Folder');
  }

  let maxDepth = 0;
  function getMaxDepth(node, depth = 1) {
    maxDepth = Math.max(maxDepth, depth);
    if (node.children && node.children.length > 0) {
      node.children.forEach(child => getMaxDepth(child, depth + 1));
    }
  }
  obj.forEach(root => getMaxDepth(root));

  const existingHeader = sheet.getLastRow() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
  const existingDepth = existingHeader.filter(h => h.startsWith("LEVEL")).length;

  const requiredDepth = Math.max(existingDepth, maxDepth);

  const header = ["classname", "groupname"];
  for (let i = 1; i <= requiredDepth; i++) header.push(`LEVEL ${i}`);

  if (existingHeader.length === 0) {
    sheet.getRange(1, 1, 1, header.length)
         .setValues([header])
         .setBackground("#d9ead3")
         .setFontWeight("bold");
  } else if (requiredDepth > existingDepth) {
    sheet.getRange(1, 1, 1, header.length).setValues([header])
          .setBackground("#d9ead3")
         .setFontWeight("bold");;
  }

  let startRow = 0; const lr = sheet.getLastRow (); 
  if (lr == 1) startRow = lr + 1 
  else startRow = lr + 2;

  const rows = [];
  let firstRow = true;
  function traverse(node, level = 0) {
    const row = Array(requiredDepth + 2).fill('');
    if (firstRow) {
      row[0] = classname;  
      row[1] = "{}";       
      firstRow = false;
    }
    row[2 + level] = node.name;
    rows.push(row);

    if (node.children && node.children.length > 0) {
      node.children.forEach(child => traverse(child, level + 1));
    }
  }

  obj.forEach(root => traverse(root));

  sheet.getRange(startRow, 1, rows.length, requiredDepth + 2).setValues(rows);

  sheet.getRange(startRow, 1).setFontWeight("bold");
}




// function test () {
//   // Test
//   const testObj = [
//     {
//       name: "lab01",
//       children: [
//         { name: "demo1", children: [{ name: "ass1", children: [] }] }
//       ]
//     },
//     {
//       name: "lab02",
//       children: [
//         { name: "demo2", children: [] }
//       ]
//     },
//         {
//       name: "lab03",
//       children: [
//         { name: "demo2", children: [] }
//       ]
//     }
//   ];

//   writeClassTree("it006", testObj);
//   writeClassTree ('it007', testObj)
//   writeClassTree ('it008', testObj)

// }


//------------------------------------------------------------------


function getAllClasses(classname = 'Class List') { // default by 'Class List'
  const sheet = SpreadsheetApp
                .getActiveSpreadsheet()
                .getSheetByName(classname);

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

function getAllGroupOfClass(classname) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Class List');

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

function getGroupMembers (classname, groupname) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet ().getSheetByName ('Class List')
  if (!sheet) return [];

  const lastRow = sheet.getLastRow ();
  if (lastRow < 2) return [];

  const data = sheet.getRange (2, 1, lastRow - 1, 14).getValues ();

  const emailsGroup = data
                      .filter (row => row[0] == classname && row[1] == groupname)
                      .map (row => {
                        return [row[4], // leader
                                row[7], // member 1
                                row[10], // member 2
                                row[13]] // member 3
                      })
                      
  Logger.log (emailsGroup[0])

  return emailsGroup[0].length > 0 ? emailsGroup[0] : []
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
function addGroupToClass(classname, groupName) {
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

  // === Apply permissions for this group ===
  const members = getGroupMembers(classname, groupName);
  // let members = []
  // if (groupName == 'group01')
  //   members = ['nhungpth.cnthongtin@gmail.com']
  // else
  //   members = ['23521718@gm.uit.edu.vn']
  applyGroupPermissions(groupFolder, members);

  // Return ID tree
  const ids = collectFolderIds(groupFolder)
  Logger.log (ids)
  return ids;
}

/**
 * === PERMISSIONS HANDLING ===
 */
// function applyGroupPermissions(folder, groupMembers) {
//   // 1. Reset link-sharing: restricted
//   folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);

//   // 2. Remove any old editors/viewers
//   removeNonOwnerPermissions(folder);

//   // 3. Grant fresh permissions
//   groupMembers.forEach(email => {
//     if (email) {
//       try {
//         folder.addEditor(email);
//       } catch (e) {
//         Logger.log("Failed to add " + email + ": " + e.message);
//       }
//     }
//   });

//   // 4. Recurse for subfolders
//   const subfolders = folder.getFolders();
//   while (subfolders.hasNext()) {
//     applyGroupPermissions(subfolders.next(), groupMembers);
//   }
// }


// /**
//  * Remove all editors and viewers except the owner
//  */
// function removeNonOwnerPermissions(folder) {
//   const file = DriveApp.getFileById(folder.getId());

//   // Remove all editors
//   file.getEditors().forEach(user => {
//     try {
//       file.removeEditor(user);
//     } catch (e) {
//       Logger.log("Cannot remove editor " + user.getEmail() + ": " + e.message);
//     }
//   });

//   // Remove all viewers
//   file.getViewers().forEach(user => {
//     try {
//       file.removeViewer(user);
//     } catch (e) {
//       Logger.log("Cannot remove viewer " + user.getEmail() + ": " + e.message);
//     }
//   });
// }


/**
 * === PERMISSIONS HANDLING (No Email, Drive API) ===
 * Requires Advanced Drive Service enabled: Resources → Advanced Google Services → Drive API
 */
function applyGroupPermissions(folder, groupMembers) {
   // 1. Reset link-sharing: restricted
  folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  const folderId = folder.getId();

  // 1. Remove all permissions except owner
  removeNonOwnerPermissions(folderId);

  // 2. Grant fresh permissions (no email)
  groupMembers.forEach(email => {
    if (email) {
      try {
        addEditorNoEmail(folderId, email);
      } catch (e) {
        Logger.log("Failed to add " + email + ": " + e.message);
      }
    }
  });

  // 3. Recurse for subfolders
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    applyGroupPermissions(subfolders.next(), groupMembers);
  }
}

/**
 * Remove all editors and viewers except the owner
 */
function removeNonOwnerPermissions(folderId) {
  const file = Drive.Files.get(folderId, { fields: 'owners,permissions(id,type,emailAddress,role)' });
  const ownerEmails = file.owners.map(o => o.emailAddress);
  const permissions = file.permissions || [];

  permissions.forEach(p => {
    if (p.type === 'user' && p.emailAddress && !ownerEmails.includes(p.emailAddress)) {
      try {
        Drive.Permissions.remove(folderId, p.id);
      } catch (e) {
        Logger.log("Cannot remove " + p.emailAddress + ": " + e.message);
      }
    }
  });
}

/**
 * Add editor without sending email
 */
function addEditorNoEmail(fileId, email) {
  const permission = {
    'type': 'user',
    'role': 'writer',
    'emailAddress': email
  };
  Drive.Permissions.create(permission, fileId, { sendNotificationEmail: false });
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

function sheetToObjectByClassname(classname='IE107') {
  if (!classname) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Create Folder");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const header = data[0];
  const classnameCol = header.indexOf("classname");
  const levelCols = header.map((h, i) => h.startsWith("LEVEL") ? i : -1).filter(i => i !== -1);

  const tree = [];
  const stack = [];
  let currentClass = null;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (row.every(cell => cell === "")) continue;

    if (row[classnameCol] && row[classnameCol].toString().trim() !== "") {
      currentClass = row[classnameCol].toString().trim();
      stack.length = 0; 
    }

    if (currentClass !== classname) continue; 

    let level = -1;
    let nodeName = null;
    for (let j = 0; j < levelCols.length; j++) {
      const val = row[levelCols[j]];
      if (val && val.toString().trim() !== "") {
        level = j;
        nodeName = val.toString().trim();
        break;
      }
    }
    if (level === -1) continue;

    const node = { name: nodeName, children: [] };

    if (level === 0) {
      tree.push(node);
      stack.length = 0;
      stack.push(node);
    } else {
      const parent = stack[level - 1];
      if (parent) {
        parent.children.push(node);
        stack[level] = node;
      }
    }
  }

  Logger.log(JSON.stringify(tree, null, 2));
  return tree;
}

