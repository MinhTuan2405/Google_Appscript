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

function getGroupMembers(classname, groupname) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Class List');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();

  const emailsGroup = data
    .filter(row => row[0] == classname && row[1] == groupname)
    .map(row => [
      row[4],  // leader
      row[7],  // member 1
      row[10], // member 2
      row[13]  // member 3
    ])
    .flat(); // gộp thành 1 array thay vì [[..]]

  const cleaned = emailsGroup.filter(e => e && e.toString().trim() !== "");

  Logger.log(`${classname} - ${groupname}: ${JSON.stringify(cleaned)}`);
  Logger.log (cleaned)

  return cleaned;
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

///////////////////////////////////////////////////////////////
// function writePermissionsSheet (classname='COMP1314') {
//   const root = getSpreadsheetParent () // root;
//   const userprofile = getOrCreateFolder (root, 'userprofile')

//   const classfolder = getOrCreateFolder (userprofile, classname)

//   Logger.log (collectFolderIds (classfolder))
// }

// function writePermissionsSheet(classname = 'COMP1314') {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   let sheet = ss.getSheetByName('Permissions');
//   if (!sheet) sheet = ss.insertSheet('Permissions');

//   // === Build cây folder của class trong userprofile ===
//   const root = getSpreadsheetParent();
//   const userprofile = getOrCreateFolder(root, 'userprofile');
//   const classfolder = getOrCreateFolder(userprofile, classname);
//   const tree = collectFolderIds(classfolder);

//   // === Tìm depth lớn nhất của cây ===
//   let maxDepth = 0;
//   (function getMaxDepth(node, depth = 1) {
//     maxDepth = Math.max(maxDepth, depth);
//     if (node.children && node.children.length > 0) {
//       node.children.forEach(child => getMaxDepth(child, depth + 1));
//     }
//   })(tree);

//   // === Header ===
//   const header = ["classname", "groupname", "Folder ID", "Emails", "Permission"];
//   for (let i = 1; i < maxDepth; i++) {
//     header.push(`LEVEL ${i}`, "Folder ID", "Emails", "Permission");
//   }
//   sheet.clear();
//   sheet.getRange(1, 1, 1, header.length)
//        .setValues([header])
//        .setBackground("#c9daf8")
//        .setFontWeight("bold");

//   // === Build rows ===
//   const rows = [];
//   const groupEmailsCache = {}; // cache email theo group

//   function traverse(node, depth = 1, groupName = null) {
//     const row = Array(header.length).fill('');

//     if (depth === 1) {
//       // cấp 1: group folder
//       groupName = node.name;
//       row[0] = classname;
//       row[1] = groupName;
//       row[2] = node.id;

//       if (!groupEmailsCache[groupName]) {
//         groupEmailsCache[groupName] = getGroupMembers(classname, groupName);
//       }
//       row[3] = groupEmailsCache[groupName].join(", ");
//       row[4] = "editor"; // mặc định
//     }

//     // đặt dữ liệu cho level tương ứng
//     const colBase = 5 + (depth - 1) * 4;
//     row[colBase] = node.name || "";
//     row[colBase + 1] = node.id || "";

//     if (groupName && groupEmailsCache[groupName]) {
//       row[colBase + 2] = groupEmailsCache[groupName].join(", ");
//       row[colBase + 3] = "editor"; // mặc định
//     }

//     rows.push(row);

//     // duyệt con
//     if (node.children && node.children.length > 0) {
//       node.children.forEach(child => traverse(child, depth + 1, groupName));
//     }
//   }

//   if (tree.children) {
//     tree.children.forEach(child => traverse(child));
//   }

//   // === Ghi dữ liệu ===
//   sheet.getRange(2, 1, rows.length, header.length).setValues(rows);

//   // === Data validation cho Emails ===
//   const groups = getAllGroupOfClass(classname);
//   groups.forEach(group => {
//     const members = getGroupMembers(classname, group);
//     if (members.length > 0) {
//       const rule = SpreadsheetApp.newDataValidation()
//         .requireValueInList(members, true) // danh sách email nhóm
//         .setAllowInvalid(true)
//         .build();

//       // gắn cho toàn bộ cột Emails có group đó
//       const data = sheet.getDataRange().getValues();
//       for (let r = 1; r < data.length; r++) {
//         if (data[r][1] === group) {
//           header.forEach((h, c) => {
//             if (h === "Emails") {
//               sheet.getRange(r + 1, c + 1).setDataValidation(rule);
//             }
//           });
//         }
//       }
//     }
//   });

//   // === Permission dropdown (default editor) ===
//   const permOptions = ["viewer", "editor", "none"];
//   const permRule = SpreadsheetApp.newDataValidation()
//     .requireValueInList(permOptions, true)
//     .setAllowInvalid(false)
//     .build();

//   header.forEach((h, c) => {
//     if (h === "Permission") {
//       const range = sheet.getRange(2, c + 1, rows.length);
//       range.setDataValidation(permRule);
//     }
//   });
// }

/**
 * writePermissionsSheet - stable version
 * - Only lists groups that actually have folders in userprofile/classname
 * - Computes maxDepth across those group folders first, then builds header
 * - Ensures each row length === header.length (avoids setValues column mismatch)
 * - Adds DataValidation for Emails (per group members) and Permission (global)
 */
function writePermissionsSheet(classname='COMP1254') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Permissions");
  if (!sheet) sheet = ss.insertSheet("Permissions");

  const root = getSpreadsheetParent();
  const userprofile = getOrCreateFolder(root, "userprofile");
  const classFolder = getOrCreateFolder(userprofile, classname);

  // 1) Lấy nhóm có folder trong Drive
  const groupsFromSheet = getAllGroupOfClass(classname);
  const groupsInfo = [];
  let globalMaxDepth = 0;

  groupsFromSheet.forEach(group => {
    const groupIter = classFolder.getFoldersByName(group);
    if (!groupIter.hasNext()) {
      Logger.log("Skip group (no folder): " + group);
      return;
    }
    const groupFolder = groupIter.next();
    const tree = collectFolderIds(groupFolder);
    const depth = getMaxDepth(tree);
    globalMaxDepth = Math.max(globalMaxDepth, depth);
    const members = getGroupMembers(classname, group);
    groupsInfo.push({ name: group, members: members, tree: tree });
  });

  if (groupsInfo.length === 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1).setValue("No groups with folders found for class: " + classname);
    return;
  }

  // 2) Header (chỉ viết 1 lần nếu sheet trống)
  const levelsCount = Math.max(0, globalMaxDepth - 1);
  const header = ["classname", "groupname", "Folder ID", "Emails", "Permission"];
  for (let i = 1; i <= levelsCount; i++) {
    header.push(`LEVEL ${i}`, "Folder ID", "Emails", "Permission");
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header])
      .setBackground("#d9ead3").setFontWeight("bold");
  }

  function levelBaseIndex(level) {
    return 5 + (level - 1) * 4; // 0-based
  }

  // 3) Build rows
  const allRows = [];
  const rowGroups = []; // để lưu group cho mỗi row (dùng khi set validation)

  groupsInfo.forEach(info => {
    const members = info.members || [];
    const membersStr = members.join(", ");

    function traverse(node, level = 0) {
      const row = Array(header.length).fill("");

      if (level === 0) {
        // root row
        row[0] = classname;
        row[1] = info.name;
        row[2] = node.id || "";
        row[3] = membersStr;
        row[4] = "editor";
      } else {
        // level row
        const base = levelBaseIndex(level);
        if (base + 3 < header.length) {
          row[base] = node.name || "";
          row[base + 1] = node.id || "";
          row[base + 2] = membersStr;
          row[base + 3] = "editor";
        }
      }

      allRows.push(row);
      rowGroups.push(info.name);

      if (node.children && node.children.length > 0) {
        node.children.forEach(child => traverse(child, level + 1));
      }
    }

    traverse(info.tree, 0);
  });

  // 4) Append rows
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, allRows.length, header.length).setValues(allRows);

  // 5) Indices
  const emailColIndices = [];
  const permColIndices = [];
  for (let i = 0; i < header.length; i++) {
    if (header[i] === "Emails") emailColIndices.push(i + 1);
    if (header[i] === "Permission") permColIndices.push(i + 1);
  }

  // 6) Permission dropdown
  const permRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["editor", "viewer", "none"], true)
    .setAllowInvalid(true)
    .build();
  permColIndices.forEach(col => {
    sheet.getRange(startRow, col, allRows.length).setDataValidation(permRule);
  });

  // 7) Email dropdown theo group
  groupsInfo.forEach(info => {
    const members = info.members || [];
    if (!members.length) return;
    const emailRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(members, true)
      .setAllowInvalid(true)
      .build();

    rowGroups.forEach((g, idx) => {
      if (g === info.name) {
        emailColIndices.forEach(col => {
          sheet.getRange(startRow + idx, col).setDataValidation(emailRule);
        });
      }
    });
  });

  Logger.log("Permissions sheet appended for class " + classname +
    " with groups: " + groupsInfo.map(g => g.name).join(", "));
}



/**
 * Multi-select cho Emails
 */
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Permissions") return;

  const headerRow = 1;
  const editedRow = e.range.getRow();
  const editedCol = e.range.getColumn();

  if (editedRow <= headerRow) return;

  const header = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colName = header[editedCol - 1];

  if (colName !== "Emails") return;

  const newValue = e.value;
  if (!newValue) return;

  let oldValue = e.oldValue || "";
  let oldList = oldValue.split(",").map(s => s.trim()).filter(s => s !== "");

  if (oldList.indexOf(newValue) === -1) {
    oldList.push(newValue);
  } else {
    // Nếu chọn lại email đã có thì xóa đi (toggle)
    oldList = oldList.filter(v => v !== newValue);
  }

  const groupName = sheet.getRange(editedRow, 2).getValue();
  const className = sheet.getRange(editedRow, 1).getValue();
  const members = getGroupMembers(className, groupName);

  // giữ đúng thứ tự gốc
  const finalList = members.filter(m => oldList.indexOf(m) !== -1);

  sheet.getRange(editedRow, editedCol).setValue(finalList.join(", "));
}

/**
 * Tính độ sâu tối đa của cây
 */
function getMaxDepth(node, depth = 1) {
  let maxDepth = depth;
  if (node.children && node.children.length > 0) {
    node.children.forEach(child => {
      maxDepth = Math.max(maxDepth, getMaxDepth(child, depth + 1));
    });
  }
  return maxDepth;
}

/**
 * Thu thập ID + details thư mục
 */
function collectFolderIdsWithDetails(folder) {
  const children = [];
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const sub = subfolders.next();
    children.push(collectFolderIdsWithDetails(sub));
  }
  return { name: folder.getName(), id: folder.getId(), children: children };
}


////////////////////////////////////////////////////////////////////////////


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

