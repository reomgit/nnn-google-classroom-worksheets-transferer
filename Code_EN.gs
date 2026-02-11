// --- CONFIGURATION ---
var ROOT_FOLDER_NAME = "Classroom";
// ---------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎓 Classroom Manager')
    .addItem('1. Find My Classes', 'listClassFolders')
    .addItem('2. Fetch Files from Selected Class', 'fetchFilesFromSelection')
    .addItem('3. Move Files', 'processMoveQueue')
    .addToUi();
}

function listClassFolders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();

  sheet.appendRow(["Select", "Class Name", "Folder ID", "Folder URL"]);
  sheet.getRange("A1:D1").setFontWeight("bold");

  var folders = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  var rootFolder;

  if (folders.hasNext()) {
    rootFolder = folders.next();
  } else {
    SpreadsheetApp.getUi().alert("Could not find a folder named '" + ROOT_FOLDER_NAME + "'.");
    return;
  }

  try {
    var subfolders = rootFolder.getFolders();
    var rows = [];

    while (subfolders.hasNext()) {
      var folder = subfolders.next();
      rows.push(["☐", folder.getName(), folder.getId(), folder.getUrl()]);
    }

    if (rows.length > 0) {
      var range = sheet.getRange(2, 1, rows.length, 4);
      range.setValues(rows);
      sheet.getRange(2, 1, rows.length, 1).insertCheckboxes();
      SpreadsheetApp.getUi().alert("Found " + rows.length + " classes.");
    } else {
      SpreadsheetApp.getUi().alert("Found 'Classroom' folder, but it looks empty.");
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert("Error accessing folders: " + e.message);
  }
}

function fetchFilesFromSelection() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var selectedFolderId = null;
  var selectedClassName = "";

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === true) {
      selectedClassName = data[i][1];
      selectedFolderId = data[i][2];
      break;
    }
  }

  if (!selectedFolderId) {
    SpreadsheetApp.getUi().alert("Please check a box next to a class folder first.");
    return;
  }

  sheet.clear();
  sheet.getRange("A:E").clearDataValidations();

  sheet.appendRow(["Source: " + selectedClassName, "ID: " + selectedFolderId]);
  sheet.appendRow(["File Name", "File Link", "Target Folder URL (Paste Here)", "Status", "File ID"]);
  sheet.getRange("A2:E2").setFontWeight("bold");

  try {
    var folder = DriveApp.getFolderById(selectedFolderId);
    var files = folder.getFiles();
    var allFiles = [];

    var shortcutCount = 0;
    var brokenShortcutCount = 0;

    var formatRegex = /^\d{4}_/;

    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      var id = file.getId();
      var url = file.getUrl();
      var isShortcut = false;
      var targetAccessError = false;

      if (file.getMimeType() === "application/vnd.google-apps.shortcut") {
        shortcutCount++;
        try {
          var targetId = file.getTargetId();

          try {
            var targetFile = DriveApp.getFileById(targetId);
            url = targetFile.getUrl();
          } catch (accessError) {
             console.log("Could not access target object for: " + name);
             targetAccessError = true;
          }

          name = file.getName();
          id = targetId;
          isShortcut = true;

        } catch (e) {
          brokenShortcutCount++;
          console.log("Totally Broken Shortcut: " + file.getName());
          continue;
        }
      }

      if (targetAccessError) {
        name = name + " [⚠️ Restricted Target]";
      }

      allFiles.push({
        name: name,
        url: url,
        id: id,
        isShortcut: isShortcut
      });
    }

    var formattedFiles = [];
    var otherFiles = [];

    allFiles.forEach(function(item) {
      if (formatRegex.test(item.name)) {
        formattedFiles.push(item);
      } else {
        otherFiles.push(item);
      }
    });

    formattedFiles.sort(function(a, b) {
      return a.name.localeCompare(b.name, undefined, {numeric: true, sensitivity: 'base'});
    });

    otherFiles.sort(function(a, b) {
      return a.name.localeCompare(b.name, 'ja');
    });

    var sortedFiles = formattedFiles.concat(otherFiles);
    var rows = [];

    sortedFiles.forEach(function(file) {
      rows.push([file.name, file.url, "", "Pending", file.id]);
    });

    if (rows.length > 0) {
      sheet.getRange(3, 1, rows.length, 5).setValues(rows);
    } else {
      var msg = "No files listed.";
      if (shortcutCount > 0) {
        msg += "

Diagnostic Info:";
        msg += "
- Shortcuts found: " + shortcutCount;
        msg += "
- Truly Broken Shortcuts: " + brokenShortcutCount;
      }
      SpreadsheetApp.getUi().alert(msg);
      sheet.appendRow(["No files found.", "", "", "", ""]);
    }
  } catch (e) {
     SpreadsheetApp.getUi().alert("Error fetching files: " + e.message);
  }
}

function processMoveQueue() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  var dataRange = sheet.getRange(3, 1, lastRow - 2, 5);
  var data = dataRange.getValues();
  var ui = SpreadsheetApp.getUi();

  var processedCount = 0;
  var myEmail = Session.getActiveUser().getEmail();

  for (var i = 0; i < data.length; i++) {
    var sheetName = data[i][0];
    var targetUrl = data[i][2];
    var status = data[i][3];
    var fileId = data[i][4];

    if (status === "Done" || status.indexOf("Done") === 0 || status === "Skipped") continue;

    if (targetUrl !== "") {
      var log = [];
      try {
        var file = DriveApp.getFileById(fileId);

        var owner = file.getOwner();
        if (!owner || owner.getEmail() !== myEmail) {
          data[i][3] = "Skipped (Not Owner)";
          continue;
        }

        var targetFolderId = extractIdFromUrl(targetUrl);
        var targetFolder = DriveApp.getFolderById(targetFolderId);
        file.moveTo(targetFolder);
        log.push("Moved");

        if (file.getName() !== sheetName) {
           file.setName(sheetName);
           log.push("Renamed to Match");
        }

        data[i][3] = "Done (" + log.join(" & ") + ")";
        processedCount++;

      } catch (e) {
        data[i][3] = "Error: " + e.message;
      }
    }
  }

  dataRange.setValues(data);

  ui.alert('Processed ' + processedCount + ' files.');
}

function extractIdFromUrl(url) {
  if (!url) throw new Error("Empty URL");
  if (url.length < 50 && url.indexOf("http") === -1) return url;
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}
