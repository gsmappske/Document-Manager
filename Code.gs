const SHEET_NAME = "Document Manager";
const DIRECT_UPLOADS_SHEET_NAME = "Direct Uploads";
const NO_FOLDER_MESSAGE = "No folders exist. Please add one.";

// Serve the main HTML page
function doGet() {
  const role = isAuthorized();

  if (!role) {
    return HtmlService.createHtmlOutput('<h1>Unauthorized Access</h1><p>You do not have permission to access this application.</p>');
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('Document Manager');
}

// Include HTML, CSS, or JS files dynamically
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Authorization check
function isAuthorized() {
  const email = Session.getActiveUser().getEmail();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User Management');

  if (!sheet) {
    throw "User Management sheet not found.";
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][2] === 'Active') { // Check email and active status
      return data[i][1]; // Return role (e.g., Admin, User)
    }
  }

  return null; // Not authorized
}

// Retrieve logged-in user details
function getLoggedUserDetails() {
  const email = Session.getActiveUser().getEmail();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User Management');

  if (!sheet) {
    throw "User Management sheet not found.";
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][2] === 'Active') {
      return {
        email: email,
        role: data[i][1] // Role: Admin or User
      };
    }
  }

  return {
    email: email,
    role: null // Not an authorized user
  };
}

// Retrieve folder hierarchy
function getFoldersAndSubfolders() {
  try {
    const rootFolderId = PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID');
    if (!rootFolderId) throw "ROOT_FOLDER_ID is not set.";

    const rootFolder = DriveApp.getFolderById(rootFolderId);
    const folders = rootFolder.getFolders();
    const folderHierarchy = {};

    while (folders.hasNext()) {
      const mainFolder = folders.next();
      const subfolders = mainFolder.getFolders();
      const subfolderNames = [];
      while (subfolders.hasNext()) {
        subfolderNames.push(subfolders.next().getName());
      }
      folderHierarchy[mainFolder.getName()] = subfolderNames;
    }

    return folderHierarchy;
  } catch (error) {
    Logger.log("Error fetching folder hierarchy: " + error);
    throw error;
  }
}

// Create a new folder
function createFolder(folderName, parentFolderName = null) {
  try {
    const rootFolderId = PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID');
    const rootFolder = DriveApp.getFolderById(rootFolderId);

    let targetFolder = rootFolder;

    if (parentFolderName) {
      const mainFolders = rootFolder.getFolders();
      let parentFound = false;
      while (mainFolders.hasNext()) {
        const folder = mainFolders.next();
        if (folder.getName() === parentFolderName) {
          targetFolder = folder;
          parentFound = true;
          break;
        }
      }

      if (!parentFound) throw `Parent folder '${parentFolderName}' not found.`;
    }

    const existingFolders = targetFolder.getFolders();
    while (existingFolders.hasNext()) {
      if (existingFolders.next().getName() === folderName) {
        return `Folder '${folderName}' already exists.`;
      }
    }

    targetFolder.createFolder(folderName);
    return `Folder '${folderName}' created successfully under '${parentFolderName || "Root"}'.`;
  } catch (error) {
    Logger.log("Error creating folder: " + error);
    return "Error creating folder: " + error.message;
  }
}

// Upload a file and log its details
function uploadFile(byteArray, fileName, fileType, mainFolderName, subFolderName = null) {
  try {
    const rootFolderId = PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID');
    const rootFolder = DriveApp.getFolderById(rootFolderId);

    // Locate the main folder
    let mainFolder = null;
    const folders = rootFolder.getFolders();
    while (folders.hasNext()) {
      const folder = folders.next();
      if (folder.getName() === mainFolderName) {
        mainFolder = folder;
        break;
      }
    }

    if (!mainFolder) {
      throw `Main folder '${mainFolderName}' not found.`;
    }

    // Locate the subfolder if specified
    let targetFolder = mainFolder;
    if (subFolderName) {
      const subfolders = mainFolder.getFolders();
      let subFolderFound = false;
      while (subfolders.hasNext()) {
        const folder = subfolders.next();
        if (folder.getName() === subFolderName) {
          targetFolder = folder;
          subFolderFound = true;
          break;
        }
      }
      if (!subFolderFound) {
        throw `Subfolder '${subFolderName}' not found under '${mainFolderName}'.`;
      }
    }

    // Upload the file
    const blob = Utilities.newBlob(new Uint8Array(byteArray), fileType, fileName);
    const uploadedFile = targetFolder.createFile(blob);
    const fileUrl = uploadedFile.getUrl();

    // Log the upload
    const userEmail = Session.getActiveUser().getEmail();
    logDirectUpload(userEmail, mainFolderName, subFolderName, fileName, fileUrl);

    return `File uploaded successfully to '${mainFolderName}${subFolderName ? " > " + subFolderName : ""}'.`;
  } catch (error) {
    Logger.log("Error uploading file: " + error);
    throw error;
  }
}

// Log direct uploads into a dedicated sheet
function logDirectlyUploadedFiles() {
  try {
    const rootFolderId = PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID');
    const rootFolder = DriveApp.getFolderById(rootFolderId);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let directUploadsSheet = ss.getSheetByName(DIRECT_UPLOADS_SHEET_NAME);
    if (!directUploadsSheet) {
      directUploadsSheet = ss.insertSheet(DIRECT_UPLOADS_SHEET_NAME);
      directUploadsSheet.appendRow([
        'Serial #',
        'User Email',
        'Main Folder',
        'Subfolder',
        'File Name',
        'File URL',
        'Timestamp'
      ]);
    }

    const existingLogs = directUploadsSheet.getDataRange().getValues();
    const loggedFiles = new Set(existingLogs.map(row => row[4])); // Column 5: File Name

    processFolder(rootFolder, loggedFiles, directUploadsSheet);
  } catch (error) {
    Logger.log(`Error logging direct uploads: ${error.message}`);
  }
}

// Helper function to recursively process folders
function processFolder(folder, loggedFiles, sheet, mainFolderName = null, subFolderName = null) {
  const files = folder.getFiles();
  let serialNumber = sheet.getLastRow(); // Start serial number from the last row

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const fileUrl = file.getUrl();
    const timestamp = new Date();
    const userEmail = Session.getActiveUser().getEmail();

    if (loggedFiles.has(fileName)) continue;

    serialNumber += 1;
    sheet.appendRow([
      serialNumber,
      userEmail,
      mainFolderName || folder.getName(),
      subFolderName || '-',
      fileName,
      fileUrl,
      timestamp
    ]);
  }

  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    const newMainFolderName = mainFolderName || folder.getName();
    const newSubFolderName = subfolder.getName();
    processFolder(subfolder, loggedFiles, sheet, newMainFolderName, newSubFolderName);
  }
}

// Log a single file upload
function logDirectUpload(userEmail, mainFolderName, subFolderName, fileName, fileUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let directUploadsSheet = ss.getSheetByName(DIRECT_UPLOADS_SHEET_NAME);
  if (!directUploadsSheet) {
    directUploadsSheet = ss.insertSheet(DIRECT_UPLOADS_SHEET_NAME);
    directUploadsSheet.appendRow([
      'Serial #',
      'User Email',
      'Main Folder',
      'Subfolder',
      'File Name',
      'File URL',
      'Timestamp'
    ]);
  }

  const serialNumber = directUploadsSheet.getLastRow();
  directUploadsSheet.appendRow([
    serialNumber,
    userEmail,
    mainFolderName,
    subFolderName || '-',
    fileName,
    fileUrl,
    new Date()
  ]);
}
