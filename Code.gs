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




const SHEET_NAME = "Document Manager";
const NO_FOLDER_MESSAGE = "No folders exist. Please add one.";

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

function extractTextContent(fileId) {
  try {
      const file = DriveApp.getFileById(fileId);
        const fileType = file.getMimeType();
      
        if (fileType === MimeType.PLAIN_TEXT || fileType === "text/csv") {
            return file.getBlob().getDataAsString();
          }else{
            return `Not a txt or csv file. File type is: ${fileType}`
          }

  } catch (e) {
      Logger.log("Error extracting file content: " + e);
      return "Error extracting file content: " + e.message;
  }
}


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

      //extract text content from the file
     let fileContent = null
      if(fileType === 'application/pdf') {
        fileContent = extractTextFromPDF(uploadedFile.getId())
     } else{
          fileContent = extractTextContent(uploadedFile.getId())
      }
    // Log the upload to the sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Document Data');
    if (!sheet) {
      // Create the sheet if it doesn't exist
      sheet = ss.insertSheet('Document Data');
      sheet.appendRow(['User Email', 'Main Folder', 'Subfolder', 'File Name', 'File URL', 'Timestamp','File content']);
    }

    const userEmail = Session.getActiveUser().getEmail();
    const timestamp = new Date();
    sheet.appendRow([userEmail, mainFolderName, subFolderName || '-', fileName, fileUrl, timestamp, fileContent]);

    return `File uploaded successfully to '${mainFolderName}${subFolderName ? " > " + subFolderName : ""}'.`;
  } catch (error) {
    Logger.log("Error uploading file: " + error);
    throw error;
  }
}

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


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();

}






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



function addUser(email, role) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('User Management');

  if (!sheet) {
    throw "User Management sheet not found.";
  }

  const currentUserDetails = getLoggedUserDetails();
  if (currentUserDetails.role !== 'Admin') {
    throw "Unauthorized action. Only administrators can add users.";
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      throw `User with email ${email} already exists.`;
    }
  }

  sheet.appendRow([email, role, 'Active']);
  return `User '${email}' added successfully as '${role}'.`;
}





// to include php functions.
function include (filename) {
return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
