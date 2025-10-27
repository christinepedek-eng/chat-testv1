// ----- CONFIGURATION -----
const SHEET_NAME = 'ChatLog';
const IMAGE_FOLDER_ID = '1GjAw2771ia45TCr96d24dzAWmMW1hsZ-';

// ----- SHEET COLUMNS -----
const COL = {
  TIMESTAMP: 1,
  USER: 2,
  MESSAGE: 3,
  IMAGE_ID: 4
};

// ----- USER IDENTIFICATION -----
const USERS = {
  'steliosgfx@gmail.com': 'Εσύ',
  'yorgooos@hotmail.com': 'Αδερφός'
};

// ---------------------------
// ----- CORE FUNCTIONS -----
// ---------------------------

function doGet(e) {
  if (e.parameter.image) {
    return serveImage(e.parameter.image);
  } else {
    return serveChatPage();
  }
}

function serveChatPage() {
  const userEmail = Session.getActiveUser().getEmail();
  const userName = USERS[userEmail] || userEmail;

  const tpl = HtmlService.createTemplateFromFile('chat');
  tpl.userName = userName;
  // *** FINAL FIX: Pass the script URL to the template correctly ***
  tpl.scriptUrl = ScriptApp.getService().getUrl();

  return tpl.evaluate()
    .setTitle('Chat App')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function serveImage(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return ContentService.createOutput(blob).setMimeType(blob.getContentType());
  } catch (error) {
    Logger.log(`Error serving image ${fileId}: ${error.toString()}`);
    return ContentService.createTextOutput('Image not found').setMimeType(ContentService.MimeType.TEXT);
  }
}

function getSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);
    newSheet.appendRow(['Timestamp', 'User', 'Message', 'Image ID']);
    newSheet.setFrozenRows(1);
    return newSheet;
  }
  return sheet;
}

function getMessages() {
  try {
    const sheet = getSheet();
    const range = sheet.getDataRange();
    const values = range.getValues();
    const messages = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (row.join('').trim() === '') continue;

      const timestamp = new Date(row[COL.TIMESTAMP - 1]).toISOString();

      messages.push({
        timestamp: timestamp,
        user: row[COL.USER - 1],
        message: row[COL.MESSAGE - 1],
        imageId: row[COL.IMAGE_ID - 1]
      });
    }
    return { ok: true, messages: messages };
  } catch (error) {
    Logger.log('Error in getMessages: ' + error.toString());
    return { ok: false, error: 'Could not fetch messages. ' + error.message };
  }
}

function sendMessage(userName, messageText) {
  if (!userName || !messageText || messageText.trim() === '') {
    return { ok: false, error: 'User or message missing.' };
  }
  try {
    getSheet().appendRow([ new Date(), userName, messageText.trim(), '' ]);
    return { ok: true };
  } catch (error) {
    Logger.log('Error in sendMessage: ' + error.toString());
    return { ok: false, error: 'Could not send message. ' + error.message };
  }
}

function uploadImage(userName, fileInfo) {
  if (!userName || !fileInfo || !fileInfo.mimeType || !fileInfo.fileData) {
     return { ok: false, error: 'Missing image data.' };
  }
  try {
    const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
    if (!folder) {
      Logger.log(`Critical Error: Folder with ID "${IMAGE_FOLDER_ID}" not found.`);
      return { ok: false, error: 'Configuration error: Image folder not found.' };
    }

    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const decodedData = Utilities.base64Decode(fileInfo.fileData);
    const blob = Utilities.newBlob(decodedData, fileInfo.mimeType, fileInfo.fileName);
    const file = folder.createFile(blob);
    const fileId = file.getId();

    getSheet().appendRow([ new Date(), userName, '', fileId ]);

    return { ok: true, imageId: fileId };

  } catch (error) {
    Logger.log('Error in uploadImage: ' + error.toString());
    return { ok: false, error: 'Could not upload image: ' + error.message };
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}