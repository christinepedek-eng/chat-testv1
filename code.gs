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

// doGet now only serves the chat page. No more routing.
function doGet(e) {
  const userEmail = Session.getActiveUser().getEmail();
  const userName = USERS[userEmail] || userEmail;

  const tpl = HtmlService.createTemplateFromFile('chat');
  tpl.userName = userName;
  return tpl.evaluate()
    .setTitle('Chat App')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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

/**
 * Gets messages and converts image IDs to Base64 Data URLs on the server.
 * This is the definitive fix.
 */
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
      const imageId = row[COL.IMAGE_ID - 1];
      let imageDataUrl = '';

      // If there's an image ID, fetch the image and convert it to a Data URL.
      if (imageId) {
        try {
          const imageFile = DriveApp.getFileById(imageId);
          const blob = imageFile.getBlob();
          const contentType = blob.getContentType();
          const base64Data = Utilities.encodeBase64(blob.getBytes());
          imageDataUrl = `data:${contentType};base64,${base64Data}`;
        } catch (e) {
          Logger.log(`Could not process image with ID ${imageId}: ${e.toString()}`);
        }
      }

      messages.push({
        timestamp: timestamp,
        user: row[COL.USER - 1],
        message: row[COL.MESSAGE - 1],
        imageData: imageDataUrl // Send the Data URL to the client
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

    const decodedData = Utilities.base64Decode(fileInfo.fileData);
    const blob = Utilities.newBlob(decodedData, fileInfo.mimeType, fileInfo.fileName);
    const file = folder.createFile(blob);
    const fileId = file.getId();

    // No sharing needed anymore. Just save the ID.

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