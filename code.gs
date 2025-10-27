// ----- CONFIGURATION -----
const SHEET_NAME = 'ChatLog';
const IMAGE_FOLDER_ID = '1GjAw2771ia45TCr96d24dzAWmMW1hsZ-'; // <-- ΒΑΛΕ ΕΔΩ ΤΟ ID ΤΟΥ ΦΑΚΕΛΟΥ ΑΠΟ ΤΟ ΒΗΜΑ 2!

// ----- SHEET COLUMNS -----
const COL = {
  TIMESTAMP: 1,
  USER: 2,
  MESSAGE: 3,
  IMAGELINK: 4
};

// ----- USER IDENTIFICATION (Simple Example) -----
// Προσάρμοσέ το με τα δικά σας emails!
const USERS = {
  'steliosgfx@gmail.com': 'Εσύ', // <-- ΒΑΛΕ ΤΟ EMAIL ΣΟΥ
  'yorgooos@hotmail.com': 'Αδερφός' // <-- ΒΑΛΕ ΤΟ EMAIL ΤΟΥ ΑΔΕΡΦΟΥ ΣΟΥ
};

// ---------------------------
// ----- CORE FUNCTIONS -----
// ---------------------------

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
    newSheet.appendRow(['Timestamp', 'User', 'Message', 'Image Link']);
    newSheet.setFrozenRows(1);
    return newSheet;
  }
  return sheet;
}

// Function called by client-side JS to get messages
function getMessages() {
  try {
    const sheet = getSheet();
    const range = sheet.getDataRange();
    const values = range.getValues();
    const messages = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (row.join('').trim() === '') continue;

      const timestampValue = row[COL.TIMESTAMP - 1];
      const timestamp = new Date(timestampValue).toISOString();

      messages.push({
        timestamp: timestamp,
        user: row[COL.USER - 1],
        message: row[COL.MESSAGE - 1],
        imageLink: row[COL.IMAGELINK - 1]
      });
    }
    return { ok: true, messages: messages };
  } catch (error) {
    Logger.log('Error in getMessages: ' + error.toString());
    return { ok: false, error: 'Could not fetch messages. ' + error.message };
  }
}

// Function called by client-side JS to send a text message
function sendMessage(userName, messageText) {
  if (!userName || !messageText || messageText.trim() === '') {
    return { ok: false, error: 'User or message missing.' };
  }
  try {
    const sheet = getSheet();
    sheet.appendRow([ new Date(), userName, messageText.trim(), '' ]);
    return { ok: true };
  } catch (error) {
    Logger.log('Error in sendMessage: ' + error.toString());
    return { ok: false, error: 'Could not send message. ' + error.message };
  }
}

// Function called by client-side JS to upload an image
function uploadImage(userName, fileInfo) {
  if (!userName || !fileInfo || !fileInfo.mimeType || !fileInfo.fileData) {
     return { ok: false, error: 'Missing image data.' };
  }
  try {
    const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
    if (!folder) {
      Logger.log(`Critical Error: The folder with ID "${IMAGE_FOLDER_ID}" was not found.`);
      return { ok: false, error: 'Configuration error: Image folder not found.' };
    }

    const decodedData = Utilities.base64Decode(fileInfo.fileData);
    const blob = Utilities.newBlob(decodedData, fileInfo.mimeType, fileInfo.fileName);
    const file = folder.createFile(blob);

    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // *** BUG FIX: Increased delay to 3 seconds for permissions to propagate. ***
    // This provides a much safer window for Google Drive's systems to sync
    // the file's sharing permissions before the link is used.
    Utilities.sleep(3000); // Pause for 3 seconds

    const fileId = file.getId();
    const imageUrl = `https://drive.google.com/uc?id=${fileId}`;

    const sheet = getSheet();
    sheet.appendRow([ new Date(), userName, '', imageUrl ]);

    return { ok: true, imageLink: imageUrl };

  } catch (error) {
    Logger.log('Error in uploadImage: ' + error.toString());
    return { ok: false, error: 'Could not upload image: ' + error.message };
  }
}

// Helper to include HTML partials
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}