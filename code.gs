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
    // If the sheet doesn't exist, create it with a header row.
    const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);
    newSheet.appendRow(['Timestamp', 'User', 'Message', 'Image Link']);
    // Freeze the header row for better readability
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

    // Skip header row (i=0)
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      // Skip empty rows
      if (row.join('').trim() === '') continue;

      messages.push({
        timestamp: row[COL.TIMESTAMP - 1],
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
    sheet.appendRow([
      new Date(), // Timestamp
      userName,   // User
      messageText.trim(), // Message
      ''          // ImageLink (empty for text messages)
    ]);
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
      // This is a critical configuration error.
      Logger.log(`Critical Error: The folder with ID "${IMAGE_FOLDER_ID}" was not found.`);
      return { ok: false, error: 'Configuration error: Image folder not found.' };
    }

    const decodedData = Utilities.base64Decode(fileInfo.fileData);
    const blob = Utilities.newBlob(decodedData, fileInfo.mimeType, fileInfo.fileName);

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId = file.getId();
    const imageUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;

    const sheet = getSheet();
    sheet.appendRow([
      new Date(), // Timestamp
      userName,   // User
      '',         // Message (empty for image messages)
      imageUrl    // ImageLink
    ]);

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