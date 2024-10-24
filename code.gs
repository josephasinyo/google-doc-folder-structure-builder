/**
 * Adds a custom menu to the Google Document UI when the document is opened.
 */
function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('🪄 Import Folder Structure')
    .addItem("Run", 'importFolderStructure')
    .addToUi();
}

/**
 * Prompts the user to input a Google Drive folder URL, retrieves the folder hierarchy, 
 * and inserts it into the active Google Document.
 */
function importFolderStructure() {
  var ui = DocumentApp.getUi();
  var response = ui.prompt('Enter the URL of the Google Drive folder:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    var folderUrl = response.getResponseText();
    var folderId = getIdFromUrl(folderUrl);

    if (folderId) {
      var emojiResponse = ui.prompt('Do you want to include emojis for file types? (yes/no)', ui.ButtonSet.OK_CANCEL);
      var includeEmojis = (emojiResponse.getResponseText().toLowerCase() === 'yes');

      var folder = DriveApp.getFolderById(folderId);
      var hierarchy = generateFolderHierarchy(folder, 0, includeEmojis); // Start with depth 0

      var doc = DocumentApp.getActiveDocument();
      var body = doc.getBody();

      hierarchy.forEach(function(item) {
        var name = item.name;
        var url = item.url;
        var isFolder = item.isFolder;
        var depth = item.depth;
        var emoji = includeEmojis ? item.emoji : '';

        var indent = '';
        for (var i = 0; i < depth; i++) {
          indent += '\t'; // Using tab for indentation
        }

        var paragraph = body.appendParagraph(indent + emoji + ' ' + name);

        var text = paragraph.editAsText();

        if (isFolder) {
          text.setBold(true);
          text.setFontSize(12); 
          text.setLinkUrl(url); 
        } else {
          text.setBold(false);
          text.setFontSize(11); 
          text.setLinkUrl(url); 
        }

        paragraph.setUnderline(false);
      });

    } else {
      ui.alert('Invalid folder URL. Please enter a valid Google Drive folder URL.');
    }
  }
}

/**
 * Generates a folder and file hierarchy recursively.
 * 
 * @param {Folder} folder - The Google Drive folder to generate hierarchy from.
 * @param {number} depth - The current depth of the hierarchy (used for indentation).
 * @param {boolean} includeEmojis - Whether to include emojis for file types.
 * @return {Array} - An array of objects representing the folder/file hierarchy.
 */
function generateFolderHierarchy(folder, depth, includeEmojis) {
  var hierarchy = [];
  var contents = folder.getFiles();

  while (contents.hasNext()) {
    var file = contents.next();
    var fileType = getFileType(file);
    var emoji = includeEmojis ? getEmojiForType(fileType) : '';
    hierarchy.push({ name: file.getName(), url: file.getUrl(), isFolder: false, depth: depth, emoji: emoji });
  }

  var subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    var subfolderHierarchy = generateFolderHierarchy(subfolder, depth + 1, includeEmojis);
    hierarchy.push({ name: subfolder.getName(), url: subfolder.getUrl(), isFolder: true, depth: depth, emoji: includeEmojis ? '📁' : '' });
    hierarchy = hierarchy.concat(subfolderHierarchy);
  }

  return hierarchy;
}

/**
 * Extracts the folder or file ID from a given Google Drive URL.
 * 
 * @param {string} url - The Google Drive URL.
 * @return {string|null} - The extracted folder/file ID or null if not found.
 */
function getIdFromUrl(url) {
  var id;
  if (url.indexOf('folders/') !== -1) {
    id = url.match(/[^/]+$/)[0];
  } else {
    id = url.match(/[-\w]{25,}/);
  }
  return id;
}

/**
 * Determines the type of file based on its MIME type.
 * 
 * @param {File} file - The Google Drive file to check.
 * @return {string} - A string representing the file type.
 */
function getFileType(file) {
  var mimeType = file.getMimeType();
  switch (mimeType) {
    case MimeType.GOOGLE_SHEETS:
      return "Spreadsheet";
    case MimeType.GOOGLE_DOCS:
      return "Document";
    case MimeType.GOOGLE_SLIDES:
      return "Presentation";
    case MimeType.GOOGLE_FORMS:
      return "Form";
    case MimeType.GOOGLE_SITES:
      return "Site";
    case MimeType.GOOGLE_JAMBOARD:
      return "Jam";
    case MimeType.GOOGLE_APPS_SCRIPT:
      return "Script";
    default:
      return "Other";
  }
}

/**
 * Returns an emoji representing the file type.
 * 
 * @param {string} fileType - The type of the file (e.g., "Document", "Spreadsheet").
 * @return {string} - An emoji representing the file type.
 */
function getEmojiForType(fileType) {
  switch (fileType) {
    case "Spreadsheet":
      return "🟢";
    case "Document":
      return "🔵";
    case "Presentation":
      return "🟠";
    case "Form":
      return "🟣";
    case "Site":
      return "🔘";
    case "Jam":
      return "🔴";
    case "Script":
      return "🟡";
    default:
      return "📄";
  }
}
