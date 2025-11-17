// Usage:
// 1. Copy this script as a file in your Appscript project. You can name it any way you like e.g. submission_sync_tools.gs
// 2. Create another "runner" script file in the same project - it can use any function from the first submission_sync_tools.gs file directly without importing it.
//    See: https://stackoverflow.com/questions/72843003/how-to-reference-one-apps-script-file-from-another
// 3. Add the following invocation to the runner script file - make sure to replace SOURCE_SHEET_ID and TARGET_FOLDER_ID with actual drive ids:
// function runSync() {
//   // TARGET_FOLDER_ID - you can see this in the browser address bar when you navigate into the dir. 
//   SheetSyncer(SOURCE_SHEET_ID, TARGET_FOLDER_ID);
// }


// Sheet column names
// TODO: replace with the correct ones if needed
const PROJECT_TITLE_COLUMN_NAME = "Project Title";
const SUBMITTER_NAME_COLUMN_NAME = "Name";
const CAREER_STAGE_COLUMN_NAME = 'Career Stage';
const FIELDS_OF_SCIENCE_COLUMN_NAME = 'Fields of Science'
const PROJECT_GRAPHIC_COLUMN_NAME = 'Project Graphic';
const SHORT_DESCRIPTION_COLUMN_NAME = 'Short Description';
const MOTIVATION_COLUMN_NAME = 'Motivation';
const PROJECT_VIDEO_COLUMN_NAME = 'Video';


// Main function to sync Google Sheet records to Google Drive folders.
// For each row, it copies media files and creates a summary document.
function SheetSyncer(sourceSheetId, targetFolderId) {
  const ss = SpreadsheetApp.openById(sourceSheetId);
  const sheet = ss.getSheets()[0]; // Assumes data is on the first sheet
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const parentFolder = DriveApp.getFolderById(targetFolderId);

  const projectNameIdx = header.indexOf(PROJECT_TITLE_COLUMN_NAME);
  const submitterNameIdx = header.indexOf(SUBMITTER_NAME_COLUMN_NAME);

  if (projectNameIdx === -1 || submitterNameIdx === -1) {
    throw new Error("A required column (Project Name or Submitter Name) is missing from the sheet.");
  }

  // Start from 1 to skip the header row.
  for (let rowNum = 1; rowNum < data.length; rowNum++) {
    try {
      const row = data[rowNum];

      const submission_id = 'SUB' + String(rowNum + 1).padStart(6, '0');
      const directory_name = submission_id;

      const existingFolders = parentFolder.getFoldersByName(directory_name);
      if (existingFolders.hasNext()) {
        Logger.log(`Skipping: Folder "${directory_name}" already exists.`);
        continue; // Skip to the next record
      }

      Logger.log(`Processing: Creating folder for "${directory_name}"...`);

      const projectFolder = parentFolder.createFolder(directory_name);
      const projectFolderId = projectFolder.getId();
    
      const projectName = row[projectNameIdx];
      const submitterName = row[submitterNameIdx];
      const submitterCareerStage = row[header.indexOf(CAREER_STAGE_COLUMN_NAME)];
      const fieldOfScience = row[header.indexOf(FIELDS_OF_SCIENCE_COLUMN_NAME)];
      const mainImageUrl = row[header.indexOf(PROJECT_GRAPHIC_COLUMN_NAME)];
      Logger.log(`submitterName: ${submitterName}, fieldOfScience: ${fieldOfScience}, mainImageUrl: ${mainImageUrl}`);
      const mainImageId = getFileIdFromUrl(mainImageUrl);
      Logger.log(`Extracted main image id ${mainImageId} from url ${mainImageUrl}`);
      const mainVideoUrl = row[header.indexOf(PROJECT_VIDEO_COLUMN_NAME)];
      var copiedVideoId = null;
      if (mainVideoUrl) {
        const mainVideoId = getFileIdFromUrl(mainVideoUrl);
        copiedVideoId = copy_supplementary_file(mainVideoId, projectFolder);
      }
      const description = row[header.indexOf(SHORT_DESCRIPTION_COLUMN_NAME)];
      const motivation = row[header.indexOf(MOTIVATION_COLUMN_NAME)];
      //const docName = projectName;
      const docName = "Description";
      const copiedImageId = copy_supplementary_file(mainImageId, projectFolder);

      createProposalDescriptionDoc(projectFolderId, docName, submission_id, projectName, submitterName, submitterCareerStage, fieldOfScience, mainImageId, description, motivation, copiedVideoId);
      
      Logger.log(`Created summary document for "${projectName}", ${submission_id}.`);
    } catch (e) {
      Logger.log(`Failed to process submission sheet row ${rowNum}. Error: ${e.message}`);
    }
  }
  Logger.log('SheetSyncer process finished.');
}


function copy_supplementary_file(file_id, dst_dir) {
  const file = DriveApp.getFileById(file_id);
  let fileName = file.getName(); // This is a modified file name with ` - {First Name} {Last Name}` suffix.
  const newFile = file.makeCopy(fileName, dst_dir);
  return newFile.getId();
}

// A helper function to extract a Google Drive file ID from various URL formats.
function getFileIdFromUrl(url) {
  const regex = /\/d\/([a-zA-Z0-9_-]+)|id=([a-zA-Z0-9_-]+)/;
  const match = url.match(regex);
  return match ? (match[1] || match[2]) : null;
}


function createProposalDescriptionDoc(parentFolderId, docName, submission_id, projectTitle, submitterName, submitterCareerStage, fieldOfScience, mainImageId, description, motivation, videoId) {
  const parentFolder = DriveApp.getFolderById(parentFolderId);
  const doc = DocumentApp.create(docName);
  const docFile = DriveApp.getFileById(doc.getId());
  
  // Move the new document from root to the parent folder
  parentFolder.addFile(docFile);
  DriveApp.getRootFolder().removeFile(docFile);
  const body = doc.getBody();
  
  // Add content to the document
  body.appendParagraph(projectTitle).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Submission ID: ${submission_id}`).setBold(true);
  body.appendParagraph('').setBold(false); // Add a blank line for spacing
  body.appendParagraph(`Author: ${submitterName}`);
  body.appendParagraph(`Career Stage: ${submitterCareerStage}`);
  body.appendParagraph(''); // Add a blank line for spacing
  body.appendParagraph(`Fields/Subfields of Science Involved: ${fieldOfScience}`);
  body.appendParagraph(''); // Add a blank line for spacing

  insert_image_into_doc(mainImageId, body);

  body.appendParagraph(''); // Add a blank line for spacing
  body.appendParagraph('Description').setHeading(DocumentApp.ParagraphHeading.HEADING4);
  body.appendParagraph(description);
  body.appendParagraph('Motivation').setHeading(DocumentApp.ParagraphHeading.HEADING4);
  body.appendParagraph(motivation);
  body.appendParagraph(''); // Add a blank line for spacing  
  if (videoId) {
    const videoFile = DriveApp.getFileById(videoId);
    const videoUrl = videoFile.getUrl();
    var videoParagraph = body.appendParagraph('Link to the Video');
    videoParagraph.setLinkUrl(videoUrl);
  }
  doc.saveAndClose();
}


function insert_image_into_doc(image_id, dst_doc_body) {
  // This prevents an error from inserting an oversized image, more details here:
  // https://stackoverflow.com/questions/54695708/getblob-causing-invalid-image-data-error-google-apps-script
  let url = "https://www.googleapis.com/drive/v3/files/" + image_id + "?fields=thumbnailLink&access_token=" + ScriptApp.getOAuthToken();
  let origThumbnailLink = JSON.parse(UrlFetchApp.fetch(url).getContentText()).thumbnailLink;
  Logger.log(`Original thumbnail link, ${origThumbnailLink}.`);
  let dst_width = 500;
  let resizedThumbnailLink = origThumbnailLink.replace(/=s\d+/, "=s" + dst_width);
  let blob = UrlFetchApp.fetch(resizedThumbnailLink).getBlob();
  if (dst_doc_body) {
    dst_doc_body.appendImage(blob)
  }
}
