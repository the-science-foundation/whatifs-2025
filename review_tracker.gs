
const PROJECT_TITLE_COLUMN = "Project Title [50 chars, excluding spaces]";
const SCIENCE_DISCIPLINE_COLUMN = 'Main Scientific Disciplines of your project (select one or more):'

/**
 * Main function to generate the Review Tracker spreadsheet.
 * @param {string} src_dir_id - The ID of the folder containing submission subfolders.
 * @param {string} src_spreadsheet_id - The ID of the master spreadsheet with submission data.
 * @param {string[]} reviewer_emails - An array of reviewer emails (e.g., ["a@test.com", "b@test.com"]).
 * @param {string} dst_dir_id - The ID of the folder where the new tracker will be saved.
 * @param {string} dst_sheet_name - The name of the Sheet tracker file.
 */
function GenerateReviewTracker(src_dir_id, src_spreadsheet_id, reviewer_emails, dst_dir_id, dst_sheet_name='Review Tracker') {
  
  // 1. Check destination for existing file
  const dstFolder = DriveApp.getFolderById(dst_dir_id);
  const existingFiles = dstFolder.getFilesByName(dst_sheet_name);
  
  if (existingFiles.hasNext()) {
    throw new Error(`Error: A file named '${dst_sheet_name}' already exists in the destination directory.`);
  }

  // 2. Gather Data
  const sourceDataMap = ParseSourceSpreadsheet(src_spreadsheet_id);
  const submissionFolders = ListSubmissionDirs(src_dir_id);
  
  if (submissionFolders.length === 0) {
    console.log("No submission folders found.");
    return;
  }

  // 3. Prepare Spreadsheet Data
  // Headers: Folder, Title, Link, Discipline, Pre-screened, [Reviewers...]
  const headers = [
    "Submission ID", 
    "Project Title", 
    "Submission Link", 
    "Science Discipline", 
    "Pre-screened", 
    ...reviewer_emails.map(email => `Reviewer ${email}`)
  ];

  const outputRows = [];

  submissionFolders.forEach(sub => {
    // Extract the numeric ID from the folder name to find the spreadsheet row
    // Note: The prompt implies the ID in the folder name corresponds to the Row Number directly.
    const rowId = GetSubmissionIdFromName(sub.name);
    
    let projectTitle = "Not Found";
    let discipline = "Not Found";

    if (sourceDataMap[rowId]) {
      projectTitle = sourceDataMap[rowId].title;
      discipline = sourceDataMap[rowId].discipline;
    }

    const row = [
      sub.name,
      projectTitle,
      `=HYPERLINK("${sub.url}", "Open Folder")`, // Clickable link formula
      discipline,
      false, // Checkbox default value (unchecked)
      ...reviewer_emails.map(() => "") // Empty cells for reviewers
    ];

    outputRows.push(row);
  });

  // 4. Create and Format the Spreadsheet
  const ss = SpreadsheetApp.create(dst_sheet_name);
  const sheet = ss.getActiveSheet();

  // Write Headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  
  // Write Data Rows
  if (outputRows.length > 0) {
    sheet.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);
    
    // Insert Checkboxes in Column 5 ("Pre-screened")
    sheet.getRange(2, 5, outputRows.length, 1).insertCheckboxes();
  }

  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Auto-resize columns for readability
  sheet.autoResizeColumns(1, headers.length);

  // 5. Move the file to the destination directory
  MoveFileToFolder(ss.getId(), dst_dir_id);
  
  console.log(`${dst_sheet_name} created successfully.`);
}

/**
 * Parses the source spreadsheet and maps Row Numbers to Data.
 * Returns an object where key = Row Number, value = {title, discipline}.
 */
function ParseSourceSpreadsheet(src_spreadsheet_id) {
  const ss = SpreadsheetApp.openById(src_spreadsheet_id);
  const sheet = ss.getSheets()[0]; // Assuming data is on the first sheet
  const data = sheet.getDataRange().getValues();
  
  if (data.length === 0) return {};

  const headers = data[0];
  
  // Find column indices (case-insensitive search)
  const titleIndex = headers.findIndex(h => h.toString() === PROJECT_TITLE_COLUMN);
  const disciplineIndex = headers.findIndex(h => h.toString() === SCIENCE_DISCIPLINE_COLUMN);

  if (titleIndex === -1 || disciplineIndex === -1) {
    throw new Error("Could not find 'Title' or 'Science Discipline' columns in source spreadsheet.");
  }

  const dataMap = {};

  // Iterate starting from row 1 (skipping header row 0)
  // The prompt says the ID corresponds to the Row Number.
  // In a spreadsheet, Row 1 is headers. If ID is 5, it means the 5th physical row (index 4).
  for (let i = 1; i < data.length; i++) {
    const rowNum = i + 1; // Physical row number in Excel/Sheets
    dataMap[rowNum] = {
      title: data[i][titleIndex],
      discipline: data[i][disciplineIndex]
    };
  }

  return dataMap;
}

/**
 * Scans the source directory for subfolders.
 * Returns a list of objects: {name, url, id}.
 */
function ListSubmissionDirs(src_dir_id) {
  const parentFolder = DriveApp.getFolderById(src_dir_id);
  const folders = parentFolder.getFolders();
  const result = [];

  while (folders.hasNext()) {
    const folder = folders.next();
    result.push({
      name: folder.getName(),
      url: folder.getUrl(),
      id: folder.getId()
    });
  }
  
  // Sort alphabetically by folder name to keep the tracker organized
  result.sort((a, b) => a.name.localeCompare(b.name));
  
  return result;
}

/**
 * Extracts the numeric ID from the submission folder name.
 * Expected format: SUB0000$ID (e.g., SUB00005 -> 5).
 */
function GetSubmissionIdFromName(folderName) {
  // Regex to find the number at the end of the string, ignoring leading zeros
  const match = folderName.match(/SUB0*(\d+)/);
  
  if (match && match[1]) {
    return parseInt(match[1], 10);
  }
  
  // Return a safe fallback or null if format doesn't match
  console.warn(`Could not parse ID from folder: ${folderName}`);
  return null;
}

/**
 * Helper to move a file from root (default creation spot) to destination folder.
 */
function MoveFileToFolder(fileId, targetFolderId) {
  const file = DriveApp.getFileById(fileId);
  const targetFolder = DriveApp.getFolderById(targetFolderId);
  file.moveTo(targetFolder);
}