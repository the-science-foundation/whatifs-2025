
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
    ...reviewer_emails.map(email => `Request review from ${email}`)
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
      ...reviewer_emails.map(() => false) // Check to request review
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
    sheet.getRange(2, 6, outputRows.length, reviewer_emails.length).insertCheckboxes();
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

function getReviewerAssignments(sheetId) {
  // 1. Open the Spreadsheet and access the data
  // We assume the data is on the first sheet. Change [0] to getSheetByName("Name") if needed.
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheets()[0]; 
  const data = sheet.getDataRange().getValues();

  // 2. Separate headers from the rest of the data
  if (data.length === 0) return [];
  const headers = data[0];
  const rows = data.slice(1);

  // 3. Identify key column indices
  const submissionIdHeader = "Submission ID";
  const submissionIdIndex = headers.indexOf(submissionIdHeader);

  if (submissionIdIndex === -1) {
    throw new Error(`Column "${submissionIdHeader}" not found.`);
  }

  // 4. map out which columns are Reviewer columns and extract their emails
  // Regex looks for "Request review from " followed by any characters (the email)
  const headerRegex = /^Request review from\s+(.+)$/i;
  
  const reviewerColumns = [];
  headers.forEach((header, index) => {
    const match = header.match(headerRegex);
    if (match) {
      reviewerColumns.push({
        index: index,
        email: match[1].trim() // The captured email group
      });
    }
  });

  // 5. Iterate through rows and build the result list
  const results = [];

  rows.forEach((row) => {
    const submissionId = row[submissionIdIndex];

    // Skip rows where Submission Id is empty
    if (!submissionId) return;

    reviewerColumns.forEach((col) => {
      // Check if the checkbox is checked (true)
      // We explicitly check for strict true to avoid false positives on non-empty strings
      if (row[col.index] === true) {
        results.push({
          submission_id: submissionId,
          reviewer_email: col.email
        });
      }
    });
  });

  // 6. Return or Log the results
  console.log(`Created ${results.length} review assignments`); // Useful for debugging in the editor
  return results;
}


function generate_review_tasks(review_entries, root_dir_id, stop_on_error=true) {
  let rootFolder = DriveApp.getFolderById(root_dir_id);

  let reviewTasks = new Map();
  for (let entry of review_entries) {
    const folderName = entry.submission_id;
    const reviewerEmail = entry.reviewer_email;
    if (!folderName || !reviewerEmail) {
      if (stop_on_error) {
        throw new Error(`Skipping row with missing data: Folder='${folderName}', Email='${reviewerEmail}'`);
      } else {
        Logger.log(`Skipping row with missing data: Folder='${folderName}', Email='${reviewerEmail}'`);
        continue;
      }
    }

    // Find the subfolder within the root directory
    const folderIter = rootFolder.getFoldersByName(folderName);
    if (!folderIter.hasNext()) {
      if (stop_on_error) {
        throw new Error(`Folder not found: '${folderName}' in root dir '${rootFolder.getName()}'`);
      } else {
        Logger.log(`Folder not found: '${folderName}' in root dir '${rootFolder.getName()}'`);
        continue; // Skip this row
      }
    }
    const folder = folderIter.next();

    // Find the "Description" Google Doc inside that folder
    const fileIter = folder.getFilesByName("Description");
    if (!fileIter.hasNext()) {
      if (stop_on_error) {
        throw new Error(`'Description' doc not found in folder '${folderName}'`)
      } else {
        Logger.log(`'Description' doc not found in folder '${folderName}'`);
        continue; // Skip this row
      }
    }
    const descriptionDoc = fileIter.next();
    
    // Ensure the file is actually a Google Doc
    if (descriptionDoc.getMimeType() !== MimeType.GOOGLE_DOCS) {
      if (stop_on_error) {
        throw new Error(`File 'Description' in '${folderName}' is not a Google Doc. Skipping.`);
      } else {
        Logger.log(`File 'Description' in '${folderName}' is not a Google Doc. Skipping.`);
        continue;
      }
    }

    // Get the document's URL
    const docUrl = descriptionDoc.getUrl();

    if (!reviewTasks.has(reviewerEmail)) {
      reviewTasks.set(reviewerEmail, []);
    }
    reviewTasks.get(reviewerEmail).push({
      folderName: folderName,
      url: docUrl
    });
  }
  return reviewTasks;
}


function do_send_review_emails(reviewTasks, dry_run=false) {
  const keysIterator = reviewTasks.keys();
  for (let email of keysIterator) {
    Logger.log(`preparing email to ${email}`);
    const tasks = reviewTasks.get(email);
    
    if (tasks.length === 0) continue; // Skip if a reviewer has no valid tasks

    const emailBody = generate_review_remainder_email_body(tasks);
    Logger.log(`Email body ${emailBody}`);
    if (dry_run) {
      continue;
    }
    // Send the email
    const subject = "Reminder: Proposals to Review";
    try {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: emailBody // Use htmlBody to render the links correctly
      });
      Logger.log(`Successfully sent email to ${email} with ${tasks.length} link(s).`);
    } catch (e) {
      Logger.log(`Failed to send email to ${email}: ${e}`);
    }
  }
}


function assign_reviewers(review_entries, root_dir_id, stop_on_error=true) {
  Logger.log("Starting permission processing...");
  let rootFolder = DriveApp.getFolderById(root_dir_id);
  for (let entry of review_entries) {
    const folder_name = entry.submission_id;
    const reviewer_email = entry.reviewer_email;
    Logger.log(`Assigning view permission for ${folder_name}, ${reviewer_email}`)

    const folderIterator = rootFolder.getFoldersByName(folder_name);

    if (!folderIterator.hasNext()) {
      if (stop_on_error) {
        throw new Error(`Folder not found with name "${folder_name}" in the root directory.`);
      } else {
        Logger.log(`Folder not found with name "${folder_name}" in the root directory. Skipping this entry.`);
        continue;
      }
    }

    const folder = folderIterator.next();
    // Log a warning if multiple folders with the same name exist
    if (folderIterator.hasNext()) {
      if (stop_on_error) {
        throw new Error(`Multiple folders found with name "${folder_name}".`);
      } else {
        Logger.log(`Multiple folders found with name "${folder_name}". Applying permission to the first one found (ID: ${folder.getId()}).`);
      }
    }

    try {
      // The `folder.addViewer` method below sends an email each time the access is granted, even when the user already has the read access.
      //folder.addViewer(reviewer_email);
      // The method below adds viewer access without sending a email notification.
      // This requires Drive advanced service v2 - enable in the Resources menu in the Script Editor
      Drive.Permissions.insert({'role': 'reader', 'type': 'user', 'value': reviewer_email}, folder.getId(), {'sendNotificationEmails': 'false'});
      Logger.log(`Granted 'View' access to "${folder_name}" for ${reviewer_email}.`);
    } catch (e) {
      if (stop_on_error) {
        throw new Error(`Failed to grant access for ${reviewer_email} to "${folder_name}". Error: ${e.message}`);
      } else {
        Logger.log(`FAILED to grant access for ${reviewer_email} to "${folder_name}". Error: ${e.message}`);
      }
    }
  }
  Logger.log("Permission processing complete.");
}

function send_review_remainder_emails(sheet_id, root_dir_id, stop_on_error=true, dry_run=false) {
  Logger.log(`Starting permission processing for sheet "${sheet_id}, root dir id ${root_dir_id}".`);
  let entries = getReviewerAssignments(sheet_id);
  assign_reviewers(entries, root_dir_id, stop_on_error)
  let reviewTasks = generate_review_tasks(entries, root_dir_id, stop_on_error);
  Logger.log(`Generated ${reviewTasks.size} review tasks`)
  do_send_review_emails(reviewTasks, dry_run);
}