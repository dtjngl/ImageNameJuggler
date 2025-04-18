function renameAndUpdateImages() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    var spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
    var startCell = firstEmptyCell(sheet) + 1;
    
    var parentFolder = spreadsheetFile.getParents().next();
    var imageFolderName = "Images";
    var imageFolder = null;
    
    var subfolders = parentFolder.getFolders();
    while (subfolders.hasNext()) {
        var folder = subfolders.next();
        if (folder.getName() === imageFolderName) {
            imageFolder = folder;
            break;
        }
    }
    
    if (!imageFolder) {
        logError("Error: The subfolder '" + imageFolderName + "' was not found.");
        return;
    }

    var data = sheet.getDataRange().getValues();
    var files = imageFolder.getFiles();
    var fileMap = {};

    while (files.hasNext()) {
        var file = files.next();
        fileMap[file.getName().trim().toLowerCase()] = file;
    }

    for (var i = 2; i < data.length; i++) {  // Start from first empty row
        var han = data[i][0].toString().trim().toLowerCase();
        var gtin = data[i][1].toString().trim();
        var matchedFiles = [];

        for (var fileName in fileMap) {
            if (fileName.includes(han)) {
                matchedFiles.push(fileName);
            }
        }

        matchedFiles.sort();
        if (matchedFiles.length > 0) { // If at least one image was found
            for (var j = 0; j < matchedFiles.length; j++) {
                sheet.getRange(i + 1, startCell + j).setValue(matchedFiles[j]);  
            }
            
            // **Step: Set processed row to YELLOW**
            sheet.getRange(i + 1, 1, 1, sheet.getLastColumn()).setBackground("#FFFF99");
        }
    }

    return {
        status: "success",
        message: "Images updated.",
        startCell: startCell,
        lastColumn: sheet.getLastColumn()
    };

}


function processSelectedDatasets() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var selectedRanges = sheet.getActiveRangeList();
    var selectedRows = [];

    if (selectedRanges) {
        selectedRanges.getRanges().forEach(range => {
            for (var i = range.getRow(); i <= range.getLastRow(); i++) {
                selectedRows.push(i);
            }
        });
    }

    // If no rows selected, find all rows where column B (images) is empty
    if (selectedRows.length === 0) {
        SpreadsheetApp.getUi().alert("No empty rows found!");
        return;
    }

    // Store selected rows
    PropertiesService.getScriptProperties().setProperty("selectedRows", JSON.stringify(selectedRows));

    // Start processing first row
    processNextRow();

    return { status: "success", message: "Dataset selection started." };

}


function processNextRow() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var selectedRows = JSON.parse(PropertiesService.getScriptProperties().getProperty("selectedRows") || "[]");

    if (selectedRows.length === 0) {
        SpreadsheetApp.getUi().alert("All selected datasets processed!");
        return;
    }

    var rowIndex = selectedRows.shift() - 1; // Get next row and remove from array
    PropertiesService.getScriptProperties().setProperty("selectedRows", JSON.stringify(selectedRows));

    // **Step 1: Set current row to ORANGE (indicating active processing)**
    sheet.getRange(rowIndex + 1, 1, 1, sheet.getLastColumn()).setBackground("#FFA500");

    var han = sheet.getRange(rowIndex + 1, 1).getValue();

    // **Step 2: Open the image selection dialog**
    showImageSelectionDialog(rowIndex, han);
}


function submitImageSelection(rowIndex) {
    try {
        var checkboxes = document.querySelectorAll('input[name="images"]:checked');
        var selectedImages = [];

        checkboxes.forEach(function(checkbox) {
            selectedImages.push(checkbox.value);
        });

        google.script.run.withFailureHandler(onFailure).saveSelectedImages(rowIndex, selectedImages);
    } catch (error) {
        console.log("Error in submitImageSelection: " + error.message);
        google.script.run.logError(error.message);  // Send error back to Apps Script
    }
}


function showImageSelectionDialog(rowIndex, han) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var allSelectedImages = getAllSelectedImages(sheet); // Get ALL selected images from the sheet
    var previousSelection = sheet.getRange(rowIndex + 1, 2, 1, sheet.getLastColumn()).getValues()[0];  

    // Dynamically find the "Images" folder
    var spreadsheetFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
    var parentFolder = spreadsheetFile.getParents().next();
    var imageFolder = null;
    var subfolders = parentFolder.getFolders();

    while (subfolders.hasNext()) {
        var folder = subfolders.next();
        if (folder.getName() === "Images") {
            imageFolder = folder;
            break;
        }
    }

    if (!imageFolder) {
        SpreadsheetApp.getUi().alert("Error: The 'Images' folder was not found in the parent directory.");
        return;
    }

    var images = listImagesInDrive(imageFolder.getId()); // Use dynamically found folder ID

    var htmlOutput = HtmlService.createHtmlOutput(`
        <h3 id="dataset-title">Select Images for ${han}</h3>
        <button type="button" onclick="submitImageSelection()">Save</button>
        <button type="button" onclick="skipDataset()">Skip</button>
        <form id="image-form">
        <br>
        <div id="image-container">
            ${generateImageHTML(images, previousSelection, allSelectedImages)}
        </div>
        </form>
        <script>
            function submitImageSelection() {
                var selectedImages = [];
                document.querySelectorAll('input[name="images"]:checked').forEach(input => {
                    selectedImages.push(input.value);
                });

                google.script.run.withSuccessHandler(function(updatedData) {
                    updateDatasetWithoutReload(updatedData);
                }).withFailureHandler(function(error) {
                    console.error("Failed to save images!", error);
                    alert("Failed to save images! Check logs.");
                }).saveSelectedImages(${rowIndex}, selectedImages);
            }

            function skipDataset() {
                google.script.run.withSuccessHandler(function(updatedData) {
                    updateDatasetWithoutReload(updatedData);
                }).processNextDataset(${rowIndex});
            }

            function updateDatasetWithoutReload(updatedData) {
                if (!updatedData) return;
                
                document.getElementById("dataset-title").innerText = "Select Images for " + updatedData.han;

                // Update checkboxes and images without clearing the entire container
                document.querySelectorAll('input[name="images"]').forEach(input => {
                    var imageName = input.value.toLowerCase();
                    var isSelected = updatedData.previousSelection.includes(imageName);
                    var isDisabled = updatedData.allSelectedImages.includes(imageName);

                    input.checked = isSelected;
                    input.disabled = isDisabled;
                    input.parentElement.style.opacity = isDisabled ? '0.4' : '1';
                });
            }
        </script>
    `);

    SpreadsheetApp.getUi().showSidebar(htmlOutput);
}


function generateImageHTML(images, previousSelection, allSelectedImages) {
    return images.map(image => {
        var imageNameLower = image.name.toLowerCase();
        var isSelectedInThisDataset = previousSelection.includes(image.name);
        var isSelectedInAnyDataset = allSelectedImages.includes(imageNameLower);

        var opacityStyle = isSelectedInAnyDataset ? 'opacity: 0.4;' : ''; 
        var disabledAttribute = isSelectedInAnyDataset ? 'disabled' : '';

        return `
            <label style="${opacityStyle}">
                <input type="checkbox" name="images" value="${image.name}" ${isSelectedInThisDataset ? 'checked' : ''} ${disabledAttribute}>
                <img src="${image.url}" width="90">
            </label><br>
        `;
    }).join('');
}


function getAllSelectedImages(sheet) {
    var data = sheet.getDataRange().getValues(); // Get all values from the sheet
    var selectedImages = new Set(); // Use a Set to avoid duplicates

    for (var i = 1; i < data.length; i++) { // Skip header row
        for (var j = 1; j < data[i].length; j++) { // Skip first column if needed
            if (data[i][j]) {
                var imageName = String(data[i][j]).trim().toLowerCase(); // Ensure it's a string
                selectedImages.add(imageName);
            }
        }
    }
    return Array.from(selectedImages); // Convert Set to Array
}


function saveSelectedImages(rowIndex, selectedImages) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var startCell = firstEmptyCell(sheet) + 1; // Pass the 'sheet' object
    
    if (selectedImages && selectedImages.length > 0) {
        sheet.getRange(rowIndex + 1, startCell, 1, selectedImages.length).setValues([selectedImages]);
    }

    // Move to the next dataset (this now includes coloring logic)
    processNextDataset(rowIndex);
}


function processNextDataset(previousRowIndex) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var selectedRows = JSON.parse(PropertiesService.getScriptProperties().getProperty("selectedRows") || "[]");

    // **Step 1: Set the previous row to YELLOW (processed)**
    if (previousRowIndex >= 0) { // Check if previousRowIndex is valid
        sheet.getRange(previousRowIndex + 1, 1, 1, sheet.getLastColumn()).setBackground("#FFFF99");
    }

    // **Step 2: Check if all datasets are processed**
    if (selectedRows.length === 0) {

        // 🚀 Send command to the sidebar to close itself
        var html = HtmlService.createHtmlOutput("<script>google.script.host.close();</script>");

        return;
    }

    var nextRowIndex = selectedRows.shift() - 1; // Get next row and remove from queue

    if (nextRowIndex >= 0) { // Check if nextRowIndex is valid
        PropertiesService.getScriptProperties().setProperty("selectedRows", JSON.stringify(selectedRows));

        // **Step 3: Highlight the new row in ORANGE (active processing)**
        sheet.getRange(nextRowIndex + 1, 1, 1, sheet.getLastColumn()).setBackground("#FFA500");

        var han = sheet.getRange(nextRowIndex + 1, 1).getValue();
        showImageSelectionDialog(nextRowIndex, han);
    }
}


function logError(errorMessage) {
    Logger.log("Error from client: " + errorMessage);
}


function listImagesInDrive(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var imageData = [];

  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType().startsWith('image/')) {
      var fileName = file.getName();  // Get file name instead of URL
      var fileId = file.getId();
      var previewUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1000"; // Still used for preview
      
      imageData.push({ name: fileName, url: previewUrl });
    }
  }

  return imageData;
}


function getPreviouslySelectedImages() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var allSelectedImages = new Set(); // Use Set to avoid duplicates

    for (var i = 1; i < data.length; i++) {  // Skip headers
        for (var j = 1; j < data[i].length; j++) {  // Skip first column (HAN)
            if (data[i][j]) {
                allSelectedImages.add(data[i][j].toLowerCase());  // Normalize case
            }
        }
    }

    return Array.from(allSelectedImages);  // Return as an array
}


function skipDataset(rowIndex) {
    // Move to the next dataset (same logic as saving, but without modifying data)
    processNextDataset(rowIndex);
}


function renameFilesBasedOnSheet() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    var spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
    var newNameColumn = firstPopulatedAfterEmpty();

    var parentFolder = spreadsheetFile.getParents().next();  
    var imageFolderName = "Images";
    var imageFolder = null;

    // If no populated column is found (i.e., returns -1), do not proceed
    if (newNameColumn === -1) {
        Logger.log("No populated cell found after the first empty cell.");
        return;
    }

    // Locate the "Images" subfolder
    var subfolders = parentFolder.getFolders();
    while (subfolders.hasNext()) {
        var folder = subfolders.next();
        if (folder.getName() === imageFolderName) {
            imageFolder = folder;
            break;
        }
    }

    if (!imageFolder) {
        logError("Error: The subfolder '" + imageFolderName + "' was not found.");
        return;
    }

    var data = sheet.getDataRange().getValues();
    var files = imageFolder.getFiles();
    var fileMap = {};

    // Store all files in the folder by their names (handling case insensitivity)
    while (files.hasNext()) {
        var file = files.next();
        fileMap[file.getName().trim().toLowerCase()] = file;
    }

    // Get selected rows
    var selectedRanges = sheet.getActiveRangeList();
    var selectedRows = [];

    if (selectedRanges) {
        selectedRanges.getRanges().forEach(range => {
            for (var i = range.getRow(); i < range.getLastRow() + 1; i++) {
                selectedRows.push(i);
            }
        });
    }

    if (selectedRows.length === 0) {
        SpreadsheetApp.getUi().alert("No rows selected! Please select rows before running.");
        return;
    }

    var renamedCount = 0;

    selectedRows.forEach(rowIndex => {
        var i = rowIndex - 1;  // Convert to zero-based index
        var gtin = data[i][1].toString().trim(); // Get GTIN from column 2
        var newFileNames = [];
        var suffix = 1;

        for (var j = 2; j < data[i].length; j++) { // Start checking from column 3 onwards
            var oldFileName = data[i][j];
            if (!oldFileName) continue; // Skip empty cells

            oldFileName = String(oldFileName).trim().toLowerCase();

            // Find the file in Drive
            if (fileMap[oldFileName]) {
                var file = fileMap[oldFileName];
                var fileExtension = oldFileName.split('.').pop(); // Get file type (.jpg, .png, etc.)
                var newFileName = gtin + "_" + suffix + "." + fileExtension; // Rename pattern: GTIN_1.jpg

                // Ensure we don't overwrite an existing file with the same name
                while (fileMap[newFileName.toLowerCase()]) {
                    suffix++;
                    newFileName = gtin + "_" + suffix + "." + fileExtension;
                }

                // Rename file
                file.setName(newFileName);
                newFileNames.push(newFileName);
                fileMap[newFileName.toLowerCase()] = file; // Update the map with the new name
                renamedCount++;
                suffix++;
            }
        }

        // Update the sheet with new file names
        if (newFileNames.length > 0) {
            sheet.getRange(rowIndex, newNameColumn, 1, newFileNames.length).setValues([newFileNames]);

            // Highlight row to indicate processing
            sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).setBackground("#FFFF99");
        }
    });

    SpreadsheetApp.getUi().alert(`Renaming completed! ${renamedCount} files updated.`);

    return { status: "success", message: "Files renamed." };

}


function firstPopulatedAfterEmpty() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();

    // Get the values from row 2 (starting from column 1) up to the last column with data
    var range = sheet.getRange(2, 1, 1, sheet.getLastColumn());
    var values = range.getValues();  // Get values of that range

    // Step 1: Find the index of the first empty cell
    var firstEmptyIndex = values[0].indexOf("");  // Get the index of the first empty cell

    if (firstEmptyIndex === -1) {
        Logger.log("No empty cell found in the row.");
        return -1;  // Return -1 if no empty cell is found, indicating no populated cell can be found after
    }

    // Step 2: Look for the first populated cell after the first empty cell
    // Start searching from the first empty cell onwards (inclusive)
    for (var i = firstEmptyIndex; i < values[0].length; i++) {
        if (values[0][i] !== "") {
            var columnNumber = i + 1;  // Adjusting because column numbers are 1-based
            Logger.log("The column number of the first populated cell after the first empty cell is: " + columnNumber);
            return columnNumber;  // Return the column number
        }
    }

    Logger.log("No populated cell found after the first empty cell.");
    return -1;  // Return -1 if no populated cell is found after the empty one
}

function doPost(e) {
  const response = { status: "success", message: "Hello from Apps Script!" };
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader("Access-Control-Allow-Origin", "*");
}

function doGet(e) {
  var output = ContentService.createTextOutput("Web app is running!");
  output.setMimeType(ContentService.MimeType.TEXT);
  // output.setHeader("Access-Control-Allow-Origin", "*");
  return output;
}


function openSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("Image Picker");
  SpreadsheetApp.getUi().showSidebar(html);
}


function processNextDatasetFromExtension(spreadsheetId) {
  const selectedRowsJson = PropertiesService.getScriptProperties().getProperty(`selectedRows_${spreadsheetId}`);
  if (!selectedRowsJson) {
    return ContentService.createTextOutput(JSON.stringify({ finished: true, message: "All datasets processed." })).setMimeType(ContentService.MimeType.JSON).setHeader("Access-Control-Allow-Origin", "*");
  }
  const selectedRows = JSON.parse(selectedRowsJson);
  if (selectedRows.length === 0) {
    PropertiesService.getScriptProperties().deleteProperty(`selectedRows_${spreadsheetId}`);
    return ContentService.createTextOutput(JSON.stringify({ finished: true, message: "All datasets processed." })).setMimeType(ContentService.MimeType.JSON).setHeader("Access-Control-Allow-Origin", "*");
  }
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getActiveSheet();
  const rowIndex = selectedRows.shift();
  PropertiesService.getScriptProperties().setProperty(`selectedRows_${spreadsheetId}`, JSON.stringify(selectedRows));
  sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).setBackground("#FFA500"); // Highlight current row
  const han = sheet.getRange(rowIndex, 1).getValue();
  const spreadsheetFile = DriveApp.getFileById(spreadsheetId);
  const parentFolder = spreadsheetFile.getParents().next();
  const imageFolder = getOrCreateImageFolder(parentFolder);
  if (!imageFolder) {
    return ContentService.createTextOutput(JSON.stringify({ error: "Images folder not found." })).setMimeType(ContentService.MimeType.JSON).setHeader("Access-Control-Allow-Origin", "*");
  }
  const images = listImagesInDrive(imageFolder.getId());
  const previousSelection = sheet.getRange(rowIndex, 2, 1, sheet.getLastColumn() - 1).getValues()[0].filter(String).map(s => s.toLowerCase());
  return ContentService.createTextOutput(JSON.stringify({ han: han, images: images, previousSelection: previousSelection, rowIndex: rowIndex, finished: false })).setMimeType(ContentService.MimeType.JSON).setHeader("Access-Control-Allow-Origin", "*");
}


function getOrCreateImageFolder(parentFolder) {
  var subfolders = parentFolder.getFolders();
  while (subfolders.hasNext()) {
    var folder = subfolders.next();
    if (folder.getName() === "Images") {
      return folder;
    }
  }
  return parentFolder.createFolder("Images"); // Optionally create if it doesn't exist
}


function firstEmptyCell (sheet) {
    var range = sheet.getRange(2, 1, 1, sheet.getLastColumn());
    var values = range.getValues();
    var empty_cell = values[0].indexOf("");
    return empty_cell === -1 ? sheet.getLastColumn() : empty_cell;
}


