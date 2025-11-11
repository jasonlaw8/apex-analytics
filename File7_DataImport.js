/**
 * ========================================
 * APEX GOLF ANALYTICS - FILE 7: DATA IMPORT
 * ========================================
 * 
 * THIS FILE CONTAINS:
 * - Import functions for Square Transactions, Items, Staff Timecards, and Apex Bookings
 * - Options for Google Drive folder monitoring or manual paste
 * 
 * Add to menu in File 1 onOpen():
 * .addItem('📥 Import New Data', 'showImportMenu')
 */

// ============================================
// IMPORT MENU
// ============================================

/**
 * Show import options menu
 */
function showImportMenu() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert(
    '📥 Import New Data',
    'Choose import method:\n\n' +
    '1. Google Drive Folder (recommended)\n' +
    '2. Paste from clipboard\n' +
    '3. Setup auto-import folder\n\n' +
    'Which method would you like to use?',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (response == ui.Button.YES) {
    // Option 1: Google Drive folder import
    importFromDriveFolder();
  } else if (response == ui.Button.NO) {
    // Option 2: Manual paste
    showPasteImportDialog();
  } else if (response == ui.Button.CANCEL) {
    // Option 3: Setup auto-import
    setupAutoImportFolder();
  }
}

// ============================================
// OPTION 1: GOOGLE DRIVE FOLDER IMPORT
// ============================================

/**
 * Import files from a designated Google Drive folder
 */
function importFromDriveFolder() {
  var ui = SpreadsheetApp.getUi();
  var userProperties = PropertiesService.getUserProperties();
  var folderId = userProperties.getProperty('importFolderId');
  
  if (!folderId) {
    ui.alert(
      'No Import Folder Set',
      'Please set up your import folder first.\n\n' +
      'Go to: 📊 Apex Analytics > 📥 Import New Data > Setup auto-import folder',
      ui.ButtonSet.OK
    );
    return;
  }
  
  try {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    
    // Organize files by type and find the LATEST of each
    var latestFiles = {
      transactions: null,
      items: null,
      timecards: null,
      bookings: null,
      customers: null
    };
    
    var latestDates = {
      transactions: new Date(0),
      items: new Date(0),
      timecards: new Date(0),
      bookings: new Date(0),
      customers: new Date(0)
    };
    
    // Process each file in the folder
    while (files.hasNext()) {
      var file = files.next();
      var fileName = file.getName().toLowerCase();
      var mimeType = file.getMimeType();
      var lastUpdated = file.getLastUpdated();
      
      // Only process Excel/CSV files
      if (mimeType !== MimeType.MICROSOFT_EXCEL && 
          mimeType !== MimeType.GOOGLE_SHEETS &&
          !fileName.endsWith('.csv') &&
          !fileName.endsWith('.xlsx')) {
        continue;
      }
      
      // Identify file type by name and track latest
      // Check for transactions: "transactions-" or "transactions_"
      if (fileName.indexOf('transaction') >= 0) {
        if (lastUpdated > latestDates.transactions) {
          latestFiles.transactions = file;
          latestDates.transactions = lastUpdated;
        }
      } 
      // Check for items: "items-" or "items_"
      else if (fileName.indexOf('item') >= 0) {
        if (lastUpdated > latestDates.items) {
          latestFiles.items = file;
          latestDates.items = lastUpdated;
        }
      } 
      // Check for customers: "customer" or "customer_round_list"
      else if (fileName.indexOf('customer') >= 0) {
        if (lastUpdated > latestDates.customers) {
          latestFiles.customers = file;
          latestDates.customers = lastUpdated;
        }
      } 
      // Check for timecards/shifts: "timecard", "staff", "payroll", "shifts", or "shift-export"
      else if (fileName.indexOf('timecard') >= 0 || 
               fileName.indexOf('staff') >= 0 || 
               fileName.indexOf('payroll') >= 0 ||
               fileName.indexOf('shift') >= 0) {
        if (lastUpdated > latestDates.timecards) {
          latestFiles.timecards = file;
          latestDates.timecards = lastUpdated;
        }
      } 
      // Check for bookings: "booking", "apex", or "export-" (generic export from Apex)
      else if (fileName.indexOf('booking') >= 0 || 
               fileName.indexOf('apex') >= 0 ||
               (fileName.indexOf('export-') >= 0 && fileName.indexOf('item') < 0 && fileName.indexOf('transaction') < 0)) {
        if (lastUpdated > latestDates.bookings) {
          latestFiles.bookings = file;
          latestDates.bookings = lastUpdated;
        }
      }
    }
    
    var imported = {
      transactions: false,
      items: false,
      timecards: false,
      bookings: false,
      customers: false
    };
    
    // Import the latest file of each type with detailed error handling
    if (latestFiles.transactions) {
      try {
        Logger.log("Importing latest transactions: " + latestFiles.transactions.getName());
        importSquareTransactions(latestFiles.transactions);
        imported.transactions = true;
      } catch (error) {
        Logger.log("ERROR importing transactions: " + error.toString());
        throw new Error("Failed to import Square Transactions from '" + latestFiles.transactions.getName() + "': " + error.toString());
      }
    }

    if (latestFiles.items) {
      try {
        Logger.log("Importing latest items: " + latestFiles.items.getName());
        importSquareItems(latestFiles.items);
        imported.items = true;
      } catch (error) {
        Logger.log("ERROR importing items: " + error.toString());
        throw new Error("Failed to import Square Items from '" + latestFiles.items.getName() + "': " + error.toString());
      }
    }

    if (latestFiles.customers) {
      try {
        Logger.log("Importing latest customers: " + latestFiles.customers.getName());
        importSquareCustomers(latestFiles.customers);
        imported.customers = true;
      } catch (error) {
        Logger.log("ERROR importing customers: " + error.toString());
        throw new Error("Failed to import Square Customers from '" + latestFiles.customers.getName() + "': " + error.toString());
      }
    }

    if (latestFiles.timecards) {
      try {
        Logger.log("Importing latest timecards: " + latestFiles.timecards.getName());
        importStaffTimecards(latestFiles.timecards);
        imported.timecards = true;
      } catch (error) {
        Logger.log("ERROR importing timecards: " + error.toString());
        throw new Error("Failed to import Staff Timecards from '" + latestFiles.timecards.getName() + "': " + error.toString());
      }
    }

    if (latestFiles.bookings) {
      try {
        Logger.log("Importing latest bookings: " + latestFiles.bookings.getName());
        importApexBookings(latestFiles.bookings);
        imported.bookings = true;
      } catch (error) {
        Logger.log("ERROR importing bookings: " + error.toString());
        throw new Error("Failed to import Apex Bookings from '" + latestFiles.bookings.getName() + "': " + error.toString());
      }
    }
    
    // Show results with file names
    var message = '✅ Import Complete!\n\n';
    message += 'Imported:\n';
    if (imported.transactions) {
      message += '• Square Transactions: ✓\n  ' + latestFiles.transactions.getName() + '\n';
    } else {
      message += '• Square Transactions: ✗ (no file found)\n';
    }
    
    if (imported.items) {
      message += '• Square Items: ✓\n  ' + latestFiles.items.getName() + '\n';
    } else {
      message += '• Square Items: ✗ (no file found)\n';
    }
    
    if (imported.customers) {
      message += '• Square Customers: ✓\n  ' + latestFiles.customers.getName() + '\n';
    } else {
      message += '• Square Customers: ✗ (no file found)\n';
    }
    
    if (imported.timecards) {
      message += '• Staff Timecards: ✓\n  ' + latestFiles.timecards.getName() + '\n';
    } else {
      message += '• Staff Timecards: ✗ (no file found)\n';
    }
    
    if (imported.bookings) {
      message += '• Apex Bookings: ✓\n  ' + latestFiles.bookings.getName() + '\n';
    } else {
      message += '• Apex Bookings: ✗ (no file found)\n';
    }
    
    ui.alert('Import Results', message, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log("Import error: " + error.toString());
    ui.alert('Import Error', 'Error: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Setup auto-import folder
 */
function setupAutoImportFolder() {
  var ui = SpreadsheetApp.getUi();
  
  var html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; max-width: 600px; }
        h2 { color: #1a73e8; }
        .instructions { background: #e8f0fe; padding: 15px; border-radius: 8px; margin: 20px 0; }
        .step { margin: 15px 0; padding-left: 10px; }
        .step-num { background: #1a73e8; color: white; border-radius: 50%; padding: 5px 10px; margin-right: 10px; }
        input { width: 100%; padding: 10px; margin: 10px 0; font-size: 14px; border: 2px solid #e0e0e0; border-radius: 4px; }
        button { background: #1a73e8; color: white; border: none; padding: 12px 24px; cursor: pointer; margin: 5px; border-radius: 4px; font-size: 14px; }
        button:hover { background: #1557b0; }
        .example { background: #f8f9fa; padding: 10px; border-left: 4px solid #1a73e8; margin: 10px 0; font-family: monospace; }
        img { max-width: 100%; border: 1px solid #e0e0e0; margin: 10px 0; }
      </style>
    </head>
    <body>
      <h2>📁 Setup Auto-Import Folder</h2>
      
      <div class="instructions">
        <div class="step">
          <span class="step-num">1</span>
          <strong>Create a folder in Google Drive</strong>
          <p>Go to <a href="https://drive.google.com" target="_blank">Google Drive</a> and create a new folder (e.g., "Apex Imports")</p>
        </div>
        
        <div class="step">
          <span class="step-num">2</span>
          <strong>Open the folder and copy the Folder ID from the URL</strong>
          <div class="example">
            Example URL:<br>
            https://drive.google.com/drive/folders/<strong style="color: #e8710a;">1a2B3c4D5e6F7g8H9i0J</strong>
            <br><br>
            The Folder ID is the part after "folders/"
          </div>
        </div>
        
        <div class="step">
          <span class="step-num">3</span>
          <strong>Paste the Folder ID below</strong>
        </div>
      </div>
      
      <label for="folderId"><strong>Google Drive Folder ID:</strong></label>
      <input type="text" id="folderId" placeholder="e.g., 1a2B3c4D5e6F7g8H9i0J">
      
      <button onclick="saveFolderId()">Save Folder</button>
      <button onclick="google.script.host.close()">Cancel</button>
      
      <div id="status"></div>
      
      <script>
        function saveFolderId() {
          var folderId = document.getElementById('folderId').value.trim();
          
          if (!folderId) {
            document.getElementById('status').innerHTML = '<p style="color: red;">Please enter a Folder ID!</p>';
            return;
          }
          
          document.getElementById('status').innerHTML = '<p>Checking folder access...</p>';
          
          google.script.run
            .withSuccessHandler(function(folderName) {
              document.getElementById('status').innerHTML = '<p style="color: green;">✅ Success! Folder: ' + folderName + '</p>';
              setTimeout(function() { google.script.host.close(); }, 2000);
            })
            .withFailureHandler(function(error) {
              document.getElementById('status').innerHTML = '<p style="color: red;">❌ Error: Could not access folder. Make sure the Folder ID is correct and you have access.</p>';
            })
            .saveImportFolder(folderId);
        }
      </script>
    </body>
    </html>
  `).setWidth(700).setHeight(600);
  
  ui.showModalDialog(html, 'Setup Auto-Import Folder');
}

/**
 * Save import folder ID (called from HTML dialog)
 */
function saveImportFolder(folderId) {
  try {
    // Test if folder is accessible
    var folder = DriveApp.getFolderById(folderId);
    
    // Save to user properties
    PropertiesService.getUserProperties().setProperty('importFolderId', folderId);
    
    return folder.getName();
    
  } catch (error) {
    throw new Error('Could not access folder: ' + error.toString());
  }
}

// ============================================
// OPTION 2: MANUAL PASTE IMPORT
// ============================================

/**
 * Show dialog for manual paste import
 */
function showPasteImportDialog() {
  var html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        h2 { color: #1a73e8; }
        select, textarea { width: 100%; margin: 10px 0; padding: 8px; }
        button { background: #1a73e8; color: white; border: none; padding: 10px 20px; cursor: pointer; margin: 5px; }
        button:hover { background: #1557b0; }
        .instructions { background: #e8f0fe; padding: 10px; border-radius: 4px; margin: 10px 0; }
      </style>
    </head>
    <body>
      <h2>📋 Paste Import Data</h2>
      
      <div class="instructions">
        <strong>Instructions:</strong>
        <ol>
          <li>Open your CSV/Excel file in Excel or Google Sheets</li>
          <li>Select ALL data (including headers)</li>
          <li>Copy (Ctrl+C or Cmd+C)</li>
          <li>Paste below and click Import</li>
        </ol>
      </div>
      
      <label for="dataType"><strong>Data Type:</strong></label>
      <select id="dataType">
        <option value="transactions">Square Transactions Export</option>
        <option value="items">Square Items Export</option>
        <option value="customers">Square Customers Export</option>
        <option value="timecards">Staff Timecards</option>
        <option value="bookings">Apex Bookings Export</option>
      </select>
      
      <label for="pasteData"><strong>Paste Data Here:</strong></label>
      <textarea id="pasteData" rows="15" placeholder="Paste your data here..."></textarea>
      
      <button onclick="importData()">Import Data</button>
      <button onclick="google.script.host.close()">Cancel</button>
      
      <div id="status"></div>
      
      <script>
        function importData() {
          var dataType = document.getElementById('dataType').value;
          var data = document.getElementById('pasteData').value;
          
          if (!data.trim()) {
            document.getElementById('status').innerHTML = '<p style="color: red;">Please paste data first!</p>';
            return;
          }
          
          document.getElementById('status').innerHTML = '<p>Importing...</p>';
          
          google.script.run
            .withSuccessHandler(function() {
              document.getElementById('status').innerHTML = '<p style="color: green;">✅ Import successful!</p>';
              setTimeout(function() { google.script.host.close(); }, 2000);
            })
            .withFailureHandler(function(error) {
              document.getElementById('status').innerHTML = '<p style="color: red;">Error: ' + error + '</p>';
            })
            .importPastedData(dataType, data);
        }
      </script>
    </body>
    </html>
  `).setWidth(600).setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Data');
}

/**
 * Import pasted data
 */
function importPastedData(dataType, pastedData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Parse TSV data (tab-separated from copy/paste)
  var rows = pastedData.split('\n');
  var data = [];

  for (var i = 0; i < rows.length; i++) {
    if (rows[i].trim()) {
      data.push(rows[i].split('\t'));
    }
  }

  if (data.length === 0) {
    throw new Error('No data found');
  }

  // Get target sheet
  var sheetName = '';
  switch(dataType) {
    case 'transactions':
      sheetName = 'Square Transactions Export';
      break;
    case 'items':
      sheetName = 'Square Item Detail Export';
      break;
    case 'customers':
      sheetName = 'Square Customer Export';
      break;
    case 'timecards':
      sheetName = 'Staff Timecards';
      break;
    case 'bookings':
      sheetName = 'Apex Bookings Export';
      break;
  }

  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }

  // Clear ALL existing data (including headers)
  if (sheet.getLastRow() > 0) {
    sheet.clear();
  }

  // Paste new data (including headers)
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  Logger.log('Imported ' + (data.length - 1) + ' rows with ' + data[0].length + ' columns into ' + sheetName);
}

// ============================================
// IMPORT HELPER FUNCTIONS
// ============================================

/**
 * Import Square Customers from file
 */
function importSquareCustomers(file) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Square Customer Export');

  Logger.log('Reading file: ' + file.getName());
  var data = readFileData(file);

  if (data.length === 0) {
    Logger.log('No data to import');
    return;
  }

  Logger.log('File has ' + data.length + ' rows and ' + data[0].length + ' columns');
  Logger.log('Sheet "Square Customer Export" currently has ' + sheet.getLastRow() + ' rows and ' + sheet.getLastColumn() + ' columns');

  // Clear all existing data (including headers)
  if (sheet.getLastRow() > 0) {
    Logger.log('Clearing existing sheet data...');
    sheet.clear();
  }

  // Import ALL data including headers from the file
  Logger.log('Writing data to range: 1,1 to ' + data.length + ',' + data[0].length);
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  Logger.log('✓ Successfully imported Square Customers: ' + (data.length - 1) + ' rows with ' + data[0].length + ' columns');
}

/**
 * Import Square Transactions from file
 */
function importSquareTransactions(file) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Square Transactions Export');

  Logger.log('Reading file: ' + file.getName());
  var data = readFileData(file);

  if (data.length === 0) {
    Logger.log('No data to import');
    return;
  }

  Logger.log('File has ' + data.length + ' rows and ' + data[0].length + ' columns');
  Logger.log('Sheet "Square Transactions Export" currently has ' + sheet.getLastRow() + ' rows and ' + sheet.getLastColumn() + ' columns');

  // Clear all existing data (including headers)
  if (sheet.getLastRow() > 0) {
    Logger.log('Clearing existing sheet data...');
    sheet.clear();
  }

  // Import ALL data including headers from the file
  Logger.log('Writing data to range: 1,1 to ' + data.length + ',' + data[0].length);
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  Logger.log('✓ Successfully imported Square Transactions: ' + (data.length - 1) + ' rows with ' + data[0].length + ' columns');
}

/**
 * Import Square Items from file
 */
function importSquareItems(file) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Square Item Detail Export');

  Logger.log('Reading file: ' + file.getName());
  var data = readFileData(file);

  if (data.length === 0) {
    Logger.log('No data to import');
    return;
  }

  Logger.log('File has ' + data.length + ' rows and ' + data[0].length + ' columns');
  Logger.log('Sheet "Square Item Detail Export" currently has ' + sheet.getLastRow() + ' rows and ' + sheet.getLastColumn() + ' columns');

  // Clear all existing data (including headers)
  if (sheet.getLastRow() > 0) {
    Logger.log('Clearing existing sheet data...');
    sheet.clear();
  }

  // Import ALL data including headers from the file
  Logger.log('Writing data to range: 1,1 to ' + data.length + ',' + data[0].length);
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  Logger.log('✓ Successfully imported Square Items: ' + (data.length - 1) + ' rows with ' + data[0].length + ' columns');
}

/**
 * Import Staff Timecards from file
 */
function importStaffTimecards(file) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Staff Timecards');

  Logger.log('Reading file: ' + file.getName());
  var data = readFileData(file);

  if (data.length === 0) {
    Logger.log('No data to import');
    return;
  }

  Logger.log('File has ' + data.length + ' rows and ' + data[0].length + ' columns');
  Logger.log('Sheet "Staff Timecards" currently has ' + sheet.getLastRow() + ' rows and ' + sheet.getLastColumn() + ' columns');

  // Clear all existing data (including headers)
  if (sheet.getLastRow() > 0) {
    Logger.log('Clearing existing sheet data...');
    sheet.clear();
  }

  // Import ALL data including headers from the file
  Logger.log('Writing data to range: 1,1 to ' + data.length + ',' + data[0].length);
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  Logger.log('✓ Successfully imported Staff Timecards: ' + (data.length - 1) + ' rows with ' + data[0].length + ' columns');
}

/**
 * Import Apex Bookings from file
 */
function importApexBookings(file) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Apex Bookings Export');

  Logger.log('Reading file: ' + file.getName());
  var data = readFileData(file);

  if (data.length === 0) {
    Logger.log('No data to import');
    return;
  }

  Logger.log('File has ' + data.length + ' rows and ' + data[0].length + ' columns');
  Logger.log('Sheet "Apex Bookings Export" currently has ' + sheet.getLastRow() + ' rows and ' + sheet.getLastColumn() + ' columns');

  // Clear all existing data (including headers)
  if (sheet.getLastRow() > 0) {
    Logger.log('Clearing existing sheet data...');
    sheet.clear();
  }

  // Import ALL data including headers from the file
  Logger.log('Writing data to range: 1,1 to ' + data.length + ',' + data[0].length);
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  Logger.log('✓ Successfully imported Apex Bookings: ' + (data.length - 1) + ' rows with ' + data[0].length + ' columns');
}

/**
 * Read data from a file (Excel or CSV)
 */
function readFileData(file) {
  var mimeType = file.getMimeType();
  var data = [];
  
  if (mimeType === MimeType.MICROSOFT_EXCEL || mimeType === MimeType.GOOGLE_SHEETS) {
    // Convert to Google Sheets temporarily to read
    var tempSheet = Drive.Files.copy({}, file.getId(), {convert: true});
    var tempSS = SpreadsheetApp.openById(tempSheet.id);
    var sheet = tempSS.getSheets()[0];
    data = sheet.getDataRange().getValues();
    
    // Delete temp file
    Drive.Files.remove(tempSheet.id);
    
  } else if (file.getName().endsWith('.csv')) {
    // Read CSV
    var csvData = file.getBlob().getDataAsString();
    var rows = csvData.split('\n');
    
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].trim()) {
        // Simple CSV parsing (handles basic cases)
        data.push(rows[i].split(','));
      }
    }
  }
  
  return data;
}