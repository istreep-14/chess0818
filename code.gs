/**
 * Chess.com Archives Manager for Google Sheets
 * Fetches monthly game archives and creates organized sheets
 */

// Configuration - UPDATE THIS WITH YOUR CHESS.COM USERNAME
const CHESS_USERNAME = 'ians141'; // Replace with your Chess.com username
const ARCHIVES_SHEET_NAME = 'Chess Archives';

/**
 * Prompt user to enter archive name for single archive population
 */
function promptPopulateSingleArchive() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Populate Archive Games', 
    'Enter the archive name (format: YYYY-MM, e.g., 2023-08):', 
    ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const archiveName = response.getResponseText().trim();
    if (archiveName) {
      populateArchiveGames(archiveName);
    } else {
      ui.alert('Please enter a valid archive name.');
    }
  }
}

/**
 * Get existing archive names from the sheet to avoid duplicates
 */
function getExistingArchiveNames(sheet) {
  const existingArchives = new Set();
  
  if (sheet.getLastRow() > 1) {
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
    const values = dataRange.getValues();
    
    values.forEach(row => {
      if (row[0]) {
        existingArchives.add(row[0].toString());
      }
    });
  }
  
  return existingArchives;
}

/**
 * Main function to fetch archives and create sheets
 */
function createChessArchives() {
  try {
    console.log('Starting Chess.com archives creation...');
    
    // Get or create the main archives sheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let archivesSheet = getOrCreateSheet(spreadsheet, ARCHIVES_SHEET_NAME);
    
    // Set up headers only if sheet is empty
    if (archivesSheet.getLastRow() === 0) {
      setupArchivesHeaders(archivesSheet);
    }
    
    // Fetch archives from Chess.com API
    const archives = fetchChessArchives();
    
    if (archives.length === 0) {
      throw new Error('No archives found. Check your username.');
    }
    
    console.log(`Found ${archives.length} archives`);
    
    // Get existing archive names to avoid duplicates
    const existingArchives = getExistingArchiveNames(archivesSheet);
    
    // Process each archive
    const newArchiveData = [];
    let newSheetsCount = 0;
    
    for (let i = 0; i < archives.length; i++) {
      const archiveUrl = archives[i];
      const archiveInfo = parseArchiveUrl(archiveUrl);
      
      if (archiveInfo) {
        // Skip if this archive already exists
        if (existingArchives.has(archiveInfo.name)) {
          console.log(`Skipping existing archive: ${archiveInfo.name}`);
          continue;
        }
        
        // Create sheet for this archive
        const sheetName = archiveInfo.sheetName;
        const sheet = getOrCreateSheet(spreadsheet, sheetName);
        
        // Hide the sheet
        sheet.hideSheet();
        
        // Create hyperlink to the sheet
        const sheetUrl = `#gid=${sheet.getSheetId()}`;
        const hyperlink = `=HYPERLINK("${sheetUrl}", "${sheetName}")`;
        
        // Add to new archive data
        newArchiveData.push([
          archiveInfo.name,        // 2023-08
          archiveInfo.monthNumber, // 8  
          archiveInfo.year,        // 23
          hyperlink,               // Sheet link
          sheetUrl,                // Sheet URL (#gid=12345)
          ''                       // Last Updated (empty initially)
        ]);
        
        newSheetsCount++;
        console.log(`Created and hid sheet: ${sheetName}`);
      }
    }
    
    // Append new archive data to the main sheet
    if (newArchiveData.length > 0) {
      const lastRow = archivesSheet.getLastRow();
      const range = archivesSheet.getRange(lastRow + 1, 1, newArchiveData.length, 6);
      range.setValues(newArchiveData);
    }
    
    console.log('Chess archives creation completed successfully!');
    SpreadsheetApp.getUi().alert(`Success! Added ${newSheetsCount} new archive sheets (${newArchiveData.length} new entries). Check the "${ARCHIVES_SHEET_NAME}" sheet for the complete list.`);
    
  } catch (error) {
    console.error('Error creating chess archives:', error);
    SpreadsheetApp.getUi().alert(`Error: Failed to create archives: ${error.message}`);
  }
}

/**
 * Fetch archives from Chess.com API
 */
function fetchChessArchives() {
  const url = `https://api.chess.com/pub/player/${CHESS_USERNAME}/games/archives`;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'User-Agent': 'Google Apps Script Chess Archives'
      }
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`API request failed with status: ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    return data.archives || [];
    
  } catch (error) {
    throw new Error(`Failed to fetch archives: ${error.message}`);
  }
}

/**
 * Parse archive URL to extract date information
 */
function parseArchiveUrl(archiveUrl) {
  // Example URL: https://api.chess.com/pub/player/username/games/2023/08
  const match = archiveUrl.match(/\/games\/(\d{4})\/(\d{2})$/);
  
  if (!match) {
    console.warn(`Could not parse archive URL: ${archiveUrl}`);
    return null;
  }
  
  const year = match[1];
  const month = match[2];
  
  // Convert to readable formats
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                     'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const monthName = monthNames[parseInt(month) - 1];
  const shortYear = year.slice(2); // 2023 -> 23
  
  return {
    name: `${year}-${month}`,           // 2023-08
    monthNumber: parseInt(month),       // 8
    year: shortYear,                    // 23
    sheetName: `a${month}${shortYear}`  // a0823
  };
}

/**
 * Get existing sheet or create new one
 */
function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    console.log(`Created new sheet: ${sheetName}`);
  }
  
  return sheet;
}

/**
 * Set up headers for the archives sheet (no formatting)
 */
function setupArchivesHeaders(sheet) {
  const headers = [['Name', 'Month', 'Year', 'Sheet', 'URL', 'Last Updated']];
  sheet.getRange(1, 1, 1, 6).setValues(headers);
}

/**
 * Format the archives sheet for better readability
 */
function formatArchivesSheet(sheet, dataRows) {
  // Auto-resize columns
  sheet.autoResizeColumns(1, 4);
  
  // Add borders
  const dataRange = sheet.getRange(1, 1, dataRows + 1, 4);
  dataRange.setBorder(true, true, true, true, true, true);
  
  // Alternate row colors for better readability
  for (let i = 2; i <= dataRows + 1; i++) {
    if (i % 2 === 0) {
      sheet.getRange(i, 1, 1, 4).setBackground('#f8f9fa');
    }
  }
  
  // Center align year and month columns
  sheet.getRange(2, 2, dataRows, 2).setHorizontalAlignment('center');
}

/**
 * Populate games data for a specific archive month
 * @param {string} archiveName - Archive name (e.g., "2023-08")
 */
function populateArchiveGames(archiveName) {
  try {
    console.log(`Starting to populate games for archive: ${archiveName}`);
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Parse the archive name to get year and month
    const match = archiveName.match(/^(\d{4})-(\d{2})$/);
    if (!match) {
      throw new Error(`Invalid archive name format: ${archiveName}. Expected format: YYYY-MM`);
    }
    
    const year = match[1];
    const month = match[2];
    const shortYear = year.slice(2);
    
    // Get the corresponding sheet
    const sheetName = `a${month}${shortYear}`;
    const archiveSheet = spreadsheet.getSheetByName(sheetName);
    
    if (!archiveSheet) {
      throw new Error(`Sheet ${sheetName} not found. Please create archives first.`);
    }
    
    // Fetch games for this specific month
    const gamesUrl = `https://api.chess.com/pub/player/${CHESS_USERNAME}/games/${year}/${month}`;
    console.log(`Fetching games from: ${gamesUrl}`);
    
    const response = UrlFetchApp.fetch(gamesUrl, {
      method: 'GET',
      headers: {
        'User-Agent': 'Google Apps Script Chess Games'
      }
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`Failed to fetch games: HTTP ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    const games = data.games || [];
    
    console.log(`Found ${games.length} games for ${archiveName}`);
    
    // Clear sheet and add comprehensive headers
    archiveSheet.clear();
    const headers = [[
      'White Username', 'White Rating', 'White Result', 'White Profile',
      'Black Username', 'Black Rating', 'Black Result', 'Black Profile',
      'White Accuracy', 'Black Accuracy', 'URL', 'FEN', 'PGN',
      'Start Time', 'End Time', 'Time Control', 'Rules', 'ECO', 'Tournament', 'Match'
    ]];
    archiveSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    
    // Process games data
    const gameData = games.map(game => {
      // Parse dates
      const startDate = game.start_time ? new Date(game.start_time * 1000) : '';
      const endDate = game.end_time ? new Date(game.end_time * 1000) : '';
      
      // Extract white player information
      const whiteUsername = game.white?.username || '';
      const whiteRating = game.white?.rating || '';
      const whiteResult = game.white?.result || '';
      const whiteProfile = game.white?.['@id'] || '';
      
      // Extract black player information
      const blackUsername = game.black?.username || '';
      const blackRating = game.black?.rating || '';
      const blackResult = game.black?.result || '';
      const blackProfile = game.black?.['@id'] || '';
      
      // Extract accuracies
      const whiteAccuracy = game.accuracies?.white || '';
      const blackAccuracy = game.accuracies?.black || '';
      
      // Extract game details
      const url = game.url || '';
      const fen = game.fen || '';
      const pgn = game.pgn || '';
      const timeControl = game.time_control || '';
      const rules = game.rules || '';
      const eco = game.eco || '';
      const tournament = game.tournament || '';
      const match = game.match || '';
      
      return [
        whiteUsername,     // White Username
        whiteRating,       // White Rating
        whiteResult,       // White Result
        whiteProfile,      // White Profile URL
        blackUsername,     // Black Username
        blackRating,       // Black Rating
        blackResult,       // Black Result
        blackProfile,      // Black Profile URL
        whiteAccuracy,     // White Accuracy
        blackAccuracy,     // Black Accuracy
        url,               // Game URL
        fen,               // Final FEN
        pgn,               // PGN
        startDate,         // Start Time
        endDate,           // End Time
        timeControl,       // Time Control
        rules,             // Rules/Variant
        eco,               // ECO Opening URL
        tournament,        // Tournament URL
        match              // Match URL
      ];
    });
    
    // Write games data to sheet
    if (gameData.length > 0) {
      const dataRange = archiveSheet.getRange(2, 1, gameData.length, gameData[0].length);
      dataRange.setValues(gameData);
    }
    
    // Update the timestamp in the main archives sheet
    updateArchiveTimestamp(archiveName);
    
    console.log(`Successfully populated ${games.length} games for ${archiveName}`);
    SpreadsheetApp.getUi().alert(`Success! Populated ${games.length} games for archive ${archiveName}.`);
    
  } catch (error) {
    console.error(`Error populating games for ${archiveName}:`, error);
    SpreadsheetApp.getUi().alert(`Error: Failed to populate games for ${archiveName}: ${error.message}`);
  }
}

/**
 * Update the timestamp for a specific archive in the main archives sheet
 * @param {string} archiveName - Archive name (e.g., "2023-08")
 */
function updateArchiveTimestamp(archiveName) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const archivesSheet = spreadsheet.getSheetByName(ARCHIVES_SHEET_NAME);
    
    if (!archivesSheet) {
      console.warn('Archives sheet not found');
      return;
    }
    
    const lastRow = archivesSheet.getLastRow();
    if (lastRow <= 1) {
      console.warn('No archive data found');
      return;
    }
    
    // Find the row with matching archive name
    const nameRange = archivesSheet.getRange(2, 1, lastRow - 1, 1);
    const names = nameRange.getValues();
    
    for (let i = 0; i < names.length; i++) {
      if (names[i][0] === archiveName) {
        // Update timestamp in column F (Last Updated)
        const timestampCell = archivesSheet.getRange(i + 2, 6);
        timestampCell.setValue(new Date());
        console.log(`Updated timestamp for ${archiveName}`);
        return;
      }
    }
    
    console.warn(`Archive ${archiveName} not found in archives sheet`);
    
  } catch (error) {
    console.error('Error updating timestamp:', error);
  }
}

/**
 * Populate games for all archives (batch operation)
 */
function populateAllArchiveGames() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const archivesSheet = spreadsheet.getSheetByName(ARCHIVES_SHEET_NAME);
    
    if (!archivesSheet) {
      throw new Error('Archives sheet not found. Please create archives first.');
    }
    
    const lastRow = archivesSheet.getLastRow();
    if (lastRow <= 1) {
      throw new Error('No archives found. Please create archives first.');
    }
    
    // Get all archive names
    const nameRange = archivesSheet.getRange(2, 1, lastRow - 1, 1);
    const names = nameRange.getValues();
    
    let successCount = 0;
    let errorCount = 0;
    
    for (let i = 0; i < names.length; i++) {
      const archiveName = names[i][0];
      if (archiveName) {
        try {
          console.log(`Processing archive ${i + 1}/${names.length}: ${archiveName}`);
          populateArchiveGames(archiveName);
          successCount++;
          
          // Add a small delay to avoid API rate limits
          Utilities.sleep(1000);
          
        } catch (error) {
          console.error(`Failed to populate ${archiveName}:`, error);
          errorCount++;
        }
      }
    }
    
    SpreadsheetApp.getUi().alert(`Batch operation completed! Successfully populated ${successCount} archives. ${errorCount} errors.`);
    
  } catch (error) {
    console.error('Error in batch populate:', error);
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
  }
}

/**
 * Combine all archive sheets into one consolidated sheet
 */
function combineAllArchiveSheets() {
  try {
    console.log('Starting to combine all archive sheets...');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const archivesSheet = spreadsheet.getSheetByName(ARCHIVES_SHEET_NAME);
    
    if (!archivesSheet) {
      throw new Error('Archives sheet not found. Please create archives first.');
    }
    
    // Create or get the combined sheet
    const combinedSheetName = 'All Games';
    let combinedSheet = spreadsheet.getSheetByName(combinedSheetName);
    
    if (!combinedSheet) {
      combinedSheet = spreadsheet.insertSheet(combinedSheetName);
      console.log(`Created new sheet: ${combinedSheetName}`);
    } else {
      // Clear existing data
      combinedSheet.clear();
      console.log('Cleared existing data from All Games sheet');
    }
    
    // Set up headers with additional Archive column
    const headers = [[
      'Archive', 'White Username', 'White Rating', 'White Result', 'White Profile',
      'Black Username', 'Black Rating', 'Black Result', 'Black Profile',
      'White Accuracy', 'Black Accuracy', 'URL', 'FEN', 'PGN',
      'Start Time', 'End Time', 'Time Control', 'Rules', 'ECO', 'Tournament', 'Match'
    ]];
    combinedSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    
    // Get all archive names from the archives sheet
    const lastRow = archivesSheet.getLastRow();
    if (lastRow <= 1) {
      throw new Error('No archives found. Please create archives first.');
    }
    
    const nameRange = archivesSheet.getRange(2, 1, lastRow - 1, 1);
    const archiveNames = nameRange.getValues().map(row => row[0]).filter(name => name);
    
    let totalGames = 0;
    let combinedData = [];
    let processedArchives = 0;
    
    // Process each archive sheet
    for (let i = 0; i < archiveNames.length; i++) {
      const archiveName = archiveNames[i];
      
      // Parse archive name to get sheet name
      const match = archiveName.match(/^(\d{4})-(\d{2})$/);
      if (!match) {
        console.warn(`Skipping invalid archive name: ${archiveName}`);
        continue;
      }
      
      const year = match[1];
      const month = match[2];
      const shortYear = year.slice(2);
      const sheetName = `a${month}${shortYear}`;
      
      const archiveSheet = spreadsheet.getSheetByName(sheetName);
      
      if (!archiveSheet) {
        console.warn(`Archive sheet ${sheetName} not found, skipping...`);
        continue;
      }
      
      const sheetLastRow = archiveSheet.getLastRow();
      if (sheetLastRow <= 1) {
        console.log(`Archive sheet ${sheetName} is empty, skipping...`);
        continue;
      }
      
      // Get data from archive sheet (skip header row)
      const dataRange = archiveSheet.getRange(2, 1, sheetLastRow - 1, 20);
      const archiveData = dataRange.getValues();
      
      // Add archive name as first column to each row
      const archiveDataWithName = archiveData.map(row => [archiveName, ...row]);
      
      // Add to combined data
      combinedData = combinedData.concat(archiveDataWithName);
      totalGames += archiveData.length;
      processedArchives++;
      
      console.log(`Processed ${archiveData.length} games from ${sheetName}`);
    }
    
    // Write combined data to sheet
    if (combinedData.length > 0) {
      const dataRange = combinedSheet.getRange(2, 1, combinedData.length, combinedData[0].length);
      dataRange.setValues(combinedData);
      
      // Sort by Start Time (column O, index 14) in descending order (newest first)
      const sortRange = combinedSheet.getRange(2, 1, combinedData.length, combinedData[0].length);
      sortRange.sort([{column: 15, ascending: false}]); // Column 15 is Start Time (1-based indexing)
    }
    
    // Format the combined sheet
    formatCombinedSheet(combinedSheet);
    
    console.log(`Successfully combined ${totalGames} games from ${processedArchives} archives`);
    SpreadsheetApp.getUi().alert(`Success! Combined ${totalGames} games from ${processedArchives} archive sheets into "${combinedSheetName}". Games are sorted by date (newest first).`);
    
  } catch (error) {
    console.error('Error combining archive sheets:', error);
    SpreadsheetApp.getUi().alert(`Error: Failed to combine archive sheets: ${error.message}`);
  }
}

/**
 * Format the combined sheet for better readability
 */
function formatCombinedSheet(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1) return;
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, lastCol);
    
    // Format headers
    const headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    
    // Add borders to all data
    const dataRange = sheet.getRange(1, 1, lastRow, lastCol);
    dataRange.setBorder(true, true, true, true, true, true);
    
    // Alternate row colors for better readability
    for (let i = 2; i <= lastRow; i++) {
      if (i % 2 === 0) {
        sheet.getRange(i, 1, 1, lastCol).setBackground('#f8f9fa');
      }
    }
    
    // Freeze the header row
    sheet.setFrozenRows(1);
    
    // Set date columns to proper date format
    if (lastRow > 1) {
      // Start Time (column O) and End Time (column P)
      const startTimeRange = sheet.getRange(2, 15, lastRow - 1, 1);
      const endTimeRange = sheet.getRange(2, 16, lastRow - 1, 1);
      startTimeRange.setNumberFormat('yyyy-mm-dd hh:mm:ss');
      endTimeRange.setNumberFormat('yyyy-mm-dd hh:mm:ss');
    }
    
    console.log('Applied formatting to combined sheet');
    
  } catch (error) {
    console.error('Error formatting combined sheet:', error);
  }
}

/**
 * Create menu item for easy access
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Chess.com Archives')
    .addItem('Create Archive Sheets', 'createChessArchives')
    .addSeparator()
    .addItem('Populate Single Archive', 'promptPopulateSingleArchive')
    .addItem('Populate All Archives', 'populateAllArchiveGames')
    .addSeparator()
    .addItem('Combine All Archives', 'combineAllArchiveSheets')
    .addSeparator()
    .addItem('Help', 'showHelp')
    .addToUi();
}

/**
 * Show help information
 */
function showHelp() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Instructions:\n\n' +
    '1. Update the CHESS_USERNAME variable in the script with your Chess.com username\n' +
    '2. Run "Create Archive Sheets" to create archive structure\n' +
    '3. Use "Populate Single Archive" to fetch games for one specific month\n' +
    '4. Use "Populate All Archives" to fetch games for all months (may take time)\n' +
    '5. Use "Combine All Archives" to merge all archive data into one "All Games" sheet\n' +
    '6. Individual sheets will be created for each archive (named like "a0825" for August 2025)\n' +
    '7. The "Last Updated" column shows when each archive was last populated\n\n' +
    'Archive format: YYYY-MM (e.g., 2023-08)\n\n' +
    'The "All Games" sheet will contain all your games sorted by date (newest first) with an Archive column to identify the source month.\n\n' +
    'Note: Make sure your Chess.com profile is public for the API to work.');
}
