/**
 * Chess.com Archives Manager for Google Sheets - FIXED VERSION
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
 * ENHANCED: Populate games data for a specific archive month with better date handling
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
    
    // Process games data with enhanced error handling
    const gameData = [];
    
    for (let gameIndex = 0; gameIndex < games.length; gameIndex++) {
      try {
        const game = games[gameIndex];
        
        // Parse dates with enhanced error handling - handle null/undefined values
        let startDate = '';
        let endDate = '';
        
        try {
          if (game.start_time) {
            if (typeof game.start_time === 'number') {
              startDate = new Date(game.start_time * 1000);
            } else {
              startDate = new Date(game.start_time);
            }
            // Validate the date
            if (isNaN(startDate.getTime())) {
              console.warn(`Invalid start_time for game ${gameIndex + 1}: ${game.start_time}`);
              startDate = '';
            }
          }
        } catch (startDateError) {
          console.warn(`Error parsing start_time for game ${gameIndex + 1}:`, startDateError);
          startDate = '';
        }
        
        try {
          if (game.end_time) {
            if (typeof game.end_time === 'number') {
              endDate = new Date(game.end_time * 1000);
            } else {
              endDate = new Date(game.end_time);
            }
            // Validate the date
            if (isNaN(endDate.getTime())) {
              console.warn(`Invalid end_time for game ${gameIndex + 1}: ${game.end_time}`);
              endDate = '';
            }
          }
        } catch (endDateError) {
          console.warn(`Error parsing end_time for game ${gameIndex + 1}:`, endDateError);
          endDate = '';
        }
        
        // Extract white player information with safe access
        const whiteUsername = (game.white && game.white.username) ? game.white.username : '';
        const whiteRating = (game.white && game.white.rating) ? game.white.rating : '';
        const whiteResult = (game.white && game.white.result) ? game.white.result : '';
        const whiteProfile = (game.white && game.white['@id']) ? game.white['@id'] : '';
        
        // Extract black player information with safe access
        const blackUsername = (game.black && game.black.username) ? game.black.username : '';
        const blackRating = (game.black && game.black.rating) ? game.black.rating : '';
        const blackResult = (game.black && game.black.result) ? game.black.result : '';
        const blackProfile = (game.black && game.black['@id']) ? game.black['@id'] : '';
        
        // Extract accuracies with safe access
        const whiteAccuracy = (game.accuracies && game.accuracies.white) ? game.accuracies.white : '';
        const blackAccuracy = (game.accuracies && game.accuracies.black) ? game.accuracies.black : '';
        
        // Extract game details with safe defaults
        const url = game.url || '';
        const fen = game.fen || '';
        const pgn = game.pgn || '';
        const timeControl = game.time_control || '';
        const rules = game.rules || '';
        const eco = game.eco || '';
        const tournament = game.tournament || '';
        const match = game.match || '';
        
        gameData.push([
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
        ]);
        
      } catch (gameError) {
        console.error(`Error processing game ${gameIndex + 1} in ${archiveName}:`, gameError);
        // Skip this game and continue with next
        continue;
      }
    }
    
    // Write games data to sheet
    if (gameData.length > 0) {
      const dataRange = archiveSheet.getRange(2, 1, gameData.length, gameData[0].length);
      dataRange.setValues(gameData);
    }
    
    // Update the timestamp in the main archives sheet
    updateArchiveTimestamp(archiveName);
    
    console.log(`Successfully populated ${gameData.length} games for ${archiveName}`);
    SpreadsheetApp.getUi().alert(`Success! Populated ${gameData.length} games for archive ${archiveName}.`);
    
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
 * FIXED: Combine all archive sheets into one consolidated sheet
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
    let errorArchives = [];
    
    // Process each archive sheet
    for (let i = 0; i < archiveNames.length; i++) {
      const archiveName = archiveNames[i];
      
      try {
        // Parse archive name to get sheet name
        const match = archiveName.match(/^(\d{4})-(\d{2})$/);
        if (!match) {
          console.warn(`Skipping invalid archive name: ${archiveName}`);
          errorArchives.push(archiveName + ' (invalid format)');
          continue;
        }
        
        const year = match[1];
        const month = match[2];
        const shortYear = year.slice(2);
        const sheetName = `a${month}${shortYear}`;
        
        const archiveSheet = spreadsheet.getSheetByName(sheetName);
        
        if (!archiveSheet) {
          console.warn(`Archive sheet ${sheetName} not found, skipping...`);
          errorArchives.push(archiveName + ' (sheet not found)');
          continue;
        }
        
        const sheetLastRow = archiveSheet.getLastRow();
        if (sheetLastRow <= 1) {
          console.log(`Archive sheet ${sheetName} is empty, skipping...`);
          continue;
        }
        
        // FIXED: Get the actual number of columns from the archive sheet
        const sheetLastCol = archiveSheet.getLastColumn();
        const expectedCols = 20; // We expect 20 columns based on our headers
        
        // Use the minimum of expected columns or actual columns to avoid range errors
        const colsToRead = Math.min(expectedCols, sheetLastCol);
        
        // Get data from archive sheet (skip header row)
        const dataRange = archiveSheet.getRange(2, 1, sheetLastRow - 1, colsToRead);
        const archiveData = dataRange.getValues();
        
        // FIXED: Process and normalize each row with better error handling
        const normalizedArchiveData = [];
        
        for (let rowIndex = 0; rowIndex < archiveData.length; rowIndex++) {
          try {
            const row = archiveData[rowIndex];
            
            // Create a new row with proper data types
            const normalizedRow = [];
            
            for (let colIndex = 0; colIndex < expectedCols; colIndex++) {
              let cellValue = '';
              
              if (colIndex < row.length && row[colIndex] != null) {
                cellValue = row[colIndex];
                
                // Special handling for date columns (13 = Start Time, 14 = End Time)
                if (colIndex === 13 || colIndex === 14) {
                  try {
                    if (cellValue instanceof Date) {
                      // Already a date object, use as-is
                      normalizedRow.push(cellValue);
                      continue;
                    } else if (typeof cellValue === 'number' && cellValue > 0) {
                      // Unix timestamp, convert to date
                      normalizedRow.push(new Date(cellValue * 1000));
                      continue;
                    } else if (typeof cellValue === 'string' && cellValue.trim()) {
                      // Try to parse as date string
                      const parsedDate = new Date(cellValue);
                      if (!isNaN(parsedDate.getTime())) {
                        normalizedRow.push(parsedDate);
                        continue;
                      }
                    }
                    // If all else fails, use empty string for dates
                    normalizedRow.push('');
                  } catch (dateError) {
                    console.warn(`Date parsing error in ${archiveName}, row ${rowIndex + 2}, col ${colIndex + 1}:`, dateError);
                    normalizedRow.push('');
                  }
                } else {
                  // For non-date columns, convert to string and handle special cases
                  if (typeof cellValue === 'object' && cellValue !== null) {
                    // Handle objects (shouldn't happen but just in case)
                    normalizedRow.push(cellValue.toString());
                  } else {
                    normalizedRow.push(cellValue);
                  }
                }
              } else {
                // Missing data, use empty string
                normalizedRow.push('');
              }
            }
            
            normalizedArchiveData.push(normalizedRow);
            
          } catch (rowError) {
            console.error(`Error processing row ${rowIndex + 2} in ${archiveName}:`, rowError);
            // Skip this row and continue with next
            continue;
          }
        }
        
        // Add archive name as first column to each row
        const archiveDataWithName = normalizedArchiveData.map((row, index) => {
          try {
            return [archiveName, ...row];
          } catch (error) {
            console.error(`Error adding archive name to row ${index + 2} in ${archiveName}:`, error);
            return null; // Mark as invalid
          }
        }).filter(row => row !== null); // Remove invalid rows
        
        // FIXED: Validate data before adding to combined data
        const validRows = archiveDataWithName.filter((row, index) => {
          try {
            // Check if the row has the expected number of columns (21 = 1 for archive + 20 for game data)
            if (row.length !== 21) {
              console.warn(`Row ${index + 2} in ${archiveName} has ${row.length} columns, expected 21`);
              return false;
            }
            return true;
          } catch (error) {
            console.error(`Error validating row ${index + 2} in ${archiveName}:`, error);
            return false;
          }
        });
        
        if (validRows.length !== archiveDataWithName.length) {
          console.warn(`Filtered out ${archiveDataWithName.length - validRows.length} invalid rows from ${sheetName}`);
        }
        
        // Add to combined data
        combinedData = combinedData.concat(validRows);
        totalGames += validRows.length;
        processedArchives++;
        
        console.log(`Processed ${validRows.length} games from ${sheetName}`);
        
      } catch (archiveError) {
        console.error(`Error processing archive ${archiveName}:`, archiveError);
        console.error(`Archive error details:`, {
          archiveName: archiveName,
          sheetName: sheetName,
          error: archiveError.message,
          stack: archiveError.stack
        });
        errorArchives.push(archiveName + ` (${archiveError.message})`);
        continue;
      }
    }
    
    // Write combined data to sheet in batches to avoid timeouts
    if (combinedData.length > 0) {
      const BATCH_SIZE = 1000;
      let rowsWritten = 0;
      
      for (let i = 0; i < combinedData.length; i += BATCH_SIZE) {
        const batch = combinedData.slice(i, Math.min(i + BATCH_SIZE, combinedData.length));
        const startRow = 2 + rowsWritten; // Start after header + already written rows
        
        try {
          const dataRange = combinedSheet.getRange(startRow, 1, batch.length, batch[0].length);
          dataRange.setValues(batch);
          rowsWritten += batch.length;
          console.log(`Wrote batch ${Math.floor(i / BATCH_SIZE) + 1}, rows ${startRow} to ${startRow + batch.length - 1}`);
        } catch (error) {
          console.error(`Error writing batch starting at row ${startRow}:`, error);
          throw new Error(`Failed to write data batch: ${error.message}`);
        }
      }
      
      // FIXED: Sort by Start Time (column 15, which is index 14) in descending order (newest first)
      try {
        if (combinedData.length > 1) {
          const sortRange = combinedSheet.getRange(2, 1, combinedData.length, combinedData[0].length);
          sortRange.sort([{column: 15, ascending: false}]); // Column 15 is Start Time
          console.log('Successfully sorted data by start time');
        }
      } catch (error) {
        console.warn('Could not sort data by start time:', error);
        // Continue without sorting rather than failing
      }
    }
    
    // Format the combined sheet
    try {
      formatCombinedSheet(combinedSheet);
    } catch (error) {
      console.warn('Error formatting sheet:', error);
      // Continue without formatting rather than failing
    }
    
    // Prepare success/error message
    let message = `Success! Combined ${totalGames} games from ${processedArchives} archive sheets into "${combinedSheetName}".`;
    if (combinedData.length > 0) {
      message += ` Games are sorted by date (newest first).`;
    }
    
    if (errorArchives.length > 0) {
      message += `\n\nWarning: ${errorArchives.length} archives had issues:\n${errorArchives.join('\n')}`;
    }
    
    console.log(`Successfully combined ${totalGames} games from ${processedArchives} archives`);
    SpreadsheetApp.getUi().alert(message);
    
  } catch (error) {
    console.error('Error combining archive sheets:', error);
    SpreadsheetApp.getUi().alert(`Error: Failed to combine archive sheets: ${error.message}`);
  }
}

/**
 * FIXED: Format the combined sheet for better readability
 */
function formatCombinedSheet(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1 || lastCol === 0) {
      console.log('No data to format in combined sheet');
      return;
    }
    
    // Auto-resize columns with reasonable limits
    for (let col = 1; col <= Math.min(lastCol, 21); col++) {
      try {
        sheet.autoResizeColumn(col);
        // Set maximum width to prevent extremely wide columns
        const currentWidth = sheet.getColumnWidth(col);
        if (currentWidth > 300) {
          sheet.setColumnWidth(col, 300);
        }
      } catch (error) {
        console.warn(`Could not resize column ${col}:`, error);
      }
    }
    
    // Format headers
    try {
      const headerRange = sheet.getRange(1, 1, 1, lastCol);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('#ffffff');
    } catch (error) {
      console.warn('Could not format headers:', error);
    }
    
    // Add borders to all data (in smaller batches to avoid timeouts)
    try {
      const BORDER_BATCH_SIZE = 500;
      for (let startRow = 1; startRow <= lastRow; startRow += BORDER_BATCH_SIZE) {
        const endRow = Math.min(startRow + BORDER_BATCH_SIZE - 1, lastRow);
        const batchRange = sheet.getRange(startRow, 1, endRow - startRow + 1, lastCol);
        batchRange.setBorder(true, true, true, true, true, true);
      }
    } catch (error) {
      console.warn('Could not add borders:', error);
    }
    
    // Alternate row colors for better readability (in batches)
    try {
      const COLOR_BATCH_SIZE = 200;
      for (let startRow = 2; startRow <= lastRow; startRow += COLOR_BATCH_SIZE) {
        const endRow = Math.min(startRow + COLOR_BATCH_SIZE - 1, lastRow);
        
        for (let i = startRow; i <= endRow; i++) {
          if (i % 2 === 0) {
            const rowRange = sheet.getRange(i, 1, 1, lastCol);
            rowRange.setBackground('#f8f9fa');
          }
        }
      }
    } catch (error) {
      console.warn('Could not apply alternating row colors:', error);
    }
    
    // Freeze the header row
    try {
      sheet.setFrozenRows(1);
    } catch (error) {
      console.warn('Could not freeze header row:', error);
    }
    
    // Set date columns to proper date format
    try {
      if (lastRow > 1 && lastCol >= 16) {
        // Start Time (column 15) and End Time (column 16)
        const startTimeRange = sheet.getRange(2, 15, lastRow - 1, 1);
        const endTimeRange = sheet.getRange(2, 16, lastRow - 1, 1);
        startTimeRange.setNumberFormat('yyyy-mm-dd hh:mm:ss');
        endTimeRange.setNumberFormat('yyyy-mm-dd hh:mm:ss');
      }
    } catch (error) {
      console.warn('Could not format date columns:', error);
    }
    
    console.log('Applied formatting to combined sheet');
    
  } catch (error) {
    console.error('Error formatting combined sheet:', error);
    // Don't throw - formatting errors shouldn't stop the main process
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
  ui.alert('Chess.com Archives Manager - Instructions:\n\n' +
    '1. UPDATE THE USERNAME: Change CHESS_USERNAME variable in the script to your Chess.com username\n' +
    '2. CREATE ARCHIVES: Run "Create Archive Sheets" to create the archive structure\n' +
    '3. POPULATE DATA: Use "Populate Single Archive" for one month or "Populate All Archives" for all months\n' +
    '4. COMBINE DATA: Use "Combine All Archives" to merge all data into one "All Games" sheet\n\n' +
    'FEATURES:\n' +
    '• Individual sheets created for each month (e.g., "a0825" for August 2025)\n' +
    '• "Last Updated" column shows when each archive was populated\n' +
    '• "All Games" sheet combines all data, sorted by date (newest first)\n' +
    '• Archive column identifies the source month for each game\n\n' +
    'ARCHIVE FORMAT: YYYY-MM (e.g., 2023-08)\n\n' +
    'TROUBLESHOOTING:\n' +
    '• Ensure your Chess.com profile is public\n' +
    '• Check username spelling in the script\n' +
    '• Large datasets may take time to process\n' +
    '• If errors occur, try processing individual archives\n\n' +
    'FIXED ISSUES IN THIS VERSION:\n' +
    '• Better error handling for missing data\n' +
    '• Improved data validation and normalization\n' +
    '• Batch processing to avoid timeouts\n' +
    '• More robust column handling\n' +
    '• Enhanced formatting with error recovery');
}

/**
 * UTILITY FUNCTIONS FOR DEBUGGING AND MAINTENANCE
 */

/**
 * Debug function to check archive sheet structure
 */
function debugArchiveSheet(archiveName) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const match = archiveName.match(/^(\d{4})-(\d{2})$/);
    
    if (!match) {
      console.log(`Invalid archive name: ${archiveName}`);
      return;
    }
    
    const year = match[1];
    const month = match[2];
    const shortYear = year.slice(2);
    const sheetName = `a${month}${shortYear}`;
    
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      console.log(`Sheet ${sheetName} not found`);
      return;
    }
    
    console.log(`=== DEBUG INFO FOR ${sheetName} ===`);
    console.log(`Last Row: ${sheet.getLastRow()}`);
    console.log(`Last Column: ${sheet.getLastColumn()}`);
    
    if (sheet.getLastRow() > 0) {
      const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
      const headers = headerRange.getValues()[0];
      console.log(`Headers (${headers.length}):`, headers);
      
      if (sheet.getLastRow() > 1) {
        const sampleRange = sheet.getRange(2, 1, 1, sheet.getLastColumn());
        const sampleData = sampleRange.getValues()[0];
        console.log(`Sample row (${sampleData.length}):`, sampleData);
      }
    }
    
  } catch (error) {
    console.error(`Debug error for ${archiveName}:`, error);
  }
}

/**
 * Clean up function to remove empty or problematic archive sheets
 */
function cleanupEmptyArchiveSheets() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    let deletedCount = 0;
    let skippedSheets = [];
    
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const sheetName = sheet.getName();
      
      // Only process archive sheets (pattern: a[0-9]{4})
      if (sheetName.match(/^a\d{4}$/)) {
        const lastRow = sheet.getLastRow();
        
        // Delete sheets with no data or only headers
        if (lastRow <= 1) {
          try {
            spreadsheet.deleteSheet(sheet);
            deletedCount++;
            console.log(`Deleted empty sheet: ${sheetName}`);
          } catch (deleteError) {
            console.warn(`Could not delete sheet ${sheetName}:`, deleteError);
            skippedSheets.push(sheetName);
          }
        }
      }
    }
    
    let message = `Cleanup completed! Deleted ${deletedCount} empty archive sheets.`;
    if (skippedSheets.length > 0) {
      message += `\nCould not delete: ${skippedSheets.join(', ')}`;
    }
    
    SpreadsheetApp.getUi().alert(message);
    
  } catch (error) {
    console.error('Error during cleanup:', error);
    SpreadsheetApp.getUi().alert(`Cleanup error: ${error.message}`);
  }
}

/**
 * Validate all archive sheets and report issues
 */
function validateAllArchiveSheets() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const archivesSheet = spreadsheet.getSheetByName(ARCHIVES_SHEET_NAME);
    
    if (!archivesSheet) {
      throw new Error('Archives sheet not found.');
    }
    
    const lastRow = archivesSheet.getLastRow();
    if (lastRow <= 1) {
      throw new Error('No archives found.');
    }
    
    const nameRange = archivesSheet.getRange(2, 1, lastRow - 1, 1);
    const archiveNames = nameRange.getValues().map(row => row[0]).filter(name => name);
    
    let validArchives = 0;
    let emptyArchives = 0;
    let missingArchives = 0;
    let issues = [];
    
    for (let i = 0; i < archiveNames.length; i++) {
      const archiveName = archiveNames[i];
      
      const match = archiveName.match(/^(\d{4})-(\d{2})$/);
      if (!match) {
        issues.push(`${archiveName}: Invalid format`);
        continue;
      }
      
      const year = match[1];
      const month = match[2];
      const shortYear = year.slice(2);
      const sheetName = `a${month}${shortYear}`;
      
      const sheet = spreadsheet.getSheetByName(sheetName);
      
      if (!sheet) {
        missingArchives++;
        issues.push(`${archiveName}: Sheet ${sheetName} not found`);
        continue;
      }
      
      const sheetLastRow = sheet.getLastRow();
      if (sheetLastRow <= 1) {
        emptyArchives++;
        issues.push(`${archiveName}: No game data`);
      } else {
        validArchives++;
      }
    }
    
    let report = `ARCHIVE VALIDATION REPORT\n\n`;
    report += `Total Archives: ${archiveNames.length}\n`;
    report += `Valid Archives with Data: ${validArchives}\n`;
    report += `Empty Archives: ${emptyArchives}\n`;
    report += `Missing Archive Sheets: ${missingArchives}\n\n`;
    
    if (issues.length > 0) {
      report += `ISSUES FOUND:\n${issues.join('\n')}`;
    } else {
      report += `No issues found! All archives are valid.`;
    }
    
    console.log(report);
    SpreadsheetApp.getUi().alert(report);
    
  } catch (error) {
    console.error('Validation error:', error);
    SpreadsheetApp.getUi().alert(`Validation error: ${error.message}`);
  }
}
