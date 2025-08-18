// Chess.com Game Data Manager
// This script manages Chess.com game data across multiple sheets

// Configuration
const SHEETS = {
  USERNAME: 'Username Input',
  ARCHIVES: 'Archives',
  GAMES: 'Game Data',
  LOGS: 'Execution Logs',
  STATS: 'Player Stats'
};

const HEADERS = {
  ARCHIVES: ['Archive URL', 'Year-Month', 'Status', 'Last Updated'],
  GAMES: [
    'Game URL',
    'Time Control',
    'Rated',
    'Time Class',
    'Rules',
    'Format',
    'End Time',
    'White Username',
    'White Rating',
    'White Result',
    'Black Username',
    'Black Rating',
    'Black Result',
    'White Accuracy',
    'Black Accuracy',
    'Event',
    'Site',
    'Date',
    'Round',
    'Opening',
    'ECO',
    'Termination',
    'UTC Date',
    'UTC Time',
    'Start Time',
    'End Date',
    'End Time',
    'Current Position',
    'Full PGN',
    'Moves & Times'
  ],
  LOGS: ['Timestamp', 'Function', 'Username', 'Status', 'Archives Processed', 'New Games Added', 'Total Games', 'Execution Time (ms)', 'Errors', 'Notes'],
  STATS: ['Path', 'Value']
};

/**
 * Initial setup - creates all necessary sheets and headers
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create Username Input sheet
  let usernameSheet = ss.getSheetByName(SHEETS.USERNAME);
  let existingUsername = '';
  
  if (!usernameSheet) {
    usernameSheet = ss.insertSheet(SHEETS.USERNAME);
  } else {
    // Preserve existing username if it exists
    const currentValue = usernameSheet.getRange('B1').getValue().toString().trim();
    if (currentValue && currentValue !== 'Enter your username here') {
      existingUsername = currentValue;
    }
  }
  
  usernameSheet.getRange('A1').setValue('Chess.com Username:');
  // Only set placeholder if no existing username
  if (!existingUsername) {
    usernameSheet.getRange('B1').setValue('Enter your username here');
  } else {
    usernameSheet.getRange('B1').setValue(existingUsername);
  }
  usernameSheet.getRange('A3').setValue('Instructions:');
  usernameSheet.getRange('A4').setValue('1. Enter your username in cell B1');
  usernameSheet.getRange('A5').setValue('2. Run fetchAllData() for initial load');
  usernameSheet.getRange('A6').setValue('3. Run fetchRecentData() for updates');
  
  // Create Archives sheet
  let archivesSheet = ss.getSheetByName(SHEETS.ARCHIVES);
  if (!archivesSheet) {
    archivesSheet = ss.insertSheet(SHEETS.ARCHIVES);
  }
  if (archivesSheet.getLastRow() === 0) {
    archivesSheet.getRange(1, 1, 1, HEADERS.ARCHIVES.length).setValues([HEADERS.ARCHIVES]);
    archivesSheet.getRange(1, 1, 1, HEADERS.ARCHIVES.length).setFontWeight('bold');
  }
  
  // Create Games sheet
  let gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  if (!gamesSheet) {
    gamesSheet = ss.insertSheet(SHEETS.GAMES);
  }
  // Always ensure headers reflect the latest schema
  gamesSheet.getRange(1, 1, 1, HEADERS.GAMES.length).setValues([HEADERS.GAMES]);
  gamesSheet.getRange(1, 1, 1, HEADERS.GAMES.length).setFontWeight('bold');
  gamesSheet.getRange(1, 1, 1, HEADERS.GAMES.length).setBackground('#4285f4').setFontColor('white');
  
  // Create Logs sheet
  let logsSheet = ss.getSheetByName(SHEETS.LOGS);
  if (!logsSheet) {
    logsSheet = ss.insertSheet(SHEETS.LOGS);
  }
  if (logsSheet.getLastRow() === 0) {
    logsSheet.getRange(1, 1, 1, HEADERS.LOGS.length).setValues([HEADERS.LOGS]);
    logsSheet.getRange(1, 1, 1, HEADERS.LOGS.length).setFontWeight('bold');
    logsSheet.getRange(1, 1, 1, HEADERS.LOGS.length).setBackground('#34a853').setFontColor('white');
  }

  // Create Player Stats sheet
  let statsSheet = ss.getSheetByName(SHEETS.STATS);
  if (!statsSheet) {
    statsSheet = ss.insertSheet(SHEETS.STATS);
  }
  if (statsSheet.getLastRow() === 0) {
    statsSheet.getRange(1, 1, 1, HEADERS.STATS.length).setValues([HEADERS.STATS]);
    statsSheet.getRange(1, 1, 1, HEADERS.STATS.length).setFontWeight('bold');
    statsSheet.getRange(1, 1, 1, HEADERS.STATS.length).setBackground('#fbbc04').setFontColor('black');
  }
}

/**
 * Gets username from the input sheet
 */
function getUsername() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usernameSheet = ss.getSheetByName(SHEETS.USERNAME);
  if (!usernameSheet) {
    throw new Error('Username sheet not found. Please run setupSheets() first.');
  }
  
  const username = usernameSheet.getRange('B1').getValue().toString().trim();
  if (!username || username === 'Enter your username here') {
    throw new Error('Please enter your Chess.com username in the Username Input sheet.');
  }
  
  return username;
}

/**
 * Fetches all archives for a user
 */
function fetchArchives(username) {
  const url = `https://api.chess.com/pub/player/${username}/games/archives`;
  
  try {
    const response = UrlFetchApp.fetch(url);
    if (response.getResponseCode() !== 200) {
      throw new Error(`Failed to fetch archives. Response code: ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    return data.archives || [];
  } catch (error) {
    console.error('Error fetching archives:', error);
    throw new Error(`Failed to fetch archives: ${error.message}`);
  }
}

/**
 * Fetch and write player stats to a dedicated sheet
 */
function fetchPlayerStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const username = getUsername();
  const statsUrl = `https://api.chess.com/pub/player/${username}/stats`;
  let data;
  try {
    const response = UrlFetchApp.fetch(statsUrl);
    if (response.getResponseCode() !== 200) {
      throw new Error(`Failed to fetch stats. Response code: ${response.getResponseCode()}`);
    }
    data = JSON.parse(response.getContentText());
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Stats fetch failed: ${error.message}`, 'Stats Error', 5);
    throw error;
  }

  // Ensure sheet and headers
  let statsSheet = ss.getSheetByName(SHEETS.STATS);
  if (!statsSheet) {
    statsSheet = ss.insertSheet(SHEETS.STATS);
  }
  statsSheet.clear();
  statsSheet.getRange(1, 1, 1, HEADERS.STATS.length).setValues([HEADERS.STATS]);
  statsSheet.getRange(1, 1, 1, HEADERS.STATS.length).setFontWeight('bold');
  statsSheet.getRange(1, 1, 1, HEADERS.STATS.length).setBackground('#fbbc04').setFontColor('black');

  // Flatten JSON to path-value rows
  const rows = [];
  function walk(obj, pathParts) {
    if (obj === null || obj === undefined) return;
    const isPrimitive = ['string', 'number', 'boolean'].includes(typeof obj);
    if (isPrimitive) {
      rows.push([pathParts.join('.'), obj]);
      return;
    }
    if (Array.isArray(obj)) {
      obj.forEach((item, index) => walk(item, pathParts.concat([String(index)])));
      return;
    }
    Object.keys(obj).forEach(key => walk(obj[key], pathParts.concat([key])));
  }
  walk(data, []);

  if (rows.length === 0) {
    rows.push(['info', 'No data']);
  }

  statsSheet.getRange(2, 1, rows.length, HEADERS.STATS.length).setValues(rows);
  SpreadsheetApp.getActiveSpreadsheet().toast(`Fetched ${rows.length} stats rows for ${username}`, 'Stats Updated', 5);
}

/**
 * Updates the archives sheet with archive URLs and their status
 */
function updateArchivesSheet(archives) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archivesSheet = ss.getSheetByName(SHEETS.ARCHIVES);
  
  // Get existing archives
  const existingData = archivesSheet.getLastRow() > 1 ? 
    archivesSheet.getRange(2, 1, archivesSheet.getLastRow() - 1, 4).getValues() : [];
  
  const existingArchives = new Set(existingData.map(row => row[0]));
  const newArchives = [];
  
  // Process each archive
  archives.forEach((archiveUrl, index) => {
    const urlParts = archiveUrl.split('/');
    const yearMonth = urlParts[urlParts.length - 2] + '-' + urlParts[urlParts.length - 1];
    const isLatest = index === archives.length - 1;
    const status = isLatest ? 'Active' : 'Inactive';
    const lastUpdated = new Date().toISOString().split('T')[0];
    
    if (!existingArchives.has(archiveUrl)) {
      newArchives.push([archiveUrl, yearMonth, status, lastUpdated]);
    } else {
      // Update status of existing archives
      const existingRowIndex = existingData.findIndex(row => row[0] === archiveUrl);
      if (existingRowIndex !== -1) {
        archivesSheet.getRange(existingRowIndex + 2, 3).setValue(status);
      }
    }
  });
  
  // Add new archives
  if (newArchives.length > 0) {
    archivesSheet.getRange(archivesSheet.getLastRow() + 1, 1, newArchives.length, 4)
      .setValues(newArchives);
  }
  
  return archives[archives.length - 1]; // Return latest archive
}

/**
 * Fetches games from a specific archive
 */
function fetchGamesFromArchive(archiveUrl) {
  try {
    const response = UrlFetchApp.fetch(archiveUrl);
    if (response.getResponseCode() !== 200) {
      throw new Error(`Failed to fetch games from ${archiveUrl}. Response code: ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    return data.games || [];
  } catch (error) {
    console.error(`Error fetching games from ${archiveUrl}:`, error);
    return [];
  }
}

/**
 * Logs execution details to the logs sheet
 */
function logExecution(functionName, username, status, archivesProcessed, newGamesAdded, totalGames, executionTime, errors = '', notes = '') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logsSheet = ss.getSheetByName(SHEETS.LOGS);
    
    if (!logsSheet) return;
    
    const logRow = [
      new Date(),
      functionName,
      username,
      status,
      archivesProcessed,
      newGamesAdded,
      totalGames,
      executionTime,
      errors,
      notes
    ];
    
    // Insert at row 2 to keep most recent logs at top
    if (logsSheet.getLastRow() > 1) {
      logsSheet.insertRows(2);
    }
    
    logsSheet.getRange(logsSheet.getLastRow() > 1 ? 2 : logsSheet.getLastRow() + 1, 1, 1, HEADERS.LOGS.length)
      .setValues([logRow]);
      
    // Color code status
    const statusCell = logsSheet.getRange(logsSheet.getLastRow() > 1 ? 2 : logsSheet.getLastRow(), 4);
    if (status === 'SUCCESS') {
      statusCell.setBackground('#d9ead3');
    } else if (status === 'ERROR') {
      statusCell.setBackground('#f4c7c3');
    } else if (status === 'WARNING') {
      statusCell.setBackground('#fce5cd');
    }
    
  } catch (error) {
    console.error('Error logging execution:', error);
  }
}

/**
 * Parses PGN string to extract metadata and moves
 */
function parsePGN(pgnString) {
  if (!pgnString) return {};
  
  const result = {
    event: '',
    site: '',
    date: '',
    round: '',
    opening: '',
    eco: '',
    termination: '',
    moves: '',
    utcDate: '',
    utcTime: '',
    startTime: '',
    endDate: '',
    endTime: '',
    currentPosition: ''
  };
  
  try {
    const lines = pgnString.split('\n');
    let movesStarted = false;
    let movesText = '';
    
    for (const line of lines) {
      const trimmedLine = line.trim();
      
      if (trimmedLine.startsWith('[') && trimmedLine.endsWith(']')) {
        // Parse metadata
        const match = trimmedLine.match(/\[(\w+)\s+"([^"]+)"\]/);
        if (match) {
          const key = match[1].toLowerCase();
          const value = match[2];
          
          if (key === 'event') result.event = value;
          else if (key === 'site') result.site = value;
          else if (key === 'date') result.date = value;
          else if (key === 'round') result.round = value;
          else if (key === 'opening') result.opening = value;
          else if (key === 'eco') result.eco = value;
          else if (key === 'termination') result.termination = value;
          else if (key === 'utcdate') result.utcDate = value;
          else if (key === 'utctime') result.utcTime = value;
          else if (key === 'starttime') result.startTime = value;
          else if (key === 'enddate') result.endDate = value;
          else if (key === 'endtime') result.endTime = value;
          else if (key === 'currentposition') result.currentPosition = value;
        }
      } else if (trimmedLine && !trimmedLine.startsWith('[')) {
        // This is moves section
        movesStarted = true;
        movesText += (movesText ? ' ' : '') + trimmedLine;
      }
    }
    
    result.moves = movesText.trim();
    
  } catch (error) {
    console.error('Error parsing PGN:', error);
  }
  
  return result;
}
/**
 * Computes a normalized game format label based on rules and time class
 */
function computeFormat(rules, timeClass) {
  const normalizedRules = (rules || '').toLowerCase();
  const normalizedTimeClass = (timeClass || '').toLowerCase();
  const isChess960 = normalizedRules.includes('960');
  const isStandardChess = normalizedRules === 'chess' || normalizedRules === '';
  if (isChess960 && normalizedTimeClass === 'daily') {
    return 'daily960';
  }
  if (isStandardChess) {
    return normalizedTimeClass || 'unknown';
  }
  return normalizedRules;
}
function gameToRow(game) {
  const pgn = game.pgn || '';
  const metadata = parsePGN(pgn);
  const rules = (game.rules || '').toLowerCase();
  const timeClass = (game.time_class || '').toLowerCase();
  const format = computeFormat(rules, timeClass);
  return [
    // 1-6
    game.url || '',
    game.time_control || '',
    game.rated || false,
    game.time_class || '',
    game.rules || '',
    format,
    // 7
    game.end_time ? new Date(game.end_time * 1000) : '',
    // 8-13
    game.white?.username || '',
    game.white?.rating || '',
    game.white?.result || '',
    game.black?.username || '',
    game.black?.rating || '',
    game.black?.result || '',
    // 14-15
    game.accuracies?.white ?? '',
    game.accuracies?.black ?? '',
    // 16-22 PGN-derived metadata
    metadata.event || '',
    metadata.site || '',
    metadata.date || '',
    metadata.round || '',
    metadata.opening || '',
    metadata.eco || '',
    metadata.termination || '',
    // 23-28 PGN extra tags
    metadata.utcDate || '',
    metadata.utcTime || '',
    metadata.startTime || '',
    metadata.endDate || '',
    metadata.endTime || '',
    metadata.currentPosition || '',
    // 29-30
    pgn,
    metadata.moves || ''
  ];
}

/**
 * Gets existing game URLs to prevent duplicates
 */
function getExistingGameUrls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  
  if (gamesSheet.getLastRow() <= 1) return new Set();
  
  const gameUrls = gamesSheet.getRange(2, 1, gamesSheet.getLastRow() - 1, 1).getValues();
  return new Set(gameUrls.flat().filter(url => url));
}

/**
 * Adds new games to the top of the games sheet (avoiding duplicates)
 */
function addNewGames(games) {
  if (!games || games.length === 0) return 0;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const existingUrls = getExistingGameUrls();
  
  // Filter out duplicates and convert to rows
  const newGameRows = games
    .filter(game => !existingUrls.has(game.url))
    .map(gameToRow)
    .reverse(); // Reverse to maintain chronological order when inserting at top
  
  if (newGameRows.length === 0) {
    console.log('No new games to add (all games already exist)');
    return 0;
  }
  
  // Insert new rows at the top (after headers)
  if (gamesSheet.getLastRow() > 1) {
    gamesSheet.insertRows(2, newGameRows.length);
  }
  
  gamesSheet.getRange(2, 1, newGameRows.length, HEADERS.GAMES.length)
    .setValues(newGameRows);
  
  console.log(`Added ${newGameRows.length} new games`);
  return newGameRows.length;
}

/**
 * Main function to fetch all data (initial load)
 */
function fetchAllData() {
  const startTime = Date.now();
  let username = '';
  let archivesProcessed = 0;
  let totalGames = 0;
  let errors = '';
  
  try {
    console.log('Starting initial data fetch...');
    
    // Only setup sheets if they don't exist, preserve username
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName(SHEETS.USERNAME) || !ss.getSheetByName(SHEETS.ARCHIVES) || !ss.getSheetByName(SHEETS.GAMES)) {
      setupSheets();
    }
    
    username = getUsername();
    console.log(`Fetching data for user: ${username}`);
    
    const archives = fetchArchives(username);
    console.log(`Found ${archives.length} archives`);
    
    if (archives.length === 0) {
      throw new Error('No archives found for this user');
    }
    
    updateArchivesSheet(archives);
    archivesProcessed = archives.length;
    
    // Fetch games from all archives
    for (let i = 0; i < archives.length; i++) {
      const archiveUrl = archives[i];
      console.log(`Fetching games from archive ${i + 1}/${archives.length}...`);
      
      const games = fetchGamesFromArchive(archiveUrl);
      const newGamesCount = addNewGames(games);
      totalGames += newGamesCount;
      
      // Add a small delay to be respectful to the API
      Utilities.sleep(100);
    }
    
    const executionTime = Date.now() - startTime;
    
    console.log(`Initial fetch complete! Added ${totalGames} total games.`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Successfully loaded ${totalGames} games from ${archives.length} archives`, 
      'Data Fetch Complete', 
      5
    );
    
    // Log successful execution
    logExecution(
      'fetchAllData', 
      username, 
      'SUCCESS', 
      archivesProcessed, 
      totalGames, 
      totalGames, 
      executionTime,
      '',
      `Initial load completed successfully`
    );
    
  } catch (error) {
    const executionTime = Date.now() - startTime;
    errors = error.message;
    
    console.error('Error in fetchAllData:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Error: ${error.message}`, 
      'Fetch Failed', 
      10
    );
    
    // Log failed execution
    logExecution(
      'fetchAllData', 
      username, 
      'ERROR', 
      archivesProcessed, 
      totalGames, 
      totalGames, 
      executionTime,
      errors,
      'Initial load failed'
    );
  }
}

/**
 * Fetches only recent data (from last few archives)
 */
function fetchRecentData(archiveCount = 3) {
  const startTime = Date.now();
  let username = '';
  let archivesProcessed = 0;
  let totalNewGames = 0;
  let errors = '';
  
  try {
    console.log('Starting recent data fetch...');
    
    username = getUsername();
    const archives = fetchArchives(username);
    
    if (archives.length === 0) {
      throw new Error('No archives found for this user');
    }
    
    updateArchivesSheet(archives);
    
    // Get the last few archives
    const recentArchives = archives.slice(-Math.min(archiveCount, archives.length));
    archivesProcessed = recentArchives.length;
    console.log(`Fetching from ${recentArchives.length} most recent archives`);
    
    // Get current total games for logging
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
    const currentTotalGames = Math.max(0, gamesSheet.getLastRow() - 1);
    
    for (const archiveUrl of recentArchives) {
      console.log(`Fetching recent games from: ${archiveUrl}`);
      
      const games = fetchGamesFromArchive(archiveUrl);
      const newGamesCount = addNewGames(games);
      totalNewGames += newGamesCount;
      
      Utilities.sleep(100);
    }
    
    const executionTime = Date.now() - startTime;
    const finalTotalGames = currentTotalGames + totalNewGames;
    
    console.log(`Recent fetch complete! Added ${totalNewGames} new games.`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Added ${totalNewGames} new games from recent archives`, 
      'Recent Update Complete', 
      5
    );
    
    // Log successful execution
    logExecution(
      'fetchRecentData', 
      username, 
      'SUCCESS', 
      archivesProcessed, 
      totalNewGames, 
      finalTotalGames, 
      executionTime,
      '',
      `Checked ${archiveCount} recent archives`
    );
    
  } catch (error) {
    const executionTime = Date.now() - startTime;
    errors = error.message;
    
    console.error('Error in fetchRecentData:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Error: ${error.message}`, 
      'Recent Fetch Failed', 
      10
    );
    
    // Log failed execution
    logExecution(
      'fetchRecentData', 
      username, 
      'ERROR', 
      archivesProcessed, 
      totalNewGames, 
      0, 
      executionTime,
      errors,
      `Recent fetch failed after ${archivesProcessed} archives`
    );
  }
}

/**
 * Fetches from the last 1 archive only (fastest update)
 */
function fetchLatestArchive() {
  const startTime = Date.now();
  let username = '';
  
  try {
    username = getUsername();
    fetchRecentData(1);
    
    // The logging is handled in fetchRecentData, but we can add a note
    const executionTime = Date.now() - startTime;
    console.log(`Latest archive fetch completed in ${executionTime}ms`);
    
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    logExecution(
      'fetchLatestArchive', 
      username, 
      'ERROR', 
      0, 
      0, 
      0, 
      executionTime,
      error.message,
      'Latest archive fetch failed'
    );
  }
}

/**
 * Menu creation function - adds custom menu to spreadsheet
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Chess.com Data')
    .addItem('Setup Sheets', 'setupSheets')
    .addSeparator()
    .addItem('Fetch All Data (Initial)', 'fetchAllData')
    .addItem('Fetch Recent Data (3 archives)', 'fetchRecentData')
    .addItem('Fetch Latest Archive Only', 'fetchLatestArchive')
    .addItem('Fetch Player Stats', 'fetchPlayerStats')
    .addSeparator()
    .addItem('View Execution Logs', 'openLogsSheet')
    .addToUi();
}

/**
 * Helper function to open logs sheet
 */
function openLogsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logsSheet = ss.getSheetByName(SHEETS.LOGS);
  if (logsSheet) {
    ss.setActiveSheet(logsSheet);
  } else {
    ss.toast('Logs sheet not found. Please run Setup Sheets first.', 'Sheet Not Found', 3);
  }
}
