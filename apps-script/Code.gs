// BEF Spelling Bee Generator - Apps Script
// Generates Google Slides and Docs from the spelling bee word database

// ============================================================================
// CONFIGURATION CONSTANTS
// ============================================================================

const CONFIG = {
  // Column names in the spreadsheet
  COLUMNS: {
    WORD: 'Word',
    PRONUNCIATION: 'Pronunciation',
    DEFINITION: 'Definition',
    SENTENCE: 'Sentence'
  },

  // Sheet name patterns
  SHEET_NAME_PATTERN: /^Round\s+(\d+)(\s|\+)?$/,
  ROUND_NUMBER_PATTERN: /Round\s+(\d+)/,

  // Output folder naming
  OUTPUT_FOLDER_PREFIX: 'BEF Spelling Bee',

  // Seeded random constants (for deterministic shuffle)
  RANDOM_SEED: {
    MULTIPLIER: 9301,
    INCREMENT: 49297,
    MODULUS: 233280
  },

  // Test defaults
  TEST: {
    DEFAULT_YEAR: '2024',
    DEFAULT_ROUND: 'Round 1'
  }
};

// ============================================================================
// MENU AND UI FUNCTIONS
// ============================================================================

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('BEF Spelling Bee')
    .addItem('Generate Materials...', 'showYearPrompt')
    .addSeparator()
    .addItem('Test: Read Round Data...', 'testReadRoundData')
    .addItem('Test: Generate Sample Slide', 'testGenerateSlide')
    .addToUi();
}

/**
 * Shows dialog to select year and generate materials
 */
function showYearPrompt() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Generate Spelling Bee Materials',
    'Enter the year (e.g., 2024, 2019Fall):',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    const year = result.getResponseText().trim();
    if (year) {
      generateMaterialsForYear(year);
    } else {
      ui.alert('Please enter a valid year');
    }
  }
}

// ============================================================================
// CORE GENERATION LOGIC
// ============================================================================

/**
 * Main function to generate materials for a specific year
 */
function generateMaterialsForYear(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Find all sheets that match "Round X" pattern
  const sheets = ss.getSheets();
  const roundSheets = sheets.filter(sheet => {
    return CONFIG.SHEET_NAME_PATTERN.test(sheet.getName());
  });

  if (roundSheets.length === 0) {
    ui.alert('No round sheets found. Make sure sheets are named like "Round 1", "Round 2", etc.');
    return;
  }

  Logger.log(`Found ${roundSheets.length} round sheets`);

  // Get or create output folder
  const outputFolder = getOrCreateOutputFolder(year);

  // Process each round
  const generatedFiles = [];
  roundSheets.forEach(sheet => {
    const roundNumber = extractRoundNumber(sheet.getName());
    Logger.log(`Processing Round ${roundNumber}`);

    const words = getRoundWords(sheet, year);
    Logger.log(`Found ${words.length} words for year ${year}`);

    if (words.length === 0) {
      Logger.log(`Skipping Round ${roundNumber} - no words found for year ${year}`);
      return;
    }

    // Shuffle words (deterministic using simple seed)
    const shuffledWords = shuffleArray(words, roundNumber);

    // Generate Slides
    const slidesFile = generateSlides(year, roundNumber, shuffledWords, outputFolder);
    if (slidesFile) {
      generatedFiles.push(`Round ${roundNumber} Slides: ${slidesFile.getUrl()}`);
    }

    // Generate Docs
    const docFile = generateDoc(year, roundNumber, shuffledWords, outputFolder);
    if (docFile) {
      generatedFiles.push(`Round ${roundNumber} Doc: ${docFile.getUrl()}`);
    }
  });

  // Show results
  if (generatedFiles.length > 0) {
    ui.alert(
      'Success!',
      `Generated ${generatedFiles.length} files in folder:\n${outputFolder.getUrl()}\n\n` +
      `Files:\n${generatedFiles.join('\n')}`,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert('No files generated. Check that the year column exists and has data.');
  }
}

// ============================================================================
// DATA EXTRACTION FUNCTIONS
// ============================================================================

/**
 * Extract round number from sheet name
 */
function extractRoundNumber(sheetName) {
  const match = sheetName.match(CONFIG.ROUND_NUMBER_PATTERN);
  return match ? match[1] : null;
}

/**
 * Read words from a round sheet for a specific year
 * @param {Sheet} sheet - The spreadsheet sheet to read from
 * @param {string} year - The year column to filter by (e.g., "2024", "2019Fall")
 * @returns {Array} Array of word objects
 */
function getRoundWords(sheet, year) {
  // Get all data
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    return []; // No data rows
  }

  // First row is headers
  const headers = data[0];

  // Find column index for the year
  const yearColIndex = headers.indexOf(year);
  if (yearColIndex === -1) {
    Logger.log(`Warning: Year column '${year}' not found in sheet ${sheet.getName()}`);
    return [];
  }

  // Find required columns
  const wordCol = headers.indexOf(CONFIG.COLUMNS.WORD);
  const pronCol = headers.indexOf(CONFIG.COLUMNS.PRONUNCIATION);
  const defCol = headers.indexOf(CONFIG.COLUMNS.DEFINITION);
  const sentCol = headers.indexOf(CONFIG.COLUMNS.SENTENCE);

  if (wordCol === -1 || pronCol === -1 || defCol === -1 || sentCol === -1) {
    Logger.log(`Warning: Required columns not found in ${sheet.getName()}`);
    Logger.log(`Expected: ${CONFIG.COLUMNS.WORD}, ${CONFIG.COLUMNS.PRONUNCIATION}, ${CONFIG.COLUMNS.DEFINITION}, ${CONFIG.COLUMNS.SENTENCE}`);
    return [];
  }

  // Filter and extract words
  const words = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Check if this word is active for this year
    const yearValue = row[yearColIndex];
    if (!yearValue || yearValue.toString().trim() === '') {
      continue; // Skip if year column is empty
    }

    // Extract word data
    const word = row[wordCol];
    if (!word || !word.toString().trim()) {
      continue; // Skip if no word
    }

    words.push({
      word: row[wordCol],
      pronunciation: row[pronCol] || '',
      definition: row[defCol] || '',
      sentence: row[sentCol] || ''
    });
  }

  return words;
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Simple deterministic shuffle using round number as seed
 * @param {Array} array - Array to shuffle
 * @param {number|string} seed - Seed value for deterministic shuffling
 * @returns {Array} Shuffled copy of array
 */
function shuffleArray(array, seed) {
  // Create a copy to avoid mutating original
  const shuffled = array.slice();

  // Simple seeded random using round number
  // This won't match the Node.js seedrandom exactly, but provides consistent shuffling
  let currentSeed = parseInt(seed) || 1;

  function seededRandom() {
    currentSeed = (currentSeed * CONFIG.RANDOM_SEED.MULTIPLIER + CONFIG.RANDOM_SEED.INCREMENT) % CONFIG.RANDOM_SEED.MODULUS;
    return currentSeed / CONFIG.RANDOM_SEED.MODULUS;
  }

  // Fisher-Yates shuffle with seeded random
  for (let i = shuffled.length - 1; i > 0; i--) {
    const j = Math.floor(seededRandom() * (i + 1));
    [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
  }

  return shuffled;
}

/**
 * Get or create output folder in Drive
 * @param {string} year - Year for the folder name
 * @returns {Folder} Google Drive folder
 */
function getOrCreateOutputFolder(year) {
  const folderName = `${CONFIG.OUTPUT_FOLDER_PREFIX} ${year}`;

  // Search for existing folder
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }

  // Create new folder
  return DriveApp.createFolder(folderName);
}

// ============================================================================
// TEST FUNCTIONS
// ============================================================================

/**
 * Test function - reads round data and logs it
 */
function testReadRoundData() {
  const ui = SpreadsheetApp.getUi();

  // Prompt for year
  const yearResult = ui.prompt(
    'Test: Read Round Data',
    `Enter the year to test (default: ${CONFIG.TEST.DEFAULT_YEAR}):`,
    ui.ButtonSet.OK_CANCEL
  );

  if (yearResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const year = yearResult.getResponseText().trim() || CONFIG.TEST.DEFAULT_YEAR;

  // Prompt for round
  const roundResult = ui.prompt(
    'Test: Read Round Data',
    `Enter the round sheet name (default: ${CONFIG.TEST.DEFAULT_ROUND}):`,
    ui.ButtonSet.OK_CANCEL
  );

  if (roundResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const roundName = roundResult.getResponseText().trim() || CONFIG.TEST.DEFAULT_ROUND;

  // Get the sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(roundName);

  if (!sheet) {
    ui.alert(`Sheet "${roundName}" not found`);
    return;
  }

  // Read words
  const words = getRoundWords(sheet, year);
  Logger.log(`Found ${words.length} words for year ${year} in ${roundName}`);

  if (words.length > 0) {
    Logger.log('First word:', JSON.stringify(words[0], null, 2));
    Logger.log('Last word:', JSON.stringify(words[words.length - 1], null, 2));
  }

  // Show in UI too
  const message = words.length > 0
    ? `Found ${words.length} words for year ${year} in ${roundName}.\n\n` +
      `First word: ${words[0].word}\n` +
      `Last word: ${words[words.length - 1].word}\n\n` +
      `Check View → Executions or Logs for full details.`
    : `No words found for year ${year} in ${roundName}.\n\n` +
      `Check that:\n` +
      `1. The year column "${year}" exists in the sheet\n` +
      `2. There are rows with non-empty values in that column`;

  ui.alert('Test Results', message, ui.ButtonSet.OK);
}
