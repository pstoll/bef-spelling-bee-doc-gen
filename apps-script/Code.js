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
    .addItem('Debug: Show Sheet Structure', 'debugShowSheetStructure')
    .addSeparator()
    .addItem('Test: Generate Sample Doc', 'testGenerateSampleDoc')
    .addItem('Test: Simple Copy & Replace', 'testSimpleCopyAndReplace')
    .addItem('Test: Final Workflow Test', 'testFinalWorkflow')
    .addToUi();
}

/**
 * Shows dialog to select year and generate materials
 */
function showYearPrompt() {
  const ui = SpreadsheetApp.getUi();

  // Ask for year
  const yearResult = ui.prompt(
    'Generate Spelling Bee Materials',
    'Enter the year (e.g., 2024, 2019Fall):',
    ui.ButtonSet.OK_CANCEL
  );

  if (yearResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const year = yearResult.getResponseText().trim();
  if (!year) {
    ui.alert('Please enter a valid year');
    return;
  }

  // Ask which rounds to generate
  const roundResult = ui.prompt(
    'Select Rounds',
    'Enter round numbers to generate (e.g., "1" or "1,2,3" or "all"):',
    ui.ButtonSet.OK_CANCEL
  );

  if (roundResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const roundInput = roundResult.getResponseText().trim().toLowerCase();

  if (roundInput === 'all') {
    generateMaterialsForYear(year);
  } else {
    // Parse comma-separated round numbers
    const roundNumbers = roundInput.split(',').map(r => r.trim()).filter(r => r);
    if (roundNumbers.length === 0) {
      ui.alert('Please enter valid round numbers');
      return;
    }
    generateMaterialsForYear(year, roundNumbers);
  }
}

// ============================================================================
// CORE GENERATION LOGIC
// ============================================================================

/**
 * Main function to generate materials for a specific year
 * @param {string} year - The year to generate for
 * @param {Array} specificRounds - Optional array of round numbers to generate (e.g., ["1", "2"])
 */
function generateMaterialsForYear(year, specificRounds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Find all sheets that match "Round X" pattern
  const sheets = ss.getSheets();
  let roundSheets = sheets.filter(sheet => {
    return CONFIG.SHEET_NAME_PATTERN.test(sheet.getName());
  });

  if (roundSheets.length === 0) {
    ui.alert('No round sheets found. Make sure sheets are named like "Round 1", "Round 2", etc.');
    return;
  }

  // Filter to specific rounds if requested
  if (specificRounds && specificRounds.length > 0) {
    roundSheets = roundSheets.filter(sheet => {
      const roundNumber = extractRoundNumber(sheet.getName());
      return specificRounds.includes(roundNumber);
    });

    if (roundSheets.length === 0) {
      ui.alert(`No sheets found for rounds: ${specificRounds.join(', ')}`);
      return;
    }

    Logger.log(`Filtered to ${roundSheets.length} round sheets: ${specificRounds.join(', ')}`);
  } else {
    Logger.log(`Found ${roundSheets.length} round sheets`);
  }

  // Get or create output folder
  const outputFolder = getOrCreateOutputFolder(year);

  // Process each round
  const generatedFiles = [];
  const errors = [];

  roundSheets.forEach(sheet => {
    const roundNumber = extractRoundNumber(sheet.getName());
    Logger.log(`Processing Round ${roundNumber}`);

    try {
      const words = getRoundWords(sheet, year);
      Logger.log(`Found ${words.length} words for year ${year}`);

      if (words.length === 0) {
        Logger.log(`Skipping Round ${roundNumber} - no words found for year ${year}`);
        errors.push(`Round ${roundNumber}: No words found`);
        return;
      }

      // Shuffle words (deterministic using simple seed)
      const shuffledWords = shuffleArray(words, roundNumber);

      // Generate Slides
      const slidesFile = generateSlides(year, roundNumber, shuffledWords, outputFolder);
      if (slidesFile) {
        generatedFiles.push(`Round ${roundNumber} Slides: ${slidesFile.getUrl()}`);
      } else {
        errors.push(`Round ${roundNumber}: Slides generation failed`);
      }

      // Generate Docs
      const docFile = generateDoc(year, roundNumber, shuffledWords, outputFolder);
      if (docFile) {
        generatedFiles.push(`Round ${roundNumber} Doc: ${docFile.getUrl()}`);
      } else {
        errors.push(`Round ${roundNumber}: Doc generation failed`);
      }
    } catch (e) {
      Logger.log(`ERROR processing Round ${roundNumber}: ${e.message}`);
      errors.push(`Round ${roundNumber}: ${e.message}`);
    }
  });

  // Show results
  let message = '';
  if (generatedFiles.length > 0) {
    message += `Generated ${generatedFiles.length} files in folder:\n${outputFolder.getUrl()}\n\n`;
    message += `Files:\n${generatedFiles.join('\n')}`;
  }

  if (errors.length > 0) {
    if (message) message += '\n\n';
    message += `Errors:\n${errors.join('\n')}`;
  }

  if (!message) {
    message = 'No files generated. Check that the year column exists and has data.';
  }

  ui.alert(generatedFiles.length > 0 ? 'Generation Complete' : 'Generation Failed', message, ui.ButtonSet.OK);
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
    Logger.log(`No data rows in sheet ${sheet.getName()}`);
    return []; // No data rows
  }

  // First row is headers - convert to strings for comparison
  const headers = data[0].map(h => h.toString());
  Logger.log(`Headers in ${sheet.getName()}: ${headers.join(', ')}`);

  // Find column index for the year
  const yearColIndex = headers.indexOf(year);
  if (yearColIndex === -1) {
    Logger.log(`ERROR: Year column '${year}' not found in sheet ${sheet.getName()}`);
    Logger.log(`Available headers: ${headers.join(', ')}`);
    return [];
  }

  Logger.log(`Found year column '${year}' at index ${yearColIndex}`);

  // Find required columns
  const wordCol = headers.indexOf(CONFIG.COLUMNS.WORD);
  const pronCol = headers.indexOf(CONFIG.COLUMNS.PRONUNCIATION);
  const defCol = headers.indexOf(CONFIG.COLUMNS.DEFINITION);
  const sentCol = headers.indexOf(CONFIG.COLUMNS.SENTENCE);

  if (wordCol === -1 || pronCol === -1 || defCol === -1 || sentCol === -1) {
    Logger.log(`Warning: Required columns not found in ${sheet.getName()}`);
    Logger.log(`Expected: ${CONFIG.COLUMNS.WORD}, ${CONFIG.COLUMNS.PRONUNCIATION}, ${CONFIG.COLUMNS.DEFINITION}, ${CONFIG.COLUMNS.SENTENCE}`);
    Logger.log(`Found indices: Word=${wordCol}, Pron=${pronCol}, Def=${defCol}, Sent=${sentCol}`);
    return [];
  }

  // Filter and extract words
  const words = [];
  let skippedEmpty = 0;
  let skippedNoWord = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Check if this word is active for this year
    const yearValue = row[yearColIndex];
    if (!yearValue || yearValue.toString().trim() === '') {
      skippedEmpty++;
      continue; // Skip if year column is empty
    }

    // Extract word data
    const word = row[wordCol];
    if (!word || !word.toString().trim()) {
      skippedNoWord++;
      continue; // Skip if no word
    }

    words.push({
      word: row[wordCol],
      pronunciation: row[pronCol] || '',
      definition: row[defCol] || '',
      sentence: row[sentCol] || ''
    });
  }

  Logger.log(`Sheet ${sheet.getName()}, Year ${year}: Found ${words.length} words, skipped ${skippedEmpty} (no year marker), skipped ${skippedNoWord} (no word)`);
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
 * Get or create output folder in Drive (in same folder as spreadsheet)
 * @param {string} year - Year for the folder name
 * @returns {Folder} Google Drive folder
 */
function getOrCreateOutputFolder(year) {
  const folderName = `${CONFIG.OUTPUT_FOLDER_PREFIX} ${year}`;

  // Get the spreadsheet's parent folder
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetFile = DriveApp.getFileById(ss.getId());
  const parentFolders = spreadsheetFile.getParents();

  if (!parentFolders.hasNext()) {
    Logger.log('Warning: Spreadsheet has no parent folder, creating folder at root');
    return DriveApp.createFolder(folderName);
  }

  const parentFolder = parentFolders.next();

  // Search for existing folder in the parent folder
  const existingFolders = parentFolder.getFoldersByName(folderName);
  if (existingFolders.hasNext()) {
    Logger.log(`Using existing folder: ${folderName}`);
    return existingFolders.next();
  }

  // Create new folder in the parent folder
  Logger.log(`Creating new folder: ${folderName} in ${parentFolder.getName()}`);
  return parentFolder.createFolder(folderName);
}

// ============================================================================
// DEBUG FUNCTIONS
// ============================================================================

/**
 * Debug function to show what columns and data exist
 */
function debugShowSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Get first Round sheet
  const sheets = ss.getSheets();
  const roundSheet = sheets.find(sheet => {
    return CONFIG.SHEET_NAME_PATTERN.test(sheet.getName());
  });

  if (!roundSheet) {
    ui.alert('No round sheet found. Make sure you have a sheet named "Round 1" or similar.');
    return;
  }

  const data = roundSheet.getDataRange().getValues();
  if (data.length < 2) {
    ui.alert('Sheet has no data rows');
    return;
  }

  const headers = data[0];
  const firstDataRow = data[1];

  let message = `Sheet: ${roundSheet.getName()}\n\n`;
  message += `Headers found:\n${headers.join(', ')}\n\n`;
  message += `First data row:\n`;

  headers.forEach((header, idx) => {
    message += `${header}: ${firstDataRow[idx]}\n`;
  });

  // Show which columns might be year columns
  const yearLikeHeaders = headers.filter(h => {
    const str = h.toString().trim();
    return str.match(/^\d{4}/) || str.toLowerCase().includes('fall');
  });

  message += `\n\nPossible year columns:\n${yearLikeHeaders.join(', ')}`;

  Logger.log(message);
  ui.alert('Sheet Structure Debug', message, ui.ButtonSet.OK);
}

function testSimpleCopyAndReplace() {
  // Simplest possible test: copy template, do ONE replacement, wait 30 seconds
  SpreadsheetApp.getUi().alert('Simple test: copy + one replacement. Takes 40 seconds.');

  Logger.log('Test: Copy template and do one global replacement...');

  const templateFile = findSlideTemplateFile(SLIDES_CONFIG.TEMPLATE.FILE_NAME);
  if (!templateFile) {
    SpreadsheetApp.getUi().alert('Template not found!');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetFile = DriveApp.getFileById(ss.getId());
  const parentFolder = spreadsheetFile.getParents().next();

  // Copy using DriveApp.makeCopy (simpler)
  const copy = templateFile.makeCopy('SIMPLE-COPY-TEST', parentFolder);
  const copyId = copy.getId();
  Logger.log(`Created copy: ${copyId}`);

  // Wait for it to be ready
  waitForFileReady(copyId, 30);

  // Open and do ONE replacement
  const pres = SlidesApp.openById(copyId);
  const count = pres.replaceAllText('{{event_date}}', 'SIMPLE-TEST-WORKED');
  Logger.log(`Replaced ${count} instances`);

  // SAVE AND CLOSE!
  Logger.log('Calling saveAndClose()...');
  pres.saveAndClose();
  Logger.log('saveAndClose() completed');

  // Wait a bit
  Logger.log('Waiting 5 seconds...');
  Utilities.sleep(5000);

  SpreadsheetApp.getUi().alert(
    'Simple test complete!\n\n' +
    'File: ' + copy.getUrl() + '\n\n' +
    'Should show SIMPLE-TEST-WORKED instead of {{event_date}}'
  );
}

function testFinalWorkflow() {
  // Test with just 3 words to verify the full workflow works
  SpreadsheetApp.getUi().alert('This will take about 45 seconds. Click OK to start.');

  Logger.log('Starting final workflow test with 3 words...');

  const testWords = [
    { word: 'apple' },
    { word: 'banana' },
    { word: 'cherry' }
  ];

  // Get output folder
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetFile = DriveApp.getFileById(ss.getId());
  const parentFolder = spreadsheetFile.getParents().next();

  // Generate slides
  const file = generateSlides('TEST', '99', testWords, parentFolder);

  if (file) {
    SpreadsheetApp.getUi().alert(
      'Test complete!\n\n' +
      'Check file: ' + file.getUrl() + '\n\n' +
      'Should have:\n' +
      '- Title slide\n' +
      '- Intro "Spellers at Work"\n' +
      '- Slide with "apple"\n' +
      '- "Spellers at Work"\n' +
      '- Slide with "banana"\n' +
      '- "Spellers at Work"\n' +
      '- Slide with "cherry"\n' +
      '- Conclusion "This concludes Round 99"\n\n' +
      'Total: ~9 slides (plus 2 template slides at end)'
    );
  } else {
    SpreadsheetApp.getUi().alert('Test failed! Check logs.');
  }
}

function testGenerateSampleDoc() {
  const testWords = [
    {
      word: 'example',
      pronunciation: 'ig-ZAM-pul',
      definition: 'A thing characteristic of its kind or illustrating a general rule.',
      sentence: 'This is an example sentence with the word example in it.'
    },
    {
      word: 'spelling',
      pronunciation: 'SPEL-ing',
      definition: 'The process or activity of writing or naming the letters of a word.',
      sentence: 'The spelling bee is a competition where students demonstrate their spelling skills.'
    }
  ];

  const folder = getOrCreateOutputFolder('TEST');
  const file = generateDoc('TEST', '99', testWords, folder);

  if (file) {
    SpreadsheetApp.getUi().alert(
      'Test Doc Created',
      `Sample doc created successfully!\n\nView it here:\n${file.getUrl()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert('Error creating test doc. Check the logs.');
  }
}
