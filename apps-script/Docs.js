// BEF Spelling Bee - Google Docs Generation

// ============================================================================
// DOCS CONFIGURATION
// ============================================================================

const DOCS_CONFIG = {
  // Template configuration
  TEMPLATE: {
    FILE_NAME: 'bef-bee-words-template'
  },

  // Typography
  FONT: {
    FAMILY: 'Calibri',
    FALLBACK: 'Arial',
    SIZE: {
      NORMAL: 20,
      SENTENCE: 20,
      TITLE_LARGE: 24,
      TITLE_MEDIUM: 18,
      FOOTER: 10
    }
  },

  // Formatting
  SPACING: {
    LINE_SPACING: 1.15,
    PARAGRAPH_SPACING_BEFORE: 0,
    PARAGRAPH_SPACING_AFTER: 10
  }
};

// ============================================================================
// MAIN DOCS GENERATION
// ============================================================================

/**
 * Generate Google Doc for a round
 * @param {string} year - The year (e.g., "2024", "2019Fall")
 * @param {string} roundNumber - The round number
 * @param {Array} words - Array of word objects
 * @param {Folder} outputFolder - Google Drive folder to save to
 * @returns {File} The created document file
 */
function generateDoc(year, roundNumber, words, outputFolder) {
  try {
    // Find template file
    const templateFile = findDocTemplateFile(DOCS_CONFIG.TEMPLATE.FILE_NAME);
    if (!templateFile) {
      Logger.log(`Error: Template file "${DOCS_CONFIG.TEMPLATE.FILE_NAME}" not found in same folder as spreadsheet`);
      return null;
    }

    // Copy template
    const fileName = `BEF Spelling Bee ${year} - Round ${roundNumber} - Words`;
    Logger.log(`Copying template document...`);

    // Delete any existing files with the same name in the folder
    const existingFiles = outputFolder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      const oldFile = existingFiles.next();
      Logger.log(`Deleting existing file: ${fileName} (${oldFile.getId()})`);
      oldFile.setTrashed(true);
    }

    const copy = templateFile.makeCopy(fileName, outputFolder);
    const docId = copy.getId();
    Logger.log(`Created document from template: ${docId}`);

    // Open the document
    Logger.log(`Opening document...`);
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    Logger.log(`Got document body, has ${body.getNumChildren()} children`);

    // Replace placeholders in title and footer (but NOT word entry placeholders yet)
    const now = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MM/dd/yyyy hh:mm a');

    Logger.log(`Replacing title/footer placeholders - year: ${year}, round: ${roundNumber}, date: ${dateStr}`);
    body.replaceText('\\{\\{year\\}\\}', year);
    body.replaceText('\\{\\{round\\}\\}', roundNumber);
    body.replaceText('\\{\\{created_date\\}\\}', dateStr);
    Logger.log(`Body title placeholder replacement complete`);

    // Also replace in footer
    const footer = doc.getFooter();
    if (footer) {
      Logger.log(`Replacing placeholders in footer...`);
      footer.replaceText('\\{\\{year\\}\\}', year);
      footer.replaceText('\\{\\{round\\}\\}', roundNumber);
      footer.replaceText('\\{\\{created_date\\}\\}', dateStr);
      Logger.log(`Footer placeholder replacement complete`);
    } else {
      Logger.log(`No footer found in document`);
    }

    // Find the template table and use it for word entries
    Logger.log(`Looking for template table to populate...`);
    const numChildren = body.getNumChildren();
    let templateTable = null;
    let templateTableIndex = -1;

    for (let i = 0; i < numChildren; i++) {
      const child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.TABLE) {
        const table = child.asTable();
        const tableText = table.getText();
        // Check if this is the template table
        if (tableText.includes('{{word}}')) {
          templateTable = table;
          templateTableIndex = i;
          Logger.log(`Found template table at index ${i}`);
          break;
        }
      }
    }

    if (!templateTable) {
      Logger.log(`ERROR: Template table not found!`);
      throw new Error('Template table with {{word}} placeholder not found');
    }

    // Get the template row (should be the first row)
    const templateRow = templateTable.getRow(0);

    // Make a pristine copy of the template row BEFORE any modifications
    const pristineTemplateRow = templateRow.copy();

    // Populate all words by duplicating and modifying the template row
    Logger.log(`Populating ${words.length} word entries...`);
    for (let i = 0; i < words.length; i++) {
      let targetRow, targetCell;

      if (i === 0) {
        // Use the existing template row for the first word
        targetRow = templateRow;
        targetCell = targetRow.getCell(0);
      } else {
        // For remaining words, copy the PRISTINE template row (before any modification)
        targetRow = pristineTemplateRow.copy();
        templateTable.appendTableRow(targetRow);
        targetCell = targetRow.getCell(0);
      }

      // Replace placeholders in the cell
      replaceInCell(targetCell, words[i], i + 1);
    }

    Logger.log(`Populated table with ${words.length} word entries`);

    // Save and close
    doc.saveAndClose();

    Logger.log(`Doc generated successfully: ${copy.getUrl()}`);
    return copy;

  } catch (error) {
    Logger.log(`Error generating doc: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    return null;
  }
}

// ============================================================================
// DOC FORMATTING FUNCTIONS
// ============================================================================

/**
 * Find doc template file in same folder as spreadsheet (case-insensitive)
 * @param {string} templateName - Template file name to find
 * @returns {File|null} The template file or null if not found
 */
function findDocTemplateFile(templateName) {
  // Get the spreadsheet's parent folder
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetFile = DriveApp.getFileById(ss.getId());
  const parentFolders = spreadsheetFile.getParents();

  if (!parentFolders.hasNext()) {
    Logger.log('Warning: Spreadsheet has no parent folder');
    return null;
  }

  const parentFolder = parentFolders.next();
  const templateNameLower = templateName.toLowerCase();

  Logger.log(`Searching for template "${templateName}" (lowercase: "${templateNameLower}") in folder: ${parentFolder.getName()}`);

  try {
    Logger.log(`DEBUG: About to list all files in folder...`);

    // Search for template file (case-insensitive)
    // Try getting all files first to see what's there
    const allFiles = parentFolder.getFiles();
    const allFilesList = [];
    while (allFiles.hasNext()) {
      const f = allFiles.next();
      allFilesList.push(`${f.getName()} (${f.getMimeType()})`);
    }
    Logger.log(`All files in folder: ${allFilesList.join(', ')}`);

    // Now search specifically for Google Docs
    Logger.log(`DEBUG: About to search for Google Docs...`);
    const files = parentFolder.getFilesByType(MimeType.GOOGLE_DOCS);
    const foundDocs = [];
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      foundDocs.push(fileName);
      Logger.log(`Checking doc: "${fileName}" vs template: "${templateName}"`);
      if (fileName.toLowerCase() === templateNameLower) {
        Logger.log(`Found template: ${fileName} in folder: ${parentFolder.getName()}`);
        return file;
      }
    }

    Logger.log(`Template "${templateName}" not found in folder: ${parentFolder.getName()}`);
    Logger.log(`Google Docs found in folder: ${foundDocs.join(', ')}`);
    return null;
  } catch (e) {
    Logger.log(`ERROR in findDocTemplateFile: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
    return null;
  }
}

/**
 * Replace placeholders in a table cell with actual word data
 * @param {TableCell} cell - Cell containing placeholders
 * @param {Object} wordObj - Word object with word, pronunciation, definition, sentence
 * @param {number} index - Word number (1-indexed)
 */
function replaceInCell(cell, wordObj, index) {
  // Get all paragraphs in the cell
  const numChildren = cell.getNumChildren();

  for (let i = 0; i < numChildren; i++) {
    const child = cell.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para = child.asParagraph();
      const paraText = para.editAsText();
      let text = paraText.getText();

      // Replace placeholders using replaceText on the paragraph
      if (text.includes('{{word}}')) {
        paraText.replaceText('\\{\\{word\\}\\}', `${index}. ${wordObj.word}`);
      }
      if (text.includes('{{pronunciation}}')) {
        paraText.replaceText('\\{\\{pronunciation\\}\\}', wordObj.pronunciation || '');
      }
      if (text.includes('{{definition}}')) {
        paraText.replaceText('\\{\\{definition\\}\\}', wordObj.definition || '');
      }
      if (text.includes('{{sentence}}')) {
        paraText.replaceText('\\{\\{sentence\\}\\}', wordObj.sentence || '');

        // Bold the word in the sentence
        if (wordObj.sentence) {
          // Refresh text after replacement
          text = paraText.getText();
          const sentenceLower = text.toLowerCase();
          const wordLower = wordObj.word.toLowerCase();
          const wordIndex = sentenceLower.indexOf(wordLower);
          if (wordIndex !== -1) {
            paraText.setBold(wordIndex, wordIndex + wordObj.word.length - 1, true);
          }
        }
      }
    }
  }
}

/**
 * Add a word entry to a table cell (keeps all parts together on same page)
 * Matches template format: {{word}}    {{pronunciation}}
 *                          {{definition}}
 *
 *                          {{sentence}}
 * @param {TableCell} cell - Table cell to add content to
 * @param {Object} wordObj - Word object with word, pronunciation, definition, sentence
 * @param {number} index - Word number (1-indexed)
 */
function addWordEntryToCell(cell, wordObj, index) {
  // Line 1: Word number and word (bold), tab, pronunciation
  const wordLine = cell.appendParagraph(`${index}. ${wordObj.word}\t${wordObj.pronunciation || ''}`);
  const wordText = wordLine.editAsText();

  // Bold just the word part (not the pronunciation)
  const wordEndIndex = `${index}. ${wordObj.word}`.length - 1;
  wordText.setBold(0, wordEndIndex, true);
  wordText.setFontFamily(DOCS_CONFIG.FONT.FAMILY);
  wordText.setFontSize(DOCS_CONFIG.FONT.SIZE.NORMAL);

  // Line 2: Definition
  if (wordObj.definition) {
    const defPara = cell.appendParagraph(wordObj.definition);
    defPara.editAsText().setFontFamily(DOCS_CONFIG.FONT.FAMILY);
    defPara.editAsText().setFontSize(DOCS_CONFIG.FONT.SIZE.NORMAL);
  }

  // Blank line
  cell.appendParagraph('');

  // Line 3: Sentence with word bolded
  if (wordObj.sentence) {
    const sentPara = cell.appendParagraph(wordObj.sentence);
    sentPara.editAsText().setFontFamily(DOCS_CONFIG.FONT.FAMILY);
    sentPara.editAsText().setFontSize(DOCS_CONFIG.FONT.SIZE.SENTENCE);
    sentPara.editAsText().setItalic(true);

    // Bold the word within the sentence
    boldWordInSentence(sentPara, wordObj.word, wordObj.sentence);
  }

}

/**
 * Bold a word within a sentence paragraph
 * @param {Paragraph} paragraph - The paragraph containing the sentence
 * @param {string} word - The word to bold
 * @param {string} sentence - The full sentence
 */
function boldWordInSentence(paragraph, word, sentence) {
  // Find the word in the sentence (case-insensitive)
  const sentenceLower = sentence.toLowerCase();
  const wordLower = word.toLowerCase();
  const wordIndex = sentenceLower.indexOf(wordLower);

  if (wordIndex === -1) {
    Logger.log(`Warning: Could not find word "${word}" in sentence "${sentence}"`);
    return;
  }

  // No prefix since sentence is on its own line
  const startIndex = wordIndex;
  const endIndex = startIndex + word.length - 1;

  // Apply bold formatting
  const text = paragraph.editAsText();
  text.setBold(startIndex, endIndex, true);
  text.setItalic(startIndex, endIndex, true); // Keep italic but also bold
}
