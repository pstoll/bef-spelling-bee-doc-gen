// BEF Spelling Bee - Google Docs Generation

// ============================================================================
// DOCS CONFIGURATION
// ============================================================================

const DOCS_CONFIG = {
  // Typography
  FONT: {
    FAMILY: 'Calibri',
    FALLBACK: 'Arial',
    SIZE: {
      NORMAL: 20,
      SENTENCE: 20
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
    // Create new document
    const doc = DocumentApp.create(`BEF Spelling Bee ${year} - Round ${roundNumber} - Words`);
    const docId = doc.getId();
    const body = doc.getBody();

    Logger.log(`Created document: ${docId}`);

    // Clear default content
    body.clear();

    // Add header
    addDocHeader(body, year, roundNumber);

    // Add each word entry
    words.forEach((wordObj, index) => {
      addWordEntry(body, wordObj, index + 1);
    });

    // Save and close
    doc.saveAndClose();

    // Move file to output folder
    const file = DriveApp.getFileById(docId);
    file.moveTo(outputFolder);

    Logger.log(`Doc generated successfully: ${file.getUrl()}`);
    return file;

  } catch (error) {
    Logger.log(`Error generating doc: ${error.message}`);
    return null;
  }
}

// ============================================================================
// DOC FORMATTING FUNCTIONS
// ============================================================================

/**
 * Add document header
 */
function addDocHeader(body, year, roundNumber) {
  // Title
  const title = body.appendParagraph(`BEF 5th Grade Spelling Bee - ${year}`);
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  // Round subtitle
  const subtitle = body.appendParagraph(`Round ${roundNumber}`);
  subtitle.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  subtitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  // Spacer
  body.appendParagraph('');
}

/**
 * Add a word entry to the document
 * @param {Body} body - Document body
 * @param {Object} wordObj - Word object with word, pronunciation, definition, sentence
 * @param {number} index - Word number (1-indexed)
 */
function addWordEntry(body, wordObj, index) {
  // Word number and word (bold)
  const wordPara = body.appendParagraph(`${index}. ${wordObj.word}`);
  const wordText = wordPara.editAsText();
  wordText.setBold(true);
  wordText.setFontFamily(DOCS_CONFIG.FONT.FAMILY);
  wordText.setFontSize(DOCS_CONFIG.FONT.SIZE.NORMAL);

  // Pronunciation (italic)
  if (wordObj.pronunciation) {
    const pronPara = body.appendParagraph(`Pronunciation: ${wordObj.pronunciation}`);
    const pronText = pronPara.editAsText();
    pronText.setItalic(true);
    pronText.setFontFamily(DOCS_CONFIG.FONT.FAMILY);
    pronText.setFontSize(DOCS_CONFIG.FONT.SIZE.NORMAL);
  }

  // Definition
  if (wordObj.definition) {
    const defPara = body.appendParagraph(`Definition: ${wordObj.definition}`);
    defPara.editAsText().setFontFamily(DOCS_CONFIG.FONT.FAMILY);
    defPara.editAsText().setFontSize(DOCS_CONFIG.FONT.SIZE.NORMAL);
  }

  // Sentence with word bolded
  if (wordObj.sentence) {
    const sentPara = body.appendParagraph(`Sentence: ${wordObj.sentence}`);
    sentPara.editAsText().setFontFamily(DOCS_CONFIG.FONT.FAMILY);
    sentPara.editAsText().setFontSize(DOCS_CONFIG.FONT.SIZE.SENTENCE);
    sentPara.editAsText().setItalic(true);

    // Bold the word within the sentence
    boldWordInSentence(sentPara, wordObj.word, wordObj.sentence);
  }

  // Add spacing after entry
  body.appendParagraph('');
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

  // Calculate offset (account for "Sentence: " prefix)
  const prefix = 'Sentence: ';
  const startIndex = prefix.length + wordIndex;
  const endIndex = startIndex + word.length - 1;

  // Apply bold formatting
  const text = paragraph.editAsText();
  text.setBold(startIndex, endIndex, true);
  text.setItalic(startIndex, endIndex, true); // Keep italic but also bold
}

// ============================================================================
// TEST FUNCTION
// ============================================================================

/**
 * Test function - generate a sample doc
 */
function testGenerateDoc() {
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
  const file = generateDoc('TEST', '1', testWords, folder);

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
