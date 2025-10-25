// BEF Spelling Bee - Google Slides Generation

// ============================================================================
// SLIDES CONFIGURATION
// ============================================================================

const SLIDES_CONFIG = {
  // BEF Brand Colors
  COLORS: {
    GREEN: '#94C601',
    DARK_GREEN: '#74A50F',
    BLACK: '#000000',
    WHITE: '#FFFFFF',
    GRAY: '#71685A'
  },

  // Typography
  FONT: {
    FAMILY: 'Questrial',
    FALLBACK: 'Arial' // In case Questrial isn't available
  },

  // Layout (based on 4:3 aspect ratio like original PPTX)
  LAYOUT: {
    SLIDE_WIDTH: 720,  // 10 inches * 72 points
    SLIDE_HEIGHT: 540  // 7.5 inches * 72 points
  },

  // Template configuration
  TEMPLATE: {
    // Template file name (should be in same folder as spreadsheet, case-insensitive)
    FILE_NAME: 'bef-bee-slide-template',
    // Slide indices in the template (0-based)
    TITLE_SLIDE_INDEX: 0,
    WORD_SLIDE_INDEX: 2,  // Swapped: word is on slide 3 (index 2)
    INFO_SLIDE_INDEX: 1   // Swapped: info is on slide 2 (index 1)
  }
};

// ============================================================================
// MAIN SLIDES GENERATION
// ============================================================================

/**
 * Generate Google Slides presentation for a round using template
 * @param {string} year - The year (e.g., "2024", "2019Fall")
 * @param {string} roundNumber - The round number
 * @param {Array} words - Array of word objects
 * @param {Folder} outputFolder - Google Drive folder to save to
 * @returns {File} The created presentation file
 */
function generateSlides(year, roundNumber, words, outputFolder) {
  try {
    // Find template file in same folder as spreadsheet
    const templateFile = findTemplateFile(SLIDES_CONFIG.TEMPLATE.FILE_NAME);
    if (!templateFile) {
      Logger.log(`Error: Template file "${SLIDES_CONFIG.TEMPLATE.FILE_NAME}" not found in same folder as spreadsheet`);
      return null;
    }

    // Copy template presentation using Drive API (DriveApp.makeCopy has issues with persistence)
    // Note: We use 3-param syntax because it's the only one that allows text replacements to persist
    // The name parameter is ignored in this syntax, so we rename afterward
    const fileName = `BEF Spelling Bee ${year} - Round ${roundNumber} - Slides`;
    Logger.log(`Creating copy via Drive API (will rename after)...`);
    Logger.log(`Template file ID: ${templateFile.getId()}`);
    Logger.log(`Output folder ID: ${outputFolder.getId()}`);
    Logger.log(`Output folder name: ${outputFolder.getName()}`);

    // Use 3-param syntax (only one that allows edits to persist)
    const copy = Drive.Files.copy({}, templateFile.getId(), {
      parents: [{id: outputFolder.getId()}]
    });
    const copyId = copy.id;

    Logger.log(`Created copy with ID: ${copyId}, renaming to: ${fileName}`);

    // Check where the copy actually ended up immediately after creation
    const copyCheck = Drive.Files.get(copyId, {fields: 'parents'});
    Logger.log(`Copy created with parents: ${JSON.stringify(copyCheck.parents)}`);

    // Move file to correct folder if it's not already there
    const currentParentId = copyCheck.parents[0];
    const targetParentId = outputFolder.getId();
    if (currentParentId !== targetParentId) {
      Logger.log(`Moving file from ${currentParentId} to ${targetParentId}...`);

      // Use DriveApp to move the file
      const tempFile = DriveApp.getFileById(copyId);
      const oldParent = DriveApp.getFolderById(currentParentId);
      oldParent.removeFile(tempFile);
      outputFolder.addFile(tempFile);

      Logger.log(`File moved successfully`);
    }

    Logger.log(`Created presentation from template: ${copyId}`);

    // Wait for Google Drive to fully initialize the file
    waitForFileReady(copyId, 30);

    // Open the presentation
    const presentation = SlidesApp.openById(copyId);

    // Verify we can read from it
    const initialSlides = presentation.getSlides();
    Logger.log(`DEBUG: Opened presentation, initial slide count: ${initialSlides.length}`);

    // Replace global placeholders
    const eventDate = getEventDate(year);
    Logger.log(`DEBUG: Replacing placeholders - year: ${year}, round: ${roundNumber}, date: ${eventDate}`);
    const yearReplacements = presentation.replaceAllText('{{year}}', year);
    const roundReplacements = presentation.replaceAllText('{{round}}', roundNumber);
    const dateReplacements = presentation.replaceAllText('{{event_date}}', eventDate);
    Logger.log(`DEBUG: Replacements made - year: ${yearReplacements}, round: ${roundReplacements}, date: ${dateReplacements}`);

    // Get template slides
    let slides = presentation.getSlides();
    Logger.log(`DEBUG: Template has ${slides.length} slides after copy`);
    const titleSlide = slides[SLIDES_CONFIG.TEMPLATE.TITLE_SLIDE_INDEX];
    const wordSlideOriginal = slides[SLIDES_CONFIG.TEMPLATE.WORD_SLIDE_INDEX];
    const infoSlideOriginal = slides[SLIDES_CONFIG.TEMPLATE.INFO_SLIDE_INDEX];

    // Validate template structure
    Logger.log(`DEBUG: Validating template structure...`);
    const wordSlideText = wordSlideOriginal.getShapes().map(s => s.getText().asString()).join(' ');
    const infoSlideText = infoSlideOriginal.getShapes().map(s => s.getText().asString()).join(' ');

    if (!wordSlideText.includes('{{word}}')) {
      Logger.log(`WARNING: Word slide (index ${SLIDES_CONFIG.TEMPLATE.WORD_SLIDE_INDEX}) does not contain {{word}} placeholder!`);
      Logger.log(`WARNING: Word slide contains: ${wordSlideText.substring(0, 100)}`);
      throw new Error(`Template validation failed: Word slide missing {{word}} placeholder. Check WORD_SLIDE_INDEX in config.`);
    }
    if (!infoSlideText.includes('{{message}}')) {
      Logger.log(`WARNING: Info slide (index ${SLIDES_CONFIG.TEMPLATE.INFO_SLIDE_INDEX}) does not contain {{message}} placeholder!`);
      Logger.log(`WARNING: Info slide contains: ${infoSlideText.substring(0, 100)}`);
      throw new Error(`Template validation failed: Info slide missing {{message}} placeholder. Check INFO_SLIDE_INDEX in config.`);
    }
    Logger.log(`DEBUG: Template validation passed`);

    // Don't remove or duplicate anything yet - just use the templates as-is
    const wordSlideTemplate = wordSlideOriginal;
    const infoSlideTemplate = infoSlideOriginal;

    slides = presentation.getSlides();
    Logger.log(`DEBUG: Using template slides directly: ${slides.length} slides`);
    Logger.log(`DEBUG: Will generate slides for ${words.length} words`);

    // Title slide is already updated with placeholders - keep it

    // Track insertion index - start after title slide
    let currentIndex = 1;

    // Add intro "Spellers at Work" slide with round number
    Logger.log(`DEBUG: Creating intro slide at index ${currentIndex}`);
    try {
      const introSlide = infoSlideTemplate.duplicate();
      introSlide.move(currentIndex);
      currentIndex++;
      Logger.log(`DEBUG: Intro slide created and moved to position ${currentIndex - 1}`);

      // Replace text by finding and modifying shapes directly
      const shapes = introSlide.getShapes();
      shapes.forEach(shape => {
        const text = shape.getText();
        let currentText = text.asString();
        if (currentText.includes('{{message}}')) {
          currentText = currentText.replace('{{message}}', `Spellers at Work\n\n..quiet please...`);
        }
        if (currentText.includes('{{round}}')) {
          currentText = currentText.replace('{{round}}', `Round ${roundNumber}`);
        }
        text.setText(currentText);
      });

      slides = presentation.getSlides();
      Logger.log(`DEBUG: After intro slide, now have ${slides.length} slides`);
    } catch (e) {
      Logger.log(`ERROR: Failed to create intro slide: ${e.message}`);
      Logger.log(`ERROR: Stack: ${e.stack}`);
      throw e;
    }

    // Add word slides with interstitial slides
    let wordSlideCount = 0;
    let interstitialCount = 0;

    words.forEach((wordObj, index) => {
      try {
        // Add word slide at current position
        const wordSlide = wordSlideTemplate.duplicate();
        wordSlide.move(currentIndex);
        currentIndex++;

        // Replace text by finding and modifying shapes directly
        const wordShapes = wordSlide.getShapes();
        wordShapes.forEach(shape => {
          const text = shape.getText();
          const currentText = text.asString();
          if (currentText.includes('{{word}}')) {
            const newText = currentText.replace('{{word}}', wordObj.word);
            text.setText(newText);
          }
        });

        wordSlideCount++;

        if (index === 0 || index === words.length - 1) {
          Logger.log(`DEBUG: Added word slide #${index + 1}: ${wordObj.word} at index ${currentIndex - 1}`);
        }

        // Add "Spellers at Work" between words (but not after last word)
        if (index < words.length - 1) {
          const infoSlide = infoSlideTemplate.duplicate();
          infoSlide.move(currentIndex);
          currentIndex++;

          // Replace text by finding and modifying shapes directly
          const infoShapes = infoSlide.getShapes();
          infoShapes.forEach(shape => {
            const text = shape.getText();
            let currentText = text.asString();
            if (currentText.includes('{{message}}')) {
              currentText = currentText.replace('{{message}}', `Spellers at Work\n\n..quiet please...`);
            }
            if (currentText.includes('{{round}}')) {
              currentText = currentText.replace('{{round}}', '');
            }
            text.setText(currentText);
          });

          interstitialCount++;
        }
      } catch (e) {
        Logger.log(`ERROR: Failed to create slide for word #${index + 1} (${wordObj.word}): ${e.message}`);
        throw e;
      }
    });

    slides = presentation.getSlides();
    Logger.log(`DEBUG: After word slides, now have ${slides.length} slides (added ${wordSlideCount} word slides, ${interstitialCount} interstitials)`);

    // Add concluding slide
    Logger.log(`DEBUG: Creating conclusion slide`);
    const conclusionSlide = infoSlideTemplate.duplicate();

    // Replace text by finding and modifying shapes directly
    const conclusionShapes = conclusionSlide.getShapes();
    conclusionShapes.forEach(shape => {
      const text = shape.getText();
      let currentText = text.asString();
      if (currentText.includes('{{message}}')) {
        currentText = currentText.replace('{{message}}', `This concludes\nRound ${roundNumber}`);
      }
      if (currentText.includes('{{round}}')) {
        currentText = currentText.replace('{{round}}', '');
      }
      text.setText(currentText);
    });

    slides = presentation.getSlides();
    Logger.log(`DEBUG: After conclusion slide, now have ${slides.length} slides`);

    // Move conclusion slide to correct position (after all word slides)
    conclusionSlide.move(currentIndex);
    currentIndex++;
    Logger.log(`DEBUG: Moved conclusion slide to position ${currentIndex - 1}`);

    // Delete the template slides (they're still at the end)
    Logger.log(`DEBUG: Deleting template slides...`);
    slides = presentation.getSlides();

    // Find and delete template slides (word and info slide templates)
    // They should be the last remaining slides with {{word}} or {{message}} placeholders
    let deletedCount = 0;
    for (let i = slides.length - 1; i >= 0; i--) {
      const slide = slides[i];
      const slideText = slide.getShapes().map(s => s.getText().asString()).join(' ');
      if (slideText.includes('{{word}}') || slideText.includes('{{message}}')) {
        Logger.log(`DEBUG: Deleting template slide at index ${i}`);
        slide.remove();
        deletedCount++;
      }
    }
    Logger.log(`DEBUG: Deleted ${deletedCount} template slides`);

    // CRITICAL: Save and close to persist all changes
    Logger.log(`Calling saveAndClose() to persist all changes...`);
    presentation.saveAndClose();
    Logger.log(`saveAndClose() completed`);

    // Rename the file now that all edits are complete
    Logger.log(`Renaming file to: ${fileName}`);
    Drive.Files.update({name: fileName}, copyId);

    // Return as DriveApp File object for compatibility
    const file = DriveApp.getFileById(copyId);

    // Verify file is in correct folder
    const parents = file.getParents();
    while (parents.hasNext()) {
      const parent = parents.next();
      Logger.log(`File is in folder: ${parent.getName()} (${parent.getId()})`);
    }

    Logger.log(`Slides generated successfully: ${file.getUrl()}`);
    return file;

  } catch (error) {
    Logger.log(`Error generating slides: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
    return null;
  }
}

// ============================================================================
// SLIDE CREATION FUNCTIONS
// ============================================================================

/**
 * Add title slide
 */
function addTitleSlide(presentation, year, roundNumber) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  // TODO: Add BEF branding elements
  // For now, add text elements

  // Event date (will need to be configurable)
  const eventDate = getEventDate(year);
  const dateText = slide.insertTextBox(eventDate);
  positionElement(dateText, 396, 119, 252, 27); // x:55%, y:22%, w:35%, h:5%
  styleText(dateText.getText(), SLIDES_CONFIG.COLORS.WHITE, 24);

  // BEF Organization
  const befText = slide.insertTextBox('Brookline\nEducation\nFoundation');
  positionElement(befText, 396, 189, 252, 135); // x:55%, y:35%, w:35%, h:25%
  styleText(befText.getText(), SLIDES_CONFIG.COLORS.GREEN, 32);

  // Title
  const titleText = slide.insertTextBox(`5th Grade\nSpelling Bee\n\nRound ${roundNumber}`);
  positionElement(titleText, 396, 351, 252, 81); // x:55%, y:65%, w:35%, h:15%
  const titleStyle = titleText.getText().getTextStyle();
  titleStyle.setBold(true);
  titleStyle.setForegroundColor(SLIDES_CONFIG.COLORS.BLACK);
  titleStyle.setFontSize(22);
  titleStyle.setFontFamily(SLIDES_CONFIG.FONT.FAMILY);
}

/**
 * Add info/interstitial slide
 */
function addInfoSlide(presentation, text, roundNumber) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  // Main text
  const mainText = slide.insertTextBox(text);
  positionElement(mainText, 144, 216, 432, 144); // x:20%, y:40%, w:60%, h:27%
  styleText(mainText.getText(), SLIDES_CONFIG.COLORS.BLACK, 36, 'center');

  // Round number overlay (if provided)
  if (roundNumber) {
    const roundText = slide.insertTextBox(`Round ${roundNumber}`);
    positionElement(roundText, 144, 90, 432, 72); // x:20%, y:16.7%, w:60%, h:13.3%
    styleText(roundText.getText(), SLIDES_CONFIG.COLORS.GREEN, 54, 'center');
  }
}

/**
 * Add word slide
 */
function addWordSlide(presentation, word) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  // Word text - centered
  const wordText = slide.insertTextBox(word);
  positionElement(wordText, 144, 216, 432, 72); // x:20%, y:40%, w:60%, h:13.3%
  styleText(wordText.getText(), SLIDES_CONFIG.COLORS.BLACK, 48, 'center');
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Find template file in same folder as spreadsheet (case-insensitive)
 * @param {string} templateName - Template file name to find
 * @returns {File|null} The template file or null if not found
 */
function findTemplateFile(templateName) {
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

  // Search for template file (case-insensitive)
  const files = parentFolder.getFilesByType(MimeType.GOOGLE_SLIDES);
  while (files.hasNext()) {
    const file = files.next();
    if (file.getName().toLowerCase() === templateNameLower) {
      Logger.log(`Found template: ${file.getName()} in folder: ${parentFolder.getName()}`);
      return file;
    }
  }

  Logger.log(`Template "${templateName}" not found in folder: ${parentFolder.getName()}`);
  return null;
}

/**
 * Position an element on the slide
 * @param {Shape} element - The slide element
 * @param {number} left - Left position in points
 * @param {number} top - Top position in points
 * @param {number} width - Width in points
 * @param {number} height - Height in points
 */
function positionElement(element, left, top, width, height) {
  element.setLeft(left);
  element.setTop(top);
  element.setWidth(width);
  element.setHeight(height);
}

/**
 * Style text
 * @param {TextRange} textRange - The text range to style
 * @param {string} color - Hex color
 * @param {number} fontSize - Font size in points
 * @param {string} alignment - Text alignment ('left', 'center', 'right')
 */
function styleText(textRange, color, fontSize, alignment) {
  const style = textRange.getTextStyle();
  style.setForegroundColor(color);
  style.setFontSize(fontSize);
  style.setFontFamily(SLIDES_CONFIG.FONT.FAMILY);

  const paragraphStyle = textRange.getParagraphStyle();
  if (alignment === 'center') {
    paragraphStyle.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  } else if (alignment === 'right') {
    paragraphStyle.setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
  } else {
    paragraphStyle.setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
  }
}

/**
 * Get event date for a given year
 * TODO: Make this configurable
 */
function getEventDate(year) {
  // Default to current date format
  const now = new Date();
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${months[now.getMonth()]} ${now.getDate()}, ${now.getFullYear()}`;
}

/**
 * Test different Drive API copy syntaxes to find what works
 */
function testDriveCopySyntaxes() {
  try {
    const templateFile = findTemplateFile(SLIDES_CONFIG.TEMPLATE.FILE_NAME);
    if (!templateFile) {
      Logger.log('Template not found');
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetFile = DriveApp.getFileById(ss.getId());
    const parentFolder = spreadsheetFile.getParents().next();
    const folderId = parentFolder.getId();

    Logger.log(`Template ID: ${templateFile.getId()}`);
    Logger.log(`Folder ID: ${folderId}`);
    Logger.log(`Folder name: ${parentFolder.getName()}`);

    // Test 1: Three-parameter syntax with id wrapper
    Logger.log('\n=== TEST 1: Three params with {id: folderId} ===');
    const test1 = Drive.Files.copy({}, templateFile.getId(), {
      name: 'TEST1-three-params-id-wrapper',
      parents: [{id: folderId}]
    });
    Logger.log(`Created: ${test1.id}`);
    const file1 = DriveApp.getFileById(test1.id);
    Logger.log(`Name: ${file1.getName()}`);
    Logger.log(`Parent: ${file1.getParents().next().getName()}`);
    Utilities.sleep(2000);

    // Test 2: Two-parameter syntax with just folderId string
    Logger.log('\n=== TEST 2: Two params with folderId string ===');
    const test2 = Drive.Files.copy({
      name: 'TEST2-two-params-string',
      parents: [folderId]
    }, templateFile.getId());
    Logger.log(`Created: ${test2.id}`);
    const file2 = DriveApp.getFileById(test2.id);
    Logger.log(`Name: ${file2.getName()}`);
    Logger.log(`Parent: ${file2.getParents().next().getName()}`);
    Utilities.sleep(2000);

    // Test 3: Two-parameter syntax with id wrapper
    Logger.log('\n=== TEST 3: Two params with {id: folderId} ===');
    const test3 = Drive.Files.copy({
      name: 'TEST3-two-params-id-wrapper',
      parents: [{id: folderId}]
    }, templateFile.getId());
    Logger.log(`Created: ${test3.id}`);
    const file3 = DriveApp.getFileById(test3.id);
    Logger.log(`Name: ${file3.getName()}`);
    Logger.log(`Parent: ${file3.getParents().next().getName()}`);

    // Now test text replacement on each
    Logger.log('\n=== Testing text replacement on TEST1 ===');
    Utilities.sleep(5000);
    const pres1 = SlidesApp.openById(test1.id);
    const count1 = pres1.replaceAllText('{{event_date}}', 'REPLACED1');
    Logger.log(`Replacements: ${count1}`);
    const reopen1 = SlidesApp.openById(test1.id);
    const shape1 = reopen1.getSlides()[0].getShapes()[0];
    const text1 = shape1.getText().asString();
    Logger.log(`Text in file: ${text1.substring(0, 50)}`);

    SpreadsheetApp.getUi().alert('Test complete! Check logs for results.');

  } catch (e) {
    Logger.log(`Error: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
  }
}

/**
 * Test if calling replaceAllText on Presentation vs TextRange makes a difference
 */
function testReplaceAllTextMethods() {
  try {
    const templateFile = findTemplateFile(SLIDES_CONFIG.TEMPLATE.FILE_NAME);
    if (!templateFile) {
      Logger.log('Template not found');
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetFile = DriveApp.getFileById(ss.getId());
    const parentFolder = spreadsheetFile.getParents().next();

    // TEST A: Call replaceAllText on PRESENTATION object
    Logger.log('\n=== TEST A: Presentation.replaceAllText() ===');
    const copyA = Drive.Files.copy({}, templateFile.getId(), {
      parents: [{id: parentFolder.getId()}]
    });
    Utilities.sleep(5000);
    const presA = SlidesApp.openById(copyA.id);
    const countA = presA.replaceAllText('{{event_date}}', 'METHOD-A-PRESENTATION');
    Logger.log(`Replacements: ${countA}`);
    const fileA = DriveApp.getFileById(copyA.id);
    Logger.log(`File A: ${fileA.getName()}`);
    Logger.log(`URL A: ${fileA.getUrl()}`);

    Utilities.sleep(2000);

    // TEST B: Call replaceAllText on SLIDE object
    Logger.log('\n=== TEST B: Slide.replaceAllText() ===');
    const copyB = Drive.Files.copy({}, templateFile.getId(), {
      parents: [{id: parentFolder.getId()}]
    });
    Utilities.sleep(5000);
    const presB = SlidesApp.openById(copyB.id);
    const slideB = presB.getSlides()[0];
    const countB = slideB.replaceAllText('{{event_date}}', 'METHOD-B-SLIDE');
    Logger.log(`Replacements: ${countB}`);
    const fileB = DriveApp.getFileById(copyB.id);
    Logger.log(`File B: ${fileB.getName()}`);
    Logger.log(`URL B: ${fileB.getUrl()}`);

    Utilities.sleep(2000);

    // TEST C: Call replaceAllText on TEXT RANGE
    Logger.log('\n=== TEST C: TextRange.replaceAllText() ===');
    const copyC = Drive.Files.copy({}, templateFile.getId(), {
      parents: [{id: parentFolder.getId()}]
    });
    Utilities.sleep(5000);
    const presC = SlidesApp.openById(copyC.id);
    const slideC = presC.getSlides()[0];
    const shapesC = slideC.getShapes();
    let countC = 0;
    shapesC.forEach(shape => {
      const text = shape.getText();
      countC += text.replaceAllText('{{event_date}}', 'METHOD-C-TEXTRANGE');
    });
    Logger.log(`Replacements: ${countC}`);
    const fileC = DriveApp.getFileById(copyC.id);
    Logger.log(`File C: ${fileC.getName()}`);
    Logger.log(`URL C: ${fileC.getUrl()}`);

    SpreadsheetApp.getUi().alert(
      'Test complete! Check these 3 files:\n\n' +
      `A (Presentation): ${fileA.getUrl()}\n` +
      `Should say: METHOD-A-PRESENTATION\n\n` +
      `B (Slide): ${fileB.getUrl()}\n` +
      `Should say: METHOD-B-SLIDE\n\n` +
      `C (TextRange): ${fileC.getUrl()}\n` +
      `Should say: METHOD-C-TEXTRANGE`
    );

  } catch (e) {
    Logger.log(`Error: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
  }
}

/**
 * Test if TEST1 syntax actually persists text changes to the real file
 */
function testTextPersistence() {
  try {
    const templateFile = findTemplateFile(SLIDES_CONFIG.TEMPLATE.FILE_NAME);
    if (!templateFile) {
      Logger.log('Template not found');
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetFile = DriveApp.getFileById(ss.getId());
    const parentFolder = spreadsheetFile.getParents().next();

    // Create with TEST1 syntax (3-param)
    Logger.log('Creating with TEST1 syntax...');
    const copy = Drive.Files.copy({}, templateFile.getId(), {
      parents: [{id: parentFolder.getId()}]
    });
    const copyId = copy.id;
    Logger.log(`Created: ${copyId}`);

    // Wait
    Logger.log('Waiting 5 seconds...');
    Utilities.sleep(5000);

    // Open and check original text
    const pres1 = SlidesApp.openById(copyId);
    const shape1 = pres1.getSlides()[0].getShapes()[0];
    const originalText = shape1.getText().asString();
    Logger.log(`Original text in API: ${originalText.substring(0, 50)}`);

    // Replace text
    Logger.log('Replacing {{event_date}} with PERSISTENCE-TEST...');
    const count = pres1.replaceAllText('{{event_date}}', 'PERSISTENCE-TEST');
    Logger.log(`Replacements made: ${count}`);

    // Check text immediately in API
    const pres2 = SlidesApp.openById(copyId);
    const shape2 = pres2.getSlides()[0].getShapes()[0];
    const textInAPI = shape2.getText().asString();
    Logger.log(`Text in API after replacement: ${textInAPI.substring(0, 50)}`);

    // Wait and check again
    Logger.log('Waiting 3 more seconds...');
    Utilities.sleep(3000);

    const pres3 = SlidesApp.openById(copyId);
    const shape3 = pres3.getSlides()[0].getShapes()[0];
    const textAfterWait = shape3.getText().asString();
    Logger.log(`Text in API after wait: ${textAfterWait.substring(0, 50)}`);

    // Get file for user to check manually
    const file = DriveApp.getFileById(copyId);

    SpreadsheetApp.getUi().alert(
      `Test complete!\n\nFile: ${file.getName()}\n\nAPI says text is: ${textAfterWait.substring(0, 40)}\n\nNow MANUALLY open this file and check if it says PERSISTENCE-TEST:\n${file.getUrl()}\n\nDoes the actual file match what the API sees?`
    );

  } catch (e) {
    Logger.log(`Error: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
  }
}

/**
 * Test the rename workaround
 */
function testRenameWorkaround() {
  try {
    const templateFile = findTemplateFile(SLIDES_CONFIG.TEMPLATE.FILE_NAME);
    if (!templateFile) {
      Logger.log('Template not found');
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetFile = DriveApp.getFileById(ss.getId());
    const parentFolder = spreadsheetFile.getParents().next();

    // Step 1: Create with 3-param syntax (allows edits)
    Logger.log('Creating copy with 3-param syntax...');
    const copy = Drive.Files.copy({}, templateFile.getId(), {
      parents: [{id: parentFolder.getId()}]
    });
    const copyId = copy.id;
    Logger.log(`Created: ${copyId}`);

    // Step 2: Wait and open
    Utilities.sleep(5000);
    const pres = SlidesApp.openById(copyId);
    Logger.log(`Opened, has ${pres.getSlides().length} slides`);

    // Step 3: Make text replacement
    const count = pres.replaceAllText('{{event_date}}', 'WORKAROUND-TEST');
    Logger.log(`Replaced ${count} instances`);

    // Step 4: Rename the file
    Logger.log('Renaming file...');
    Drive.Files.update({name: 'RENAME-WORKAROUND-TEST'}, copyId);

    // Step 5: Verify everything
    const reopened = SlidesApp.openById(copyId);
    const text = reopened.getSlides()[0].getShapes()[0].getText().asString();
    const file = DriveApp.getFileById(copyId);

    Logger.log(`Final name: ${file.getName()}`);
    Logger.log(`Text in file: ${text.substring(0, 50)}`);

    SpreadsheetApp.getUi().alert(
      `Test complete!\n\nName: ${file.getName()}\nText: ${text.substring(0, 30)}\n\nCheck file: ${file.getUrl()}`
    );

  } catch (e) {
    Logger.log(`Error: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
  }
}

/**
 * Wait for a Drive file to be fully initialized (have a modifiedTime)
 */
function waitForFileReady(fileId, maxWaitSeconds = 30) {
  Logger.log(`Waiting for file ${fileId} to be ready...`);
  const startTime = Date.now();
  let attempts = 0;

  while (Date.now() - startTime < maxWaitSeconds * 1000) {
    attempts++;
    const fileCheck = Drive.Files.get(fileId, {fields: 'id,name,modifiedTime'});
    Logger.log(`Attempt ${attempts}: modifiedTime = ${fileCheck.modifiedTime}`);

    if (fileCheck.modifiedTime) {
      Logger.log(`File ready after ${attempts} attempts (${Date.now() - startTime}ms)`);
      return true;
    }

    Utilities.sleep(1000); // Wait 1 second between checks
  }

  Logger.log(`WARNING: File not ready after ${maxWaitSeconds} seconds`);
  return false;
}

/**
 * Test creating a brand new presentation (not a copy) to see if THAT can be modified
 */
function testCreateNewPresentation() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetFile = DriveApp.getFileById(ss.getId());
    const parentFolder = spreadsheetFile.getParents().next();

    // Create a BRAND NEW presentation from scratch
    Logger.log('Creating brand new presentation...');
    const newPres = SlidesApp.create('BRAND-NEW-TEST');
    const presId = newPres.getId();
    Logger.log(`Created new presentation: ${presId}`);

    // Move it to the right folder
    const file = DriveApp.getFileById(presId);
    file.moveTo(parentFolder);
    Logger.log(`Moved to folder: ${parentFolder.getName()}`);

    // Use the EXISTING placeholder shapes instead of inserting new ones
    const slide = newPres.getSlides()[0];
    const shapes = slide.getShapes();
    Logger.log(`Slide has ${shapes.length} existing shapes`);

    // Modify the first placeholder
    if (shapes.length > 0) {
      const firstShape = shapes[0];
      firstShape.getText().setText('PLACEHOLDER-TEXT');
      Logger.log('Set text on existing shape to PLACEHOLDER-TEXT');
    } else {
      Logger.log('ERROR: No existing shapes to modify!');
      throw new Error('No shapes available');
    }

    // Wait BEFORE reopening (maybe Google needs time to queue the changes?)
    Logger.log('Waiting 5 seconds before reopen...');
    Utilities.sleep(5000);

    // REOPEN to trigger save!
    Logger.log('Reopening to trigger save...');
    const afterInsert = SlidesApp.openById(presId);
    Logger.log('Reopened after insert');

    // Wait again
    Logger.log('Waiting 5 seconds...');
    Utilities.sleep(5000);

    // Now replace text
    Logger.log('Replacing PLACEHOLDER-TEXT with FINAL-TEXT...');
    const count = afterInsert.replaceAllText('PLACEHOLDER-TEXT', 'FINAL-TEXT');
    Logger.log(`Replaced ${count} instances`);

    // Wait before final reopen
    Logger.log('Waiting 5 seconds before final reopen...');
    Utilities.sleep(5000);

    // REOPEN AGAIN to trigger save of the replacement!
    Logger.log('Reopening again to trigger save of replacement...');
    const afterReplace = SlidesApp.openById(presId);
    Logger.log('Reopened after replace');

    Logger.log('Waiting 5 seconds for final save...');
    Utilities.sleep(5000);

    // Check what's there
    const finalShapes = afterReplace.getSlides()[0].getShapes();
    Logger.log(`Final presentation has ${finalShapes.length} shapes`);

    const finalText = finalShapes.length > 0 ? finalShapes[0].getText().asString() : '(no shapes)';
    Logger.log(`Final text: ${finalText}`);

    SpreadsheetApp.getUi().alert(
      `Brand new presentation test!\n\n` +
      `API says: ${finalText}\n\n` +
      `Shapes found: ${shapes.length}\n\n` +
      `Check file: ${file.getUrl()}\n\n` +
      `Does it have a text box?`
    );

  } catch (e) {
    Logger.log(`Error: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
  }
}

/**
 * Test using DriveApp.makeCopy instead of Drive.Files.copy
 */
function testDriveAppMakeCopy() {
  try {
    const templateFile = findTemplateFile(SLIDES_CONFIG.TEMPLATE.FILE_NAME);
    if (!templateFile) {
      Logger.log('Template not found');
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetFile = DriveApp.getFileById(ss.getId());
    const parentFolder = spreadsheetFile.getParents().next();

    // Use DriveApp.makeCopy (old method)
    Logger.log('Creating copy via DriveApp.makeCopy...');
    const copy = templateFile.makeCopy('DRIVEAPP-MAKECOPY-TEST', parentFolder);
    const copyId = copy.getId();
    Logger.log(`Created copy: ${copyId}`);

    // Wait for it to be ready
    const isReady = waitForFileReady(copyId, 30);
    if (!isReady) {
      throw new Error('File never became ready');
    }

    // Now modify
    const pres = SlidesApp.openById(copyId);
    Logger.log(`Opened, has ${pres.getSlides().length} slides`);

    const count = pres.replaceAllText('{{event_date}}', 'DRIVEAPP-TEST');
    Logger.log(`Replaced ${count} instances`);

    Utilities.sleep(3000);

    // Check
    const reopened = SlidesApp.openById(copyId);
    const text = reopened.getSlides()[0].getShapes()[0].getText().asString();
    Logger.log(`API says text is: ${text.substring(0, 50)}`);

    SpreadsheetApp.getUi().alert(
      `DriveApp.makeCopy test complete!\n\n` +
      `API says: ${text.substring(0, 40)}\n\n` +
      `Check file: ${copy.getUrl()}\n\n` +
      `Does it show DRIVEAPP-TEST?`
    );

  } catch (e) {
    Logger.log(`Error: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
  }
}

/**
 * MINIMAL test - just copy template and replace one text
 */
function testMinimalCopy() {
  try {
    // Find template
    const templateFile = findTemplateFile(SLIDES_CONFIG.TEMPLATE.FILE_NAME);
    if (!templateFile) {
      Logger.log('Template not found');
      return;
    }

    // Get parent folder
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetFile = DriveApp.getFileById(ss.getId());
    const parentFolder = spreadsheetFile.getParents().next();

    // Make copy using Drive API directly (using the syntax that works!)
    Logger.log(`Creating copy via Drive API...`);
    const copy = Drive.Files.copy({}, templateFile.getId(), {
      name: 'MINIMAL TEST',
      parents: [{id: parentFolder.getId()}]
    });
    const copyId = copy.id;
    Logger.log(`Created copy: ${copyId}`);

    // Wait for file to be ACTUALLY ready (have a modifiedTime)
    const isReady = waitForFileReady(copyId, 30);
    if (!isReady) {
      throw new Error('File never became ready');
    }

    // Verify file exists and is ready via Drive API
    const fileCheck = Drive.Files.get(copyId);
    Logger.log(`File status: ${fileCheck.mimeType}, modifiedTime: ${fileCheck.modifiedTime}`);

    // Now open with SlidesApp
    const pres = SlidesApp.openById(copyId);
    const presId = pres.getId();
    const presUrl = pres.getUrl();
    Logger.log(`Opened presentation ID: ${presId}`);
    Logger.log(`Presentation URL: ${presUrl}`);
    Logger.log(`IDs match: ${copyId === presId}`);
    Logger.log(`Opened, has ${pres.getSlides().length} slides`);

    // Replace ONE thing that we know exists
    const count = pres.replaceAllText('{{event_date}}', 'TEST123');
    Logger.log(`Replaced ${count} instances of {{event_date}} with TEST123`);

    // Try to force a save by making another tiny change
    pres.replaceAllText('TEST123', 'TEST123'); // No-op replacement to trigger save

    // Force flush by waiting
    Utilities.sleep(3000); // Wait 3 seconds for Google to save

    Logger.log(`Waited 3 seconds for save to complete`);

    // Check
    const reopened = SlidesApp.openById(copyId);
    Logger.log(`Reopened, has ${reopened.getSlides().length} slides`);

    // Try to read text back
    const slides = reopened.getSlides();
    if (slides.length > 0) {
      const firstSlide = slides[0];
      const shapes = firstSlide.getShapes();
      Logger.log(`First slide has ${shapes.length} shapes`);
      if (shapes.length > 0) {
        const textContent = shapes[0].getText().asString();
        Logger.log(`First shape text: ${textContent.substring(0, 100)}`);
      }
    }

    SpreadsheetApp.getUi().alert(`Minimal test complete.\n\nFile ID: ${copyId}\n\nCheck file: ${presUrl}\n\nDoes it say TEST123?`);

  } catch (e) {
    Logger.log(`Error: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
  }
}

/**
 * Test function - generate a sample slide
 */
function testGenerateSlide() {
  const testWords = [
    { word: 'example', pronunciation: 'ig-ZAM-pul', definition: 'A thing serving to illustrate a rule', sentence: 'This is an example sentence.' },
    { word: 'spelling', pronunciation: 'SPEL-ing', definition: 'The process of writing words', sentence: 'The spelling bee is tomorrow.' }
  ];

  const folder = getOrCreateOutputFolder('TEST');
  const file = generateSlides('TEST', '1', testWords, folder);

  if (file) {
    SpreadsheetApp.getUi().alert(
      'Test Slide Created',
      `Sample slide created successfully!\n\nView it here:\n${file.getUrl()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert('Error creating test slide. Check the logs.');
  }
}
