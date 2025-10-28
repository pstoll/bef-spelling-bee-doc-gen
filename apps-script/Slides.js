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
    TITLE_SLIDE_INDEX: 0,        // Slide 1: Title slide with {{year}}, {{round}}, {{date}}
    INTERSTITIAL_SLIDE_INDEX: 1, // Slide 2: "Spellers at Work" - hardcoded text
    WORD_SLIDE_INDEX: 2,         // Slide 3: Word slide with {{word}}
    CONCLUSION_SLIDE_INDEX: 3    // Slide 4: "This concludes Round {{round}}"
  },

  // Event configuration
  // Event date is read from ScriptConfig sheet (B1)
  // This is a fallback if ScriptConfig sheet is not found
  EVENT_DATE_FALLBACK: 'Nov 16, 2025'
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
    const templateFile = findSlideTemplateFile(SLIDES_CONFIG.TEMPLATE.FILE_NAME);
    if (!templateFile) {
      Logger.log(`Error: Template file "${SLIDES_CONFIG.TEMPLATE.FILE_NAME}" not found in same folder as spreadsheet`);
      return null;
    }

    // Copy template presentation using Drive API (DriveApp.makeCopy has issues with persistence)
    // Note: We use 3-param syntax because it's the only one that allows text replacements to persist
    // The name parameter is ignored in this syntax, so we rename afterward
    const fileName = `BEF Spelling Bee ${year} - Round ${roundNumber} - Slides`;

    // Delete any existing files with the same name in the folder
    const existingFiles = outputFolder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      const oldFile = existingFiles.next();
      Logger.log(`Deleting existing file: ${fileName} (${oldFile.getId()})`);
      oldFile.setTrashed(true);
    }

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
    const interstitialSlideTemplate = slides[SLIDES_CONFIG.TEMPLATE.INTERSTITIAL_SLIDE_INDEX];
    const wordSlideTemplate = slides[SLIDES_CONFIG.TEMPLATE.WORD_SLIDE_INDEX];
    const conclusionSlideTemplate = slides[SLIDES_CONFIG.TEMPLATE.CONCLUSION_SLIDE_INDEX];

    // Validate template structure
    Logger.log(`DEBUG: Validating template structure...`);
    const wordSlideText = wordSlideTemplate.getShapes().map(s => s.getText().asString()).join(' ');

    if (!wordSlideText.includes('{{word}}')) {
      Logger.log(`WARNING: Word slide (index ${SLIDES_CONFIG.TEMPLATE.WORD_SLIDE_INDEX}) does not contain {{word}} placeholder!`);
      Logger.log(`WARNING: Word slide contains: ${wordSlideText.substring(0, 100)}`);
      throw new Error(`Template validation failed: Word slide missing {{word}} placeholder. Check WORD_SLIDE_INDEX in config.`);
    }
    Logger.log(`DEBUG: Template validation passed`);

    slides = presentation.getSlides();
    Logger.log(`DEBUG: Using template slides directly: ${slides.length} slides`);
    Logger.log(`DEBUG: Will generate slides for ${words.length} words`);

    // Structure: [Title (0), Interstitial (1), Word template (2), Conclusion (3)]
    // Keep the first interstitial (slide 1) - it's the opening "Round X... quiet please" slide
    // Make copies of templates for duplication, then delete word template original
    Logger.log(`DEBUG: Making pristine copies of templates...`);
    const wordTemplateForDuplication = wordSlideTemplate.duplicate();
    const interstitialTemplateForDuplication = interstitialSlideTemplate.duplicate();

    Logger.log(`DEBUG: Deleting word template original...`);
    wordSlideTemplate.remove();
    // NOTE: We keep the original interstitial at index 1 as the opening slide

    // After deleting word template, structure is: [Title (0), Interstitial (1), Conclusion (2), copies...]
    // We'll insert word slides starting at index 2 (after opening interstitial)

    slides = presentation.getSlides();
    Logger.log(`DEBUG: After deleting word template, have ${slides.length} slides`);

    let currentIndex = 2; // Start after title (0) and opening interstitial (1)

    // Add word slides with interstitial slides
    let wordSlideCount = 0;
    let interstitialCount = 0;

    words.forEach((wordObj, index) => {
      try {
        // Add word slide at current position
        const wordSlide = wordTemplateForDuplication.duplicate();
        wordSlide.move(currentIndex);

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

        if (index === 0 || index === words.length - 1 || index === words.length - 2) {
          Logger.log(`DEBUG: Added word slide #${index + 1}: ${wordObj.word} at index ${currentIndex}`);
        }

        // Add "Spellers at Work" interstitial between words (but not after last word)
        // IMPORTANT: .move(X) inserts BEFORE index X, so to put interstitial AFTER the word we just added,
        // we need to move it to currentIndex + 1 (not currentIndex, which would put it BEFORE the word)
        if (index < words.length - 1) {
          const interstitialSlide = interstitialTemplateForDuplication.duplicate();
          interstitialSlide.move(currentIndex + 1);
          interstitialCount++;
          currentIndex += 2;  // Skip both word and interstitial
        } else {
          currentIndex++;  // Just skip the word
          Logger.log(`DEBUG: NOT adding interstitial after last word #${index + 1}: ${wordObj.word}`);
        }
      } catch (e) {
        Logger.log(`ERROR: Failed to create slide for word #${index + 1} (${wordObj.word}): ${e.message}`);
        throw e;
      }
    });

    slides = presentation.getSlides();
    Logger.log(`DEBUG: After word slides, now have ${slides.length} slides (added ${wordSlideCount} word slides, ${interstitialCount} interstitials)`);

    // Delete the pristine template copies we made for duplication (but NOT the original interstitial at index 1)
    Logger.log(`DEBUG: Deleting pristine template copies...`);
    wordTemplateForDuplication.remove();
    interstitialTemplateForDuplication.remove();

    // Move conclusion slide to end
    Logger.log(`DEBUG: Moving conclusion slide to end at index ${currentIndex}`);
    conclusionSlideTemplate.move(currentIndex);
    Logger.log(`DEBUG: Conclusion slide moved to position ${currentIndex}`);

    // Log the last few slides to verify structure
    slides = presentation.getSlides();
    Logger.log(`DEBUG: Total slides before save: ${slides.length}`);
    for (let i = Math.max(0, slides.length - 5); i < slides.length; i++) {
      const slide = slides[i];
      const slideText = slide.getShapes().map(s => s.getText().asString()).join(' ').substring(0, 50);
      Logger.log(`DEBUG: Slide ${i}: ${slideText}`);
    }

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
 * Find slide template file in same folder as spreadsheet (case-insensitive)
 * @param {string} templateName - Template file name to find
 * @returns {File|null} The template file or null if not found
 */
function findSlideTemplateFile(templateName) {
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
 * Get event date from ScriptConfig sheet
 * Reads from ScriptConfig sheet, row 1: A1="Event Date", B1=actual date
 */
function getEventDate(year) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('ScriptConfig');

    if (!configSheet) {
      Logger.log('WARNING: ScriptConfig sheet not found, using fallback date');
      return SLIDES_CONFIG.EVENT_DATE_FALLBACK;
    }

    // Read from B1 (A1 should be the label "Event Date")
    const eventDate = configSheet.getRange('B1').getValue();

    if (!eventDate) {
      Logger.log('WARNING: Event Date (B1) is empty in ScriptConfig sheet, using fallback');
      return SLIDES_CONFIG.EVENT_DATE_FALLBACK;
    }

    // Convert to string if it's a Date object
    if (eventDate instanceof Date) {
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      return `${months[eventDate.getMonth()]} ${eventDate.getDate()}, ${eventDate.getFullYear()}`;
    }

    return eventDate.toString();
  } catch (e) {
    Logger.log(`ERROR reading event date from ScriptConfig: ${e.message}`);
    return 'TBD';
  }
}
