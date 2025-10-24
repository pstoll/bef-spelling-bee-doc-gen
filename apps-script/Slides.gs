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

  // Template placeholder IDs (to be set after creating template)
  TEMPLATE: {
    // Will store the template presentation ID
    // This should be set in a script property or config
    PRESENTATION_ID: null
  }
};

// ============================================================================
// MAIN SLIDES GENERATION
// ============================================================================

/**
 * Generate Google Slides presentation for a round
 * @param {string} year - The year (e.g., "2024", "2019Fall")
 * @param {string} roundNumber - The round number
 * @param {Array} words - Array of word objects
 * @param {Folder} outputFolder - Google Drive folder to save to
 * @returns {File} The created presentation file
 */
function generateSlides(year, roundNumber, words, outputFolder) {
  try {
    // Create new presentation
    const presentation = SlidesApp.create(`BEF Spelling Bee ${year} - Round ${roundNumber}`);
    const presentationId = presentation.getId();

    Logger.log(`Created presentation: ${presentationId}`);

    // Get all slides (there's a default blank slide)
    let slides = presentation.getSlides();

    // Remove default slide if exists
    if (slides.length > 0) {
      slides[0].remove();
    }

    // Add title slide
    addTitleSlide(presentation, year, roundNumber);

    // Add "Spellers at Work" intro slide
    addInfoSlide(presentation, `Spellers at Work\n\n..quiet please...`, roundNumber);

    // Add word slides
    words.forEach((wordObj, index) => {
      addWordSlide(presentation, wordObj.word);

      // Add "Spellers at Work" between words (but not after last word)
      if (index < words.length - 1) {
        addInfoSlide(presentation, `Spellers at Work\n\n..quiet please...`);
      }
    });

    // Add concluding slide
    addInfoSlide(presentation, `This concludes\nRound ${roundNumber}`);

    // Move file to output folder
    const file = DriveApp.getFileById(presentationId);
    file.moveTo(outputFolder);

    Logger.log(`Slides generated successfully: ${file.getUrl()}`);
    return file;

  } catch (error) {
    Logger.log(`Error generating slides: ${error.message}`);
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
