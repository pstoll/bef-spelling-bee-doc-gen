// Restore onOpen and other functions
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('BEF Spelling Bee')
    .addItem('Generate Materials...', 'showYearPrompt')
    .addSeparator()
    .addItem('Test: Simple Copy & Replace', 'testSimpleCopyAndReplace')
    .addItem('Test: Final Workflow Test', 'testFinalWorkflow')
    .addToUi();
}

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

function testSimpleCopyAndReplace() {
  // Simplest possible test: copy template, do ONE replacement, wait 30 seconds
  SpreadsheetApp.getUi().alert('Simple test: copy + one replacement. Takes 40 seconds.');

  Logger.log('Test: Copy template and do one global replacement...');

  const templateFile = findTemplateFile(SLIDES_CONFIG.TEMPLATE.FILE_NAME);
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
