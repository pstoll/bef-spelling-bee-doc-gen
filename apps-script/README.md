# BEF Spelling Bee - Apps Script

This directory contains Google Apps Script code to generate spelling bee materials directly from the Google Sheet word database.

## Overview

The Apps Script:
- Reads word data directly from the Google Sheet (no download/upload needed)
- Filters words by year column (e.g., "2024", "2019Fall")
- Generates Google Slides presentations with BEF branding
- Generates Google Docs word lists with definitions and sentences
- Saves all files to a Google Drive folder
- Can be run by anyone with access to the spreadsheet

## Files

- **Code.gs** - Main script with menu, data extraction, and coordination logic
- **Slides.gs** - Google Slides generation functions
- **Docs.gs** - Google Docs generation functions

## Setup Instructions

### 1. Open Your Google Sheet

Open the BEF Spelling Bee Word Database spreadsheet in Google Sheets.

### 2. Open Apps Script Editor

1. Click **Extensions → Apps Script**
2. This opens the Apps Script editor in a new tab

### 3. Add the Script Files

In the Apps Script editor:

1. Delete any existing code in `Code.gs`
2. Copy the entire contents of `apps-script/Code.gs` and paste it
3. Click the `+` next to "Files" to add a new script file
4. Name it `Slides` (the .gs extension is added automatically)
5. Copy the entire contents of `apps-script/Slides.gs` and paste it
6. Add another new script file named `Docs`
7. Copy the entire contents of `apps-script/Docs.gs` and paste it

### 4. Save the Project

1. Click the disk icon or press `Cmd+S` / `Ctrl+S`
2. Name the project: "BEF Spelling Bee Generator"

### 5. Authorize the Script

1. Click **Run** (the play button) next to the function dropdown
2. Select `onOpen` from the dropdown if not already selected
3. Click **Run**
4. You'll see a permissions dialog - click **Review Permissions**
5. Select your Google account
6. Click **Advanced** → **Go to BEF Spelling Bee Generator (unsafe)**
7. Click **Allow**

This grants the script permission to:
- Read data from your spreadsheet
- Create Google Slides and Docs
- Create folders in Google Drive

### 6. Reload Your Spreadsheet

1. Go back to your Google Sheet tab
2. Refresh the page (Cmd+R / Ctrl+R)
3. You should now see a new menu: **BEF Spelling Bee**

## Usage

### Generate Materials

1. Click **BEF Spelling Bee → Generate Materials...**
2. Enter the year when prompted (e.g., `2024` or `2019Fall`)
3. The script will:
   - Find all sheets named "Round 1", "Round 2", etc.
   - Filter words for the specified year
   - Generate slides and docs for each round
   - Save everything to a Google Drive folder named "BEF Spelling Bee [YEAR]"
4. When complete, you'll see a dialog with links to the folder and files

### Test Functions

Before generating materials, you can test:

- **Test: Read Round Data...** - Reads and displays word count from a specific round
- **Test: Generate Sample Slide** - Creates a sample presentation to verify formatting

## Configuration

### Customizing Settings

Edit the `CONFIG` objects in each file to customize:

**Code.gs:**
- `CONFIG.COLUMNS` - Column names in your spreadsheet
- `CONFIG.SHEET_NAME_PATTERN` - Pattern for round sheet names
- `CONFIG.OUTPUT_FOLDER_PREFIX` - Folder name prefix

**Slides.gs:**
- `SLIDES_CONFIG.COLORS` - BEF brand colors
- `SLIDES_CONFIG.FONT` - Typography settings

**Docs.gs:**
- `DOCS_CONFIG.FONT` - Font and size settings
- `DOCS_CONFIG.SPACING` - Paragraph spacing

### Event Date

Currently the event date is generated automatically. To customize:
1. Edit the `getEventDate()` function in `Slides.gs`
2. Or store the date in a cell in your spreadsheet and read it

## Troubleshooting

### "No round sheets found"
- Make sure your sheets are named exactly "Round 1", "Round 2", etc.
- Check that there are no extra spaces or characters

### "No files generated"
- Verify the year column exists in your sheet
- Check that there are rows with non-empty values in the year column
- Use "Test: Read Round Data..." to debug

### "Permission denied" errors
- Re-run the authorization process (Step 5 above)
- Make sure you're logged into the correct Google account

### View Logs
1. In Apps Script editor: **View → Executions**
2. Click on a recent execution to see detailed logs
3. Errors will appear in red

## Development

### Testing Locally

To test changes without affecting production:
1. Make a copy of the spreadsheet
2. Attach the script to the copy
3. Test with a sample year

### Viewing Logs

Use `Logger.log()` in the code, then:
- Apps Script editor → **View → Executions**
- Or **View → Logs** (deprecated but still works)

### Debugging

1. Set breakpoints by clicking line numbers
2. Select a function from the dropdown
3. Click **Debug** (bug icon)
4. Step through code and inspect variables

## Differences from Node.js Version

### Shuffling Algorithm
The Apps Script version uses a simplified seeded random algorithm. It provides consistent shuffling but won't produce the exact same order as the Node.js version which uses the `seedrandom` library.

### BEF Branding
The Slides currently use basic text formatting. The Node.js version includes:
- Bee graphics (background, left, right images)
- Gradient backgrounds
- Custom slide masters

To add these to Apps Script, you would need to:
1. Upload bee graphics to Google Drive
2. Insert images using `slide.insertImage()`
3. Or create a template presentation and copy it

### Fonts
Google Slides may have different font availability than PowerPoint. The script uses Questrial (BEF brand font) but falls back to Arial if unavailable.

## Next Steps

Potential enhancements:
- Create a template Google Slides with BEF branding
- Add bee graphics to slides
- Make event date configurable via spreadsheet cell
- Add PDF export functionality
- Add email notification when generation completes
- Allow selecting specific rounds instead of all rounds
