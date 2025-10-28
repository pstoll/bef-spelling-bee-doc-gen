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

- **Code.js** - Main script with menu, data extraction, and coordination logic
- **Slides.js** - Google Slides generation functions
- **Docs.js** - Google Docs generation functions
- **appsscript.json** - Apps Script project configuration
- **.clasp.json** - Configuration for clasp CLI deployment
- **.claspignore** - Files to exclude from deployment

## For End Users (Running the Script)

**If the script is already installed in your spreadsheet, you just need to authorize it once.**

### First Time Authorization (Required Once Per User)

1. **Open the BEF Spelling Bee Word Database spreadsheet** in Google Sheets

2. **Look for the "BEF Spelling Bee" menu** at the top
   - It appears between "Help" and your profile picture
   - If you don't see it, wait a few seconds for the sheet to fully load
   - If it still doesn't appear, contact the administrator

3. **Click any menu item** (e.g., "BEF Spelling Bee → Generate Materials...")

4. **You'll see an "Authorization Required" dialog**
   - Click **"Continue"** or **"Review Permissions"**

5. **Select your Google account**
   - Choose the account that has access to this spreadsheet

6. **You'll see a warning: "Google hasn't verified this app"**
   - **This is normal!** Google shows this for all custom scripts
   - Click **"Advanced"** (small text at bottom left)
   - Click **"Go to BEF Spelling Bee Generator (unsafe)"**
   - Despite "unsafe", this is safe - it's a custom script for BEF, not a public app

7. **Review the permissions and click "Allow"**
   - The script needs permission to read the spreadsheet and create Slides/Docs

8. **You're done!** The authorization is complete
   - You won't need to do this again unless permissions change
   - The menu function you clicked should now run

### What If I Get "Permission Denied"?

If you get permission errors:
1. Try running the menu item again - sometimes the first attempt fails
2. Make sure you're logged into the correct Google account
3. Close the spreadsheet tab and reopen it
4. If still failing, contact the administrator

## For Developers (Deploying Code Updates)

### Prerequisites

- Node.js and npm installed
- Access to the Apps Script project
- Clasp CLI installed (`npm install -g @google/clasp`)

### Template Files Setup

The Apps Script requires two template files to be present in the same Google Drive folder as your spreadsheet:

1. **`bef-bee-slide-template`** - Google Slides template
2. **`bef-bee-words-template`** - Google Docs template

**Template backups are stored in this repo** at `apps-script/templates/` as PPTX and DOCX files.

#### Uploading Templates to Google Drive

If you need to set up the templates (e.g., in a new Google account or if they were deleted):

1. **Navigate to the folder** containing your BEF Spelling Bee spreadsheet in Google Drive

2. **Upload the Slides template:**
   - Upload `apps-script/templates/bef-bee-slide-template.pptx` to the folder
   - Google Drive will automatically convert it to Google Slides format
   - Rename it to **exactly** `bef-bee-slide-template` (case-insensitive, but must match)

3. **Upload the Docs template:**
   - Upload `apps-script/templates/bef-bee-words-template.docx` to the folder
   - Google Drive will automatically convert it to Google Docs format
   - Rename it to **exactly** `bef-bee-words-template` (case-insensitive, but must match)

4. **Verify the templates are in the correct location:**
   - They should be in the same folder as your spreadsheet
   - The script will search for them by name in this folder

#### Updating Template Backups

If you modify the templates in Google Drive and want to update the repo backups:

1. **Download from Google Drive:**
   - Open the template in Google Slides/Docs
   - File → Download → Microsoft PowerPoint (.pptx) or Microsoft Word (.docx)

2. **Replace the backup files:**
   - Save to `apps-script/templates/` (overwriting the old files)
   - Commit to git with a description of what changed

### Initial Setup

1. **Clone this repository**
   ```bash
   git clone https://github.com/pstoll/bef-spelling-bee-doc-gen.git
   cd bef-spelling-bee-doc-gen
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Login to clasp** (if not already logged in)
   ```bash
   npm run clasp:login
   ```
   - This opens a browser to authorize clasp
   - Use the Google account that owns the Apps Script project

4. **Verify connection**
   ```bash
   npm run clasp:push
   ```
   - Should show "Script is already up to date" or push the files

### Making Code Changes

1. **Edit the code locally** in your editor (VS Code, etc.)
   - Files are in `apps-script/Code.js`, `Slides.js`, `Docs.js`

2. **Push changes to Apps Script**
   ```bash
   npm run clasp:push
   ```
   - This uploads your local changes to the Apps Script project
   - Users will see the changes immediately (may need to reload the spreadsheet)

3. **Test the changes**
   - Open the spreadsheet in your browser
   - Reload the page to pick up the new code
   - Run the menu items to test

4. **Commit to git**
   ```bash
   git add apps-script/
   git commit -m "Description of changes"
   git push
   ```

### Useful Commands

```bash
# Push local code to Apps Script
npm run clasp:push

# Pull latest code from Apps Script (if you edited in the web editor)
npm run clasp:pull

# Open the Apps Script project in browser
npm run clasp:open
```

## Usage

### Generate Materials for a Spelling Bee

1. **Open the BEF Spelling Bee Word Database spreadsheet**

2. **Click the menu: BEF Spelling Bee → Generate Materials...**

3. **Enter the year** when prompted
   - Examples: `2024`, `2019Fall`
   - This must match a column name in your spreadsheet exactly

4. **Enter which rounds to generate**
   - Type `1` to generate just Round 1
   - Type `1,2,3` to generate Rounds 1, 2, and 3
   - Type `all` to generate all rounds (may timeout for many rounds)
   - **Tip**: Generate one round at a time to avoid timeouts

5. **Wait for generation to complete**
   - The script will process each round (this may take 30-60 seconds per round)
   - You'll see a progress message while it runs

6. **Check the results dialog**
   - Shows links to all generated files
   - Files are saved in a Google Drive folder named "BEF Spelling Bee [YEAR]"
   - The folder is created in the same directory as your spreadsheet

### What Gets Generated

For each round, the script creates:
- **Slides** - Google Slides presentation with one slide per word
- **Docs** - Google Docs word list with definitions and example sentences

### Requirements

- A slide template named **"bef-bee-slide-template"** must exist in the same folder as your spreadsheet
- A doc template named **"bef-bee-words-template"** must exist in the same folder as your spreadsheet
- Spreadsheet sheets must be named **"Round 1", "Round 2",** etc.
- The year column (e.g., "2024") must exist with "X" or "x" marks for words to include

### Troubleshooting

**"Exceeded maximum execution time"**
- Generate rounds one at a time instead of all at once
- Try generating just "1" instead of "all"

**"No files generated"**
- Check that the year you entered matches a column name exactly
- Use "Debug: Show Sheet Structure" menu item to see available year columns

**"Template file not found"**
- Make sure "bef-bee-slide-template" exists in the same folder as your spreadsheet
- Check that the name matches exactly (case-insensitive)

### Debug Tools

Available in the menu for troubleshooting:

- **Debug: Show Sheet Structure** - Shows column names and sample data from Round 1

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

## Status

**Both Docs and Slides generation are fully working and tested!**

- ✅ Tested end-to-end with all 6 rounds (2024)
- ✅ Generates 12 files (6 slides + 6 docs) successfully
- ✅ Docs use template table approach to keep word entries together on same page
- ✅ Slides use template-based approach with BEF branding
- ✅ Handles duplicate file names (deletes old files before creating new ones)
- ✅ No execution timeouts with full 6-round generation

## TODO

### High Priority
- [ ] **Create detailed video walkthrough** for non-technical users showing:
  - How to authorize the script
  - How to run generation
  - What to do if errors occur

### Medium Priority
- [ ] **Better error messages** - Make errors more user-friendly for non-technical users
- [ ] **Add "Setup Check" menu item** - Verify template exists, folder structure is correct, etc.
- [ ] **Let user select the output folder** - Allow choosing destination instead of auto-creating in same directory
- [ ] **Make event date configurable** - Allow users to set it in a cell instead of auto-generating
- [ ] **Add progress indicator** - Show progress when generating multiple rounds
- [ ] **Optimize waitForFileReady timeout** - Reduce from 30s to 5s (currently exits early anyway)

### Low Priority
- [ ] **Add PDF export functionality** - Auto-convert Slides/Docs to PDF
- [ ] **Add email notification** when generation completes
- [ ] **Improve template validation** - Better error messages if template is missing placeholders
- [ ] **Add "Preview" feature** - Generate just first 3 words to preview formatting

## Completed

- ✅ Allow selecting specific rounds instead of all rounds
- ✅ Create output folder in same directory as spreadsheet
- ✅ Test and verify Doc generation - working with all 68 words in Round 1
- ✅ Test and verify Slides generation - working with template-based approach
- ✅ Fix Docs formatting using pristine template row copy
- ✅ Add file deletion to prevent duplicate file names

## Potential Future Enhancements

- Enhanced BEF branding in slides (currently template-based)
- Automatic year detection from current date
- Word count statistics and reports
