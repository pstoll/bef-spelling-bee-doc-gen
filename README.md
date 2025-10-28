# BEF Spelling Bee Document Generator

Tools to generate spelling bee materials (word lists and presentation slides) for the Brookline Education Foundation's 5th Grade Spelling Bee.

## Two Approaches

### 1. Apps Script (Recommended)

Generate Google Docs and Slides directly from Google Sheets. No local setup required.

**Features:**
- Works entirely in the cloud
- Anyone with spreadsheet access can run it
- Generates Google Docs word lists and Google Slides presentations
- No file downloads/uploads needed

**Documentation:** See [apps-script/README.md](apps-script/README.md)

**Quick Start:**
1. Open the BEF Spelling Bee Google Sheet
2. Use the "BEF Spelling Bee" menu → "Generate Materials..."
3. Enter year and rounds to generate
4. Files are created in Google Drive

---

### 2. Node.js/Docker (Legacy)

Generate DOCX and PPTX files from downloaded Excel files. Requires local setup.

**Features:**
- Generates Microsoft Word and PowerPoint files
- Deterministic shuffling with seedrandom
- Custom BEF branding and formatting

**Documentation:** See [CLAUDE.md](CLAUDE.md) for details

**When to use:**
- Need DOCX/PPTX files instead of Google formats
- Working offline
- Need exact reproducibility with seeded shuffling

---

## Repository Structure

```
├── apps-script/          # Google Apps Script code (recommended)
│   ├── Code.js          # Main script and menu
│   ├── Slides.js        # Google Slides generation
│   ├── Docs.js          # Google Docs generation
│   └── templates/       # Template file backups
├── src/                 # Node.js generators (legacy)
├── inputs/              # Templates and graphics
└── word_db/            # Word database files (gitignored)
```

## Getting Started

**For end users:** Use the Apps Script approach - it's already set up in your Google Sheet.

**For developers:** See [apps-script/README.md](apps-script/README.md) for deployment instructions.
