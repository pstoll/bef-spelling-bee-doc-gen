# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This repository contains tools to generate spelling bee materials (word lists and presentation slides) for the Brookline Education Foundation's 5th Grade Spelling Bee.

**Two approaches available:**

1. **Apps Script (Recommended)** - Generate Google Docs and Slides directly from Google Sheets
   - No local setup required
   - Works directly in the cloud
   - Anyone with Sheet access can run it
   - See `apps-script/` directory and README

2. **Node.js/Docker (Legacy)** - Generate DOCX/PPTX from downloaded Excel files
   - Requires local setup and Docker
   - Outputs DOCX/PPTX files that need PDF conversion
   - See tag `nodejs-docker-only` for the original version
   - Still maintained in `src/` directory

---

# Apps Script Approach (Recommended)

See detailed documentation in `apps-script/README.md`

## Quick Start

1. Open the BEF Spelling Bee Google Sheet
2. Extensions → Apps Script
3. Copy code from `apps-script/Code.gs`, `Slides.gs`, and `Docs.gs`
4. Reload spreadsheet
5. Use "BEF Spelling Bee" menu → "Generate Materials..."

## Architecture

- **Code.gs** - Menu, data extraction, coordination
- **Slides.gs** - Google Slides generation with BEF branding
- **Docs.gs** - Google Docs generation with formatted word lists

All configuration is in `CONFIG` objects at the top of each file. No hardcoded values.

---

# Node.js/Docker Approach (Legacy)

## Key Dependencies

- `xlsx` - Reading Excel spreadsheet databases containing word lists
- `pptxgenjs` - Generating PowerPoint presentations with word slides
- `docxtemplater` with `pizzip` - Filling DOCX templates with word data
- `seedrandom` - Deterministic shuffling of word order using seeded RNG

## Core Architecture

The repository contains two main generation scripts:

1. **`src/gen-docx-template.js`** - Generates DOCX word lists with definitions
   - Reads Excel file sheets matching pattern `Round \d+`
   - Filters words by year column (e.g., "2024", "2019Fall")
   - Shuffles words using seeded random (`seedrandom('hello.')`)
   - Applies custom XML formatting to bold words within example sentences
   - Fills DOCX template with word, pronunciation, definition, and sentence
   - Usage: `node gen-docx-template.js <year> <xlsx_file> <template_file> <output_dir>`

2. **`src/make-ppts.js`** - Generates PowerPoint slides
   - Creates custom BEF-branded slide masters with bee graphics and gradients
   - Generates title slide, "Spellers at Work" interstitial slides, and word slides
   - Uses Questrial font and specific BEF color scheme (green: #94C601, etc.)
   - Usage: `node make-ppts.js <year> <xlsx_file> <output_dir>`

Both scripts share the same core logic:
- Parse Excel file to extract sheets named like "Round 1", "Round 2", etc.
- Filter records where the specified year column is non-empty
- Shuffle using `seedrandom('hello.')` for reproducible order
- Generate one output file per round

## Running with Docker

The `commands` file contains the full workflow. The typical process:

```bash
# 1. Copy Excel database to word_db/
cp "$HOME/Downloads/spelling-bee-word-database-2024-20241113-2000.xlsx" word_db/

# 2. Build Docker image
docker build -t pstoll/bef-spelling-bee-pdf-gen .

# 3. Run container with mounted volumes
docker run -it \
  -v "${PWD}/word_db":/word_db \
  -v "${PWD}/inputs":/inputs \
  -v "${PWD}/outputs":/outputs \
  -v "${PWD}/src":/src \
  pstoll/bef-spelling-bee-pdf-gen

# 4. Inside container, generate outputs
export year=2024
export word_db="spelling-bee-word-database-2024-20241113-2000.xlsx"
export doctemplate="bee-template-2024.docx"
node ./gen-docx-template.js "${year}" "/word_db/${word_db}" "/inputs/${doctemplate}" /outputs/2024
node make-ppts.js "${year}" "/word_db/${word_db}" /outputs/2024
```

## Input Files Structure

- **`word_db/`** - Excel word databases (gitignored)
  - Excel files contain multiple sheets, one per round (e.g., "Round 1", "Round 2")
  - Each row has columns: Word, Pronunciation, Definition, Sentence, and year columns (2018, 2019, 2019Fall, 2022, 2023, 2024, etc.)

- **`inputs/`** - DOCX templates and graphics (committed to git)
  - DOCX templates use docxtemplater syntax for mail-merge-style generation
  - Bee graphics used in presentations (`bee-background.png`, `bee-left.png`, `bee-right.png`)

## Output Structure

Generated files are saved to `outputs/<year>/`:
- `bee-words-<year>-round<N>.docx` - Word lists with definitions
- `bee-slides-<year>-round<N>.pptx` - Presentation slides

## Manual Post-Processing

After generation, DOCX files are manually converted to PDF for delivery to the BEF team. This is done by opening the DOCX files in Word or Pages and saving as PDF.

## Important Implementation Details

- **Deterministic shuffling**: Both scripts use `seedrandom('hello.')` to shuffle words in the same order every time
- **Sheet name parsing**: Sheets must match regex `/Round\s+(\d+)(\s|\+)?$/` to be processed
- **Year column filtering**: Records are included only if the year column is non-empty (not null, not whitespace)
- **BEF branding constants**: Colors and fonts defined in `befDefaults` object in make-ppts.js
- **Bold sentence formatting**: gen-docx-template.js manually injects Word XML to bold the target word within example sentences
