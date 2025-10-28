# Node.js/Docker Document Generator (Legacy)

This directory contains the original Node.js-based document generators for creating DOCX and PPTX files from Excel word databases.

**Note:** This is the legacy approach. For new usage, see the Apps Script approach in `../apps-script/` which works directly in Google Drive.

## Key Dependencies

- `xlsx` - Reading Excel spreadsheet databases containing word lists
- `pptxgenjs` - Generating PowerPoint presentations with word slides
- `docxtemplater` with `pizzip` - Filling DOCX templates with word data
- `seedrandom` - Deterministic shuffling of word order using seeded RNG

## Core Scripts

### 1. `gen-docx-template.js` - Generate DOCX word lists

Generates Word documents with word lists including definitions and example sentences.

**Features:**
- Reads Excel file sheets matching pattern `Round \d+`
- Filters words by year column (e.g., "2024", "2019Fall")
- Shuffles words using seeded random (`seedrandom('hello.')`)
- Applies custom XML formatting to bold words within example sentences
- Fills DOCX template with word, pronunciation, definition, and sentence

**Usage:**
```bash
node gen-docx-template.js <year> <xlsx_file> <template_file> <output_dir>
```

### 2. `make-ppts.js` - Generate PowerPoint slides

Generates PowerPoint presentations with BEF branding.

**Features:**
- Creates custom BEF-branded slide masters with bee graphics and gradients
- Generates title slide, "Spellers at Work" interstitial slides, and word slides
- Uses Questrial font and specific BEF color scheme (green: #94C601, etc.)

**Usage:**
```bash
node make-ppts.js <year> <xlsx_file> <output_dir>
```

## Running with Docker

The `../commands` file contains the full workflow. Typical process:

```bash
# 1. Copy Excel database to word_db/
cp "$HOME/Downloads/spelling-bee-word-database-2024.xlsx" ../word_db/

# 2. Build Docker image
docker build -t pstoll/bef-spelling-bee-pdf-gen ..

# 3. Run container with mounted volumes
docker run -it \
  -v "${PWD}/../word_db":/word_db \
  -v "${PWD}/../inputs":/inputs \
  -v "${PWD}/../outputs":/outputs \
  -v "${PWD}":/src \
  pstoll/bef-spelling-bee-pdf-gen

# 4. Inside container, generate outputs
export year=2024
export word_db="spelling-bee-word-database-2024.xlsx"
export doctemplate="bee-template-2024.docx"
node ./gen-docx-template.js "${year}" "/word_db/${word_db}" "/inputs/${doctemplate}" /outputs/2024
node make-ppts.js "${year}" "/word_db/${word_db}" /outputs/2024
```

## Directory Structure

- **`../word_db/`** - Excel word databases (gitignored)
  - Excel files contain multiple sheets, one per round (e.g., "Round 1", "Round 2")
  - Each row has columns: Word, Pronunciation, Definition, Sentence, and year columns (2018, 2019, 2019Fall, 2022, 2023, 2024, etc.)

- **`../inputs/`** - DOCX templates and graphics (committed to git)
  - DOCX templates use docxtemplater syntax for mail-merge-style generation
  - Bee graphics used in presentations (`bee-background.png`, `bee-left.png`, `bee-right.png`)

- **`../outputs/`** - Generated files (gitignored)
  - `bee-words-<year>-round<N>.docx` - Word lists with definitions
  - `bee-slides-<year>-round<N>.pptx` - Presentation slides

## Post-Processing

After generation, DOCX files are manually converted to PDF for delivery to the BEF team. This is done by opening the DOCX files in Word or Pages and saving as PDF.

## Implementation Details

- **Deterministic shuffling**: Both scripts use `seedrandom('hello.')` to shuffle words in the same order every time
- **Sheet name parsing**: Sheets must match regex `/Round\s+(\d+)(\s|\+)?$/` to be processed
- **Year column filtering**: Records are included only if the year column is non-empty (not null, not whitespace)
- **BEF branding constants**: Colors and fonts defined in `befDefaults` object in make-ppts.js
- **Bold sentence formatting**: gen-docx-template.js manually injects Word XML to bold the target word within example sentences
