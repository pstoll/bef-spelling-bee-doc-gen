# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This repository contains tools to generate spelling bee materials (word lists and presentation slides) for the Brookline Education Foundation's 5th Grade Spelling Bee.

**Two approaches available:**

1. **Apps Script (Recommended)** - Generate Google Docs and Slides directly from Google Sheets
   - No local setup required
   - Works directly in the cloud
   - Anyone with Sheet access can run it
   - See `apps-script/README.md` for details

2. **Node.js/Docker** - Generate DOCX/PPTX from downloaded Excel files
   - Requires local setup and Docker
   - Outputs DOCX/PPTX files that need PDF conversion
   - See `nodejs-generators/README.md` for details

---

# Apps Script Approach (Recommended)

See detailed documentation in `apps-script/README.md`

## Quick Start

1. Open the BEF Spelling Bee Google Sheet
2. Extensions → Apps Script
3. Copy code from `apps-script/Code.js`, `Slides.js`, and `Docs.js`
4. Reload spreadsheet
5. Use "BEF Spelling Bee" menu → "Generate Materials..."

## Architecture

- **Code.js** - Menu, data extraction, coordination
- **Slides.js** - Google Slides generation with BEF branding
- **Docs.js** - Google Docs generation with formatted word lists

All configuration is in `CONFIG` objects at the top of each file. No hardcoded values.
