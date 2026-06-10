# PDF Converter

Simple CLI tool to convert `.docx` and `.pptx` to PDF using Microsoft Office.

## Features

- Convert single file or entire folder
- Batch processing
- Supports DOCX and PPTX

## Requirements

- Windows
- Microsoft Word & PowerPoint installed
- Python 3.8+

## Install

```bash
pip install .
```

## Usage

### 1. Convert single file
pdfconverter file.docx

### 2. Convert multiple files
pdfconverter a.docx b.pptx

### 3. Convert folder
pdfconverter ./documents

## Notes
- Output PDF is saved next to original file.
- Existing PDFs will be overwritten.