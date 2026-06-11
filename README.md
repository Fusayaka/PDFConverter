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
```bash
pdfconverter file.docx
```

### 2. Convert multiple files
```bash
pdfconverter a.docx b.pptx
```

### 3. Convert folder
```bash
pdfconverter ./documents
```

### 4. Specify output directory
```bash
pdfconverter ./documents -o ./output
```

## Notes
- By default, each PDF is saved in the same directory as its source file.
- Use `-o` / `--output` to specify a custom output directory.
- Existing PDF files with the same name will be overwritten.