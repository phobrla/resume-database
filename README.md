# Resume Database

This repository provides a Python script and supporting tools to scan, parse, and maintain a master SQLite database of resume files in `.docx`, `.doc`, and `.pdf` formats.

## Features

- **Bulk scanning:** Recursively scan a directory for resume files.
- **Parsing:** Extracts text from DOCX, DOC (via LibreOffice), and PDF files.
- **Database:** Stores extracted content, filename, and path in a SQLite database for easy querying.
- **Easy setup:** Includes instructions and requirements for macOS and Linux.

## Quick Start

1. Place your resumes in a folder (e.g. `resumes/`).
2. Install required Python packages:
   ```
   pip install python-docx docx2txt PyPDF2
   ```
   For `.doc` support, install LibreOffice via Homebrew:
   ```
   brew install --cask libreoffice
   ```
3. Run the script:
   ```
   python resume_parser_to_db.py /path/to/resumes
   ```

## License

MIT License.
