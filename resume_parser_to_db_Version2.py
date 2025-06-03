import os
import sys
import glob
import sqlite3
import traceback
from typing import List, Dict

try:
    from docx import Document
except ImportError:
    print("Missing 'python-docx'. Install with: pip install python-docx")
    sys.exit(1)

try:
    import docx2txt
except ImportError:
    print("Missing 'docx2txt'. Install with: pip install docx2txt")
    sys.exit(1)

try:
    import PyPDF2
except ImportError:
    print("Missing 'PyPDF2'. Install with: pip install PyPDF2")
    sys.exit(1)

def extract_text_from_docx(file_path: str) -> str:
    try:
        doc = Document(file_path)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception:
        # Fallback to docx2txt
        return docx2txt.process(file_path)

def extract_text_from_pdf(file_path: str) -> str:
    text = ""
    with open(file_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text += page.extract_text() or ""
    return text

def extract_text_from_doc(file_path: str) -> str:
    """
    Attempts to extract text from .doc files by converting to .docx using LibreOffice (if available).
    Requires LibreOffice CLI (soffice).
    """
    try:
        import subprocess
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            output_docx = os.path.join(tmpdir, "converted.docx")
            cmd = [
                "soffice",
                "--headless",
                "--convert-to",
                "docx",
                "--outdir",
                tmpdir,
                file_path
            ]
            subprocess.run(cmd, capture_output=True, check=True)
            if os.path.exists(output_docx):
                return extract_text_from_docx(output_docx)
            else:
                raise Exception("Conversion to docx failed.")
    except Exception as e:
        print(f"Failed to extract .doc file '{file_path}': {e}")
        return ""

def scan_resumes(folder: str) -> List[Dict[str, str]]:
    resumes_data = []
    # Collect all pdf, doc, docx files (case-insensitive)
    file_patterns = ["*.docx", "*.DOCX", "*.doc", "*.DOC", "*.pdf", "*.PDF"]
    files = []
    for pat in file_patterns:
        files.extend(glob.glob(os.path.join(folder, pat)))
    for file_path in files:
        ext = file_path.lower().split(".")[-1]
        try:
            if ext == "docx":
                text = extract_text_from_docx(file_path)
            elif ext == "doc":
                text = extract_text_from_doc(file_path)
            elif ext == "pdf":
                text = extract_text_from_pdf(file_path)
            else:
                print(f"Unsupported file type: {file_path}")
                continue
            resumes_data.append({
                "filename": os.path.basename(file_path),
                "filepath": file_path,
                "content": text.strip()
            })
        except Exception as e:
            print(f"Failed to extract {file_path}: {e}")
            traceback.print_exc()
    return resumes_data

def init_db(db_file: str):
    conn = sqlite3.connect(db_file)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS resumes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            filepath TEXT,
            content TEXT
        )
    """)
    conn.commit()
    conn.close()

def insert_resume(db_file: str, resume: Dict[str, str]):
    conn = sqlite3.connect(db_file)
    c = conn.cursor()
    c.execute("""
        INSERT INTO resumes (filename, filepath, content)
        VALUES (?, ?, ?)
    """, (resume['filename'], resume['filepath'], resume['content']))
    conn.commit()
    conn.close()

def build_master_database(resume_folder: str, db_file: str):
    print(f"Initializing database: {db_file}")
    init_db(db_file)
    print(f"Scanning resumes in folder: {resume_folder}")
    resumes = scan_resumes(resume_folder)
    print(f"Found {len(resumes)} resumes.")
    for resume in resumes:
        insert_resume(db_file, resume)
    print("Master database created/updated successfully.")

if __name__ == "__main__":
    resume_folder = '/Users/phobrla/Documents/Career/Application Materials'
    db_file = "master_resumes.db"
    if len(sys.argv) > 1:
        resume_folder = sys.argv[1]
    if len(sys.argv) > 2:
        db_file = sys.argv[2]
    if not os.path.isdir(resume_folder):
        print(f"Resume folder does not exist: {resume_folder}")
        sys.exit(1)
    build_master_database(resume_folder, db_file)
    print(f"Database '{db_file}' ready! You can query the resumes now.")