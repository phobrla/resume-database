import os
import sys
import glob
import sqlite3
import traceback
import re
from typing import List, Dict, Any

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

IGNORE_FILES = {
    '/Users/phobrla/Documents/Career/Application Materials/vollman----interview-notes-unstructured.docx',
    '/Users/phobrla/Documents/Career/Application Materials/Project Manager- IT - Salem, VA 24153 - Indeed.com.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/Interview Questions and Answers.docx',
    '/Users/phobrla/Documents/Career/Application Materials/Results.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/TopResume JobScan Results.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/project-management-prepcast.com-PMP Formulas and Calculations - The Complete Guide.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/_Master Files/zzz - Old Templates/business-cl--phil-hobrla1.docx',
    '/Users/phobrla/Documents/Career/Application Materials/Moog/Business Analyst IT Project Manager - April 2019/Business Analyst_IT Project Manager - Moog, Inc. Careers.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/_Master Files/zzz - Old Templates/business-cl--phil-hobrla.docx',
    '/Users/phobrla/Documents/Career/Application Materials/Transcon Environmental/Project Coordinator - April 2019/Cabell_Foundation_Grant_Proposal_2017___Writing_Sample.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/_Master Files/zzz - Old Templates/CLForJobScan_2019-03-20T10-13-09.docx',
    '/Users/phobrla/Documents/Career/Application Materials/_Master Files/zzz - Old Templates/Cover Letter Template.dotx',
    '/Users/phobrla/Documents/Career/Application Materials/TMEIC/Epson_03282019104428.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/hobrla-cover-letter.docx',
    '/Users/phobrla/Documents/Career/Application Materials/American Family RV/Project Coordinator - March 2019/indeed.com-Project Coordinator.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/Carilion Clinic/Application Analyst I/JobScan - Report 1.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/Moog/Business Analyst IT Project Manager - April 2019/moog.careers-Business AnalystIT Project Manager.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/Norfolk Southern/NS CL - Phil Hobrla.docx',
    '/Users/phobrla/Documents/Career/Application Materials/Phil Hobrla - LinkedIn Recommendations as of April 1 2019.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/Phil Hobrla - References (1).docx',
    '/Users/phobrla/Documents/Career/Application Materials/Phil Hobrla - References.docx',
    '/Users/phobrla/Documents/Career/Application Materials/Workforce Recruitment Program/Phil Vollman - Recommendations.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/PSRRPT-190322111419-42911.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/Resume for Phil.docx',
    '/Users/phobrla/Documents/Career/Application Materials/Resume_10203452191975339.docx',
    '/Users/phobrla/Documents/Career/Application Materials/Robertson Marketing Group/ResumePhilHobrlaRobertsonMarketingGroupProjectManager100.docx',
    '/Users/phobrla/Documents/Career/Application Materials/ResumePhilHobrlaVenveoProjectManager98.docx',
    '/Users/phobrla/Documents/Career/Application Materials/Robertson Marketing Group/Robertson Interview Prep.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/Workforce Recruitment Program/Schedule A Letter.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/Workforce Recruitment Program/Schedule A Lettter II.pdf',
}

def normalize_path(path: str) -> str:
    return os.path.abspath(os.path.normpath(path))

def should_ignore(file_path: str, ignore_files_norm: set) -> bool:
    abs_path = normalize_path(file_path)
    basename = os.path.basename(file_path).lower()
    if abs_path in ignore_files_norm:
        return True
    if 'philhobrlacl' in basename or 'cover letter' in basename:
        return True
    return False

def extract_text_from_docx(file_path: str) -> str:
    try:
        doc = Document(file_path)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception:
        return docx2txt.process(file_path)

def extract_text_from_pdf(file_path: str) -> str:
    text = ""
    with open(file_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text += page.extract_text() or ""
    return text

def extract_text_from_doc(file_path: str) -> str:
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

def parse_resume(text: str) -> Dict[str, Any]:
    # Robust section parsing using regex and section titles
    sections = {}
    current_section = "header"
    sections[current_section] = []
    lines = [line.strip() for line in text.splitlines()]
    section_titles = {
        "summary": re.compile(r"^(summary|profile|objective)$", re.I),
        "skills": re.compile(r"^(skills|technical\s+skills|core\s+competencies)$", re.I),
        "experience": re.compile(r"^(experience|work\s+history|professional\s+experience|employment|career\s+history)$", re.I),
        "education": re.compile(r"^education$", re.I),
    }
    current_section_found = False

    for line in lines:
        line_clean = line.strip()
        found_section = None
        for name, pat in section_titles.items():
            if pat.match(line_clean):
                found_section = name
                break
        if found_section:
            current_section = found_section
            if current_section not in sections:
                sections[current_section] = []
            current_section_found = True
            continue
        if not current_section_found and not line_clean:
            # First blank line after header
            current_section_found = True
            continue
        if current_section_found:
            sections.setdefault(current_section, []).append(line_clean)
        else:
            sections[current_section].append(line_clean)

    # Compose outputs
    header = "\n".join([l for l in sections.get("header", []) if l])
    summary = " ".join([l for l in sections.get("summary", []) if l])
    # Parse experiences
    employers = []
    if "experience" in sections:
        exp_lines = sections["experience"]
        emp = None
        for line in exp_lines:
            if not line:
                continue
            # Heuristic: employer header likely in Title Case or ALL CAPS, not a bullet
            if not line.startswith(("-", "*", "•")) and (line.istitle() or line.isupper()):
                if emp:
                    employers.append(emp)
                emp = {"header": line, "summary": "", "highlights": []}
            elif emp and (line.startswith("•") or line.startswith("-") or line.startswith("*")):
                emp["highlights"].append(line)
            elif emp:
                if emp["summary"]:
                    emp["summary"] += " " + line
                else:
                    emp["summary"] = line
        if emp:
            employers.append(emp)
    skills = []
    if "skills" in sections:
        skill_lines = [l for l in sections["skills"] if l]
        skills = [{"header": "Skills", "skills": skill_lines}]
    return {
        "header": header,
        "summary": summary,
        "employers": employers,
        "skills": skills
    }

def scan_resumes(folder: str) -> List[Dict[str, Any]]:
    resumes_data = []
    file_patterns = ["**/*.docx", "**/*.DOCX", "**/*.doc", "**/*.DOC", "**/*.pdf", "**/*.PDF"]
    files = []
    for pat in file_patterns:
        files.extend(glob.glob(os.path.join(folder, pat), recursive=True))
    ignore_files_norm = {normalize_path(f) for f in IGNORE_FILES}
    for file_path in files:
        if should_ignore(file_path, ignore_files_norm):
            print(f"Skipping ignored file: {file_path}")
            continue
        ext = file_path.lower().split(".")[-1]
        print(f"Parsing: {file_path}")
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
            if not text.strip():
                print(f"WARNING: No text extracted from {file_path}")
            parsed = parse_resume(text)
            print(f"  Header: {parsed['header'][:40]}...")
            print(f"  Summary: {parsed['summary'][:40]}...")
            print(f"  Employers found: {len(parsed['employers'])}")
            print(f"  Skills found: {len(parsed['skills'])}")
            resumes_data.append({
                "filename": os.path.basename(file_path),
                "filepath": normalize_path(file_path),
                "content": text.strip(),
                "header": parsed["header"],
                "summary": parsed["summary"],
                "employers": parsed["employers"],
                "skills": parsed["skills"]
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
            content TEXT,
            header TEXT,
            summary TEXT
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS employers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            resume_id INTEGER,
            employer_header TEXT,
            employer_summary TEXT,
            highlights TEXT,
            FOREIGN KEY(resume_id) REFERENCES resumes(id)
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS skills (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            resume_id INTEGER,
            skill_header TEXT,
            skills TEXT,
            FOREIGN KEY(resume_id) REFERENCES resumes(id)
        )
    """)
    conn.commit()
    conn.close()

def insert_resume(db_file: str, resume: Dict[str, Any]):
    conn = sqlite3.connect(db_file)
    c = conn.cursor()
    c.execute("""
        INSERT INTO resumes (filename, filepath, content, header, summary)
        VALUES (?, ?, ?, ?, ?)
    """, (resume['filename'], resume['filepath'], resume['content'], resume['header'], resume['summary']))
    resume_id = c.lastrowid

    for emp in resume.get("employers", []):
        c.execute("""
            INSERT INTO employers (resume_id, employer_header, employer_summary, highlights)
            VALUES (?, ?, ?, ?)
        """, (
            resume_id,
            emp.get("header", ""),
            emp.get("summary", ""),
            "\n".join(emp.get("highlights", []))
        ))
    for s in resume.get("skills", []):
        c.execute("""
            INSERT INTO skills (resume_id, skill_header, skills)
            VALUES (?, ?, ?)
        """, (
            resume_id,
            s.get("header", ""),
            "\n".join(s.get("skills", []))
        ))
    conn.commit()
    conn.close()

def build_master_database(resume_folder: str, db_file: str):
    print(f"Will save DB to: {db_file}")
    try:
        # Check write permission to parent folder
        db_dir = os.path.dirname(db_file)
        if not os.access(db_dir, os.W_OK):
            print(f"ERROR: No write permission to folder {db_dir}")
            sys.exit(1)
        init_db(db_file)
        print(f"Recursively scanning resumes in folder: {resume_folder}")
        resumes = scan_resumes(resume_folder)
        print(f"Found {len(resumes)} resumes to insert.")
        for resume in resumes:
            insert_resume(db_file, resume)
        print("Master database created/updated successfully.")
    except Exception as e:
        print(f"ERROR: Could not create database at {db_file}: {e}")
        sys.exit(1)

if __name__ == "__main__":
    resume_folder = '/Users/phobrla/Documents/Career/Application Materials'
    db_file = os.path.join(resume_folder, "master_resumes.db")
    if len(sys.argv) > 1:
        resume_folder = sys.argv[1]
        db_file = os.path.join(resume_folder, "master_resumes.db")
    if len(sys.argv) > 2:
        db_file = sys.argv[2]
    print(f"Running from: {os.getcwd()}")
    print(f"Input folder: {resume_folder}")
    print(f"DB target: {db_file}")
    if not os.path.isdir(resume_folder):
        print(f"Resume folder does not exist: {resume_folder}")
        sys.exit(1)
    build_master_database(resume_folder, db_file)
    print(f"Database '{db_file}' ready! You can query the resumes now.")
