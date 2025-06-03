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

# Explicit list of files to ignore (absolute, normalized paths),
# but do NOT include files containing 'PhilHobrlaCL' or 'Cover Letter' (case-insensitive) in their name.
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
    '/Users/phobrla/Documents/Career/Application Materials/PhilHobrlaCL (1).docx',  # Covered by skip logic, can be removed
    '/Users/phobrla/Documents/Career/Application Materials/Carilion Clinic/Help Desk Specialist I - May 2019/PhilHobrlaCL (2).docx',  # Covered by skip logic, can be removed
    '/Users/phobrla/Documents/Career/Application Materials/PSRRPT-190322111419-42911.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/Resume for Phil.docx',
    '/Users/phobrla/Documents/Career/Application Materials/Resume_10203452191975339.docx',
    '/Users/phobrla/Documents/Career/Application Materials/Robertson Marketing Group/ResumePhilHobrlaRobertsonMarketingGroupProjectManager100.docx',
    '/Users/phobrla/Documents/Career/Application Materials/ResumePhilHobrlaVenveoProjectManager98.docx',
    '/Users/phobrla/Documents/Career/Application Materials/Robertson Marketing Group/Robertson Interview Prep.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/Workforce Recruitment Program/Schedule A Letter.pdf',
    '/Users/phobrla/Documents/Career/Application Materials/Workforce Recruitment Program/Schedule A Lettter II.pdf',
}

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

def normalize_path(path: str) -> str:
    return os.path.abspath(os.path.normpath(path))

def should_ignore(file_path: str, ignore_files_norm: set) -> bool:
    """
    Returns True if file should be skipped based on:
      - Explicit ignore file list (normalized)
      - Filename contains 'PhilHobrlaCL' or 'Cover Letter' (case-insensitive)
    """
    abs_path = normalize_path(file_path)
    basename = os.path.basename(file_path).lower()
    if abs_path in ignore_files_norm:
        return True
    if 'philhobrlacl' in basename or 'cover letter' in basename:
        return True
    return False

def scan_resumes(folder: str) -> List[Dict[str, Any]]:
    resumes_data = []
    # Recursively collect all pdf, doc, docx files (case-insensitive)
    file_patterns = ["**/*.docx", "**/*.DOCX", "**/*.doc", "**/*.DOC", "**/*.pdf", "**/*.PDF"]
    files = []
    for pat in file_patterns:
        files.extend(glob.glob(os.path.join(folder, pat), recursive=True))
    # Normalize ignore file paths for comparison
    ignore_files_norm = {normalize_path(f) for f in IGNORE_FILES}
    for file_path in files:
        if should_ignore(file_path, ignore_files_norm):
            print(f"Skipping ignored file: {file_path}")
            continue
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
            parsed = parse_resume(text)
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

def parse_resume(text: str) -> Dict[str, Any]:
    """
    Parse header, summary at top, each employer (summary+highlights), and skills with headers.
    This uses simple heuristics for common resume layouts. For best results, use consistently-formatted resumes.
    """
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    # Header: Assume at the very top (name and contact, usually 1-4 lines)
    header = []
    summary = ""
    employers = []
    skills = []

    # Identify header (first block before a blank line or a known section)
    header_lines = []
    for idx, line in enumerate(lines):
        if (
            re.search(r'^(summary|objective|profile|professional|experience|employment|work|skills)', line, re.I)
            or line == ""
        ):
            break
        header_lines.append(line)
    header = "\n".join(header_lines)

    # Find summary: lines after header and before first work/employment section
    summary_lines = []
    i = len(header_lines)
    while i < len(lines):
        line = lines[i]
        if re.search(r'^(experience|employment|work\s+history|professional\s+experience|career\s+history)', line, re.I):
            break
        if re.search(r'^(skills|technical\s+skills|core\s+competencies)', line, re.I):
            break
        summary_lines.append(line)
        i += 1
    summary = " ".join(summary_lines).strip()

    # Find employers: look for experience sections and parse summaries/highlights
    employers = []
    employer_section = False
    employer = {}
    highlights = []
    in_employer = False
    for j in range(i, len(lines)):
        line = lines[j]
        # Start of experience section
        if re.search(r'^(experience|employment|work\s+history|professional\s+experience|career\s+history)', line, re.I):
            employer_section = True
            continue
        # Start of skills section = end of employer parsing
        if re.search(r'^(skills|technical\s+skills|core\s+competencies)', line, re.I):
            if employer:
                employer["highlights"] = highlights
                employers.append(employer)
            break

        # Heuristic: employer header lines often contain company name and role, or are ALL CAPS or Title Case
        if employer_section and (re.match(r'^[A-Z][A-Za-z0-9&,\.\-\s]+$', line) and not line.endswith(":")):
            if employer:
                employer["highlights"] = highlights
                employers.append(employer)
                highlights = []
            employer = {"header": line, "summary": "", "highlights": []}
            in_employer = True
            # Next line(s) may be summary or job title/dates
            continue
        if in_employer and ('•' in line or re.match(r'^[-*•]\s+', line)):
            # Highlight/bullet
            highlights.append(line)
        elif in_employer and not employer.get("summary"):
            # First line after employer header that is not a bullet is likely the summary
            employer["summary"] = line
        elif in_employer and employer.get("summary"):
            # More summary lines until first bullet
            if not ('•' in line or re.match(r'^[-*•]\s+', line)):
                employer["summary"] += " " + line
            else:
                highlights.append(line)
    if employer and employer not in employers:
        employer["highlights"] = highlights
        employers.append(employer)

    # Find skills and headers
    skills = []
    skill_section_re = re.compile(r'^(skills|technical\s+skills|core\s+competencies)', re.I)
    current_header = None
    skill_lines = []
    found_skill_section = False
    for idx, line in enumerate(lines):
        if skill_section_re.match(line):
            found_skill_section = True
            current_header = line
            continue
        if found_skill_section:
            # Section ends with a new section or empty line
            if re.match(r'^(experience|employment|work\s+history|professional\s+experience|career\s+history|education|certification)', line, re.I):
                if current_header and skill_lines:
                    skills.append({"header": current_header, "skills": skill_lines})
                break
            # Header within skills
            if line and not re.search(r'[-•*]', line) and len(line.split()) < 7:
                if current_header and skill_lines:
                    skills.append({"header": current_header, "skills": skill_lines})
                current_header = line
                skill_lines = []
            else:
                skill_lines.append(line)
    if found_skill_section and current_header and skill_lines:
        skills.append({"header": current_header, "skills": skill_lines})

    return {
        "header": header,
        "summary": summary,
        "employers": employers,
        "skills": skills
    }

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

    # Insert employers
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
    # Insert skills
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
    print(f"Initializing database: {db_file}")
    init_db(db_file)
    print(f"Recursively scanning resumes in folder: {resume_folder}")
    resumes = scan_resumes(resume_folder)
    print(f"Found {len(resumes)} resumes.")
    for resume in resumes:
        insert_resume(db_file, resume)
    print("Master database created/updated successfully.")

if __name__ == "__main__":
    # Default folder and database file
    resume_folder = '/Users/phobrla/Documents/Career/Application Materials'
    db_file = os.path.join(resume_folder, "master_resumes.db")
    if len(sys.argv) > 1:
        resume_folder = sys.argv[1]
        db_file = os.path.join(resume_folder, "master_resumes.db")
    if len(sys.argv) > 2:
        db_file = sys.argv[2]
    if not os.path.isdir(resume_folder):
        print(f"Resume folder does not exist: {resume_folder}")
        sys.exit(1)
    build_master_database(resume_folder, db_file)
    print(f"Database '{db_file}' ready! You can query the resumes now.")
