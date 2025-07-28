# Filename: merge_resumes.py
# Version: 1.2.0
#
# Merges two DOCX files at the XML level (unzipping, merging, and re-zipping) using only Python standard and pip-installable packages.
# No Homebrew/system dependencies required.
#
# Prerequisites:
# 1. If your environment is not yet activated:
#      cd ~/Documents && source myenv/bin/activate && pip install lxml
# 2. If your environment is already active:
#      pip install lxml
#
# Usage:
#     python merge_resumes.py
#
# Output:
#     Merged DOCX file at the path specified by MERGED_PATH

import os
import zipfile
import shutil
import tempfile
from lxml import etree

# --- Configurable File Paths (macOS/Unix style) ---
# Please verify these filenames match exactly as they appear in your Downloads folder
PART1_PATH = '/Users/phobrla/Library/Mobile Documents/com~apple~CloudDocs/Downloads/Phil Hobrla - Instructional Designer Resume Part 1.docx'
PART2_PATH = '/Users/phobrla/Library/Mobile Documents/com~apple~CloudDocs/Downloads/Phil Hobrla - Instructional Designer Resume Part 2.docx'
MERGED_PATH = '/Users/phobrla/Library/Mobile Documents/com~apple~CloudDocs/Downloads/Phil Hobrla - Instructional Designer Resume - Merged.docx'

def unzip_docx(docx_path, extract_to):
    """Unzip a .docx file to a given directory."""
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)

def rezip_docx(folder_path, output_path):
    """Zip a folder back into a .docx file."""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx_zip:
        for foldername, subfolders, filenames in os.walk(folder_path):
            for filename in filenames:
                abs_path = os.path.join(foldername, filename)
                arcname = os.path.relpath(abs_path, folder_path)
                docx_zip.write(abs_path, arcname)

def merge_document_xml(part1_xml, part2_xml, merged_xml):
    """
    Merge the bodies of document.xml files from two DOCX files.
    - Appends body elements from part2 to part1, preserves final sectPr from part2.
    """
    NSMAP = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    parser = etree.XMLParser(remove_blank_text=True)

    # Parse both document.xml files
    tree1 = etree.parse(part1_xml, parser)
    tree2 = etree.parse(part2_xml, parser)

    body1 = tree1.find('.//w:body', namespaces=NSMAP)
    body2 = tree2.find('.//w:body', namespaces=NSMAP)

    # Remove last sectPr from body1 (if present)
    sectPr1 = body1.find('w:sectPr', namespaces=NSMAP)
    if sectPr1 is not None:
        body1.remove(sectPr1)

    # Append all children from body2 except its sectPr
    for child in body2:
        if child.tag.endswith('sectPr'):
            continue
        body1.append(child)

    # Append section properties from part2 at the end
    sectPr2 = body2.find('w:sectPr', namespaces=NSMAP)
    if sectPr2 is not None:
        body1.append(sectPr2)

    # Write merged XML back
    tree1.write(merged_xml, xml_declaration=True, encoding='UTF-8', standalone='yes')

def main():
    # Ensure input files exist
    if not os.path.exists(PART1_PATH):
        print(f"Error: '{PART1_PATH}' not found.")
        exit(1)
    if not os.path.exists(PART2_PATH):
        print(f"Error: '{PART2_PATH}' not found.")
        exit(1)

    with tempfile.TemporaryDirectory() as tmpdir:
        part1_dir = os.path.join(tmpdir, 'part1')
        part2_dir = os.path.join(tmpdir, 'part2')
        merged_dir = os.path.join(tmpdir, 'merged')
        os.makedirs(part1_dir, exist_ok=True)
        os.makedirs(part2_dir, exist_ok=True)
        os.makedirs(merged_dir, exist_ok=True)

        # Unzip both DOCX files
        unzip_docx(PART1_PATH, part1_dir)
        unzip_docx(PART2_PATH, part2_dir)

        # Copy all contents of part1 to merged_dir
        for item in os.listdir(part1_dir):
            s = os.path.join(part1_dir, item)
            d = os.path.join(merged_dir, item)
            if os.path.isdir(s):
                shutil.copytree(s, d, dirs_exist_ok=True)
            else:
                shutil.copy2(s, d)

        # Merge word/document.xml files
        part1_xml = os.path.join(part1_dir, 'word', 'document.xml')
        part2_xml = os.path.join(part2_dir, 'word', 'document.xml')
        merged_xml = os.path.join(merged_dir, 'word', 'document.xml')
        merge_document_xml(part1_xml, part2_xml, merged_xml)

        # Rezip merged folder into new DOCX
        rezip_docx(merged_dir, MERGED_PATH)

    print(f"Merged DOCX created at: {MERGED_PATH}")

if __name__ == '__main__':
    main()