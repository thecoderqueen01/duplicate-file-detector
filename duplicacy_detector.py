import os
import hashlib
import csv

import fitz  # PyMuPDF
from docx import Document
import pandas as pd


# folder where input files are present
INPUT_DIR = "input_files"

# csv file for duplicate report
OUTPUT_CSV = "duplicate_report.csv"


def readTextFile(path):
    try:
        f = open(path, "r", encoding="utf-8", errors="ignore")
        data = f.read()
        f.close()
        return data
    except:
        return ""


def readPdfFile(path):
    text = ""
    try:
        doc = fitz.open(path)
        for page in doc:
            text += page.get_text()
        doc.close()
    except:
        pass
    return text


def readDocxFile(path):
    text = ""
    try:
        doc = Document(path)
        for p in doc.paragraphs:
            text += p.text
    except:
        pass
    return text


def readExcelFile(path):
    text = ""
    try:
        sheets = pd.read_excel(path, sheet_name=None)
        for sheet in sheets.values():
            text += sheet.to_string()
    except:
        pass
    return text


def getFileContent(path):
    ext = os.path.splitext(path)[1].lower()

    if ext == ".txt":
        return readTextFile(path)

    if ext == ".pdf":
        return readPdfFile(path)

    if ext == ".docx":
        return readDocxFile(path)

    if ext == ".xls" or ext == ".xlsx":
        return readExcelFile(path)

    return ""


def md5FromText(text):
    return hashlib.md5(
        text.encode("utf-8", errors="ignore")
    ).hexdigest()


def md5FromFile(path):
    h = hashlib.md5()
    try:
        f = open(path, "rb")
        while True:
            chunk = f.read(8192)
            if not chunk:
                break
            h.update(chunk)
        f.close()
    except:
        pass
    return h.hexdigest()


def findDuplicates():
    hashMap = {}

    for root, _, files in os.walk(INPUT_DIR):
        for file in files:
            fullPath = os.path.join(root, file)

            content = getFileContent(fullPath)

            if content.strip() != "":
                fileHash = md5FromText(content)
            else:
                fileHash = md5FromFile(fullPath)

            if fileHash not in hashMap:
                hashMap[fileHash] = []

            hashMap[fileHash].append(fullPath)

    return hashMap


def writeCsv(result):
    out = open(OUTPUT_CSV, "w", newline="", encoding="utf-8")
    writer = csv.writer(out)
    writer.writerow(["group_id", "file_path"])

    gid = 1
    for h in result:
        files = result[h]
        if len(files) > 1:
            for f in files:
                writer.writerow([gid, f])
            gid += 1

    out.close()


def main():
    print("checking files...")
    result = findDuplicates()
    writeCsv(result)
    print("done, output saved in", OUTPUT_CSV)


main()
