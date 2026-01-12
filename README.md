
# PDF → Excel Converter (Context-Aware OCR)

This project extracts tabular data from scanned PDF documents (image-based) using **EasyOCR** and **img2table**, applies **context-aware corrections** (USNs, marks, names), and outputs clean structured data in both **CSV** and **Excel** formats.

---

## 📋 Prerequisites

1. **Python version**
   - Must be **Python 3.11**
   - Check your version:
     ```bash
     python --version
     ```
   - If not 3.11, install Python 3.11 before proceeding.

2. **Sample files included**
   - `phase1_extract.py` → Extracts raw table from PDF with OCR corrections
   - `phase2_process.py` → Processes CSV into structured Excel workbook
   - `requirements.txt` → Pinned dependencies for reproducible environment
   - `sample-doc-from-megha-1.pdf` → Example input PDF

---

## ⚙️ Setup Instructions

1. **Create and activate virtual environment**
   ```bash
   python -m venv venv
   ```
   - On Git Bash / MINGW64:
     ```bash
     source venv/Scripts/activate
     ```
   - On Command Prompt (cmd):
     ```cmd
     venv\Scripts\activate
     ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

---

## 🚀 Usage

### Phase 1 – Extract raw CSV from PDF
Run the OCR extraction script:
```bash
python phase1_extract.py sample-input-doc-received.pdf
```

- Output: `debug_raw_table_context.csv`
- Context-aware corrections applied:
  - USNs: `O → 0`, `I → 1`, `B → 8`
  - Marks: OCR number fixes
  - Names/Text: Keep letters intact

### Phase 2 – Process CSV into Excel
Run the processor script:
```bash
python phase2_process.py debug_raw_table_context.csv
```

- Outputs:
  - `perfect_student_marks.csv` → Clean structured data
  - `perfect_student_marks.xlsx` → Formatted Excel workbook with multiple sheets:
    - **Summary**
    - **Student_Marks**
    - **Marks_Aliases**

---

## ✅ Verification

- Check that USNs are correctly normalized (e.g., `1BM21CS006` not `1BM21CSOO6`).
- Ensure no stray OCR artifacts (e.g., `€` symbols).
- Confirm student names and marks are aligned across subjects.

---

## 📂 Folder Structure

```
TEMP-PDF-IMG-TO-EXCEL/
├── debug_raw_table_context.csv
├── perfect_student_marks.csv
├── perfect_student_marks.xlsx
├── phase1_extract.py
├── phase2_process.py
├── requirements.txt
├── sample-input-doc-received.pdf
└── venv/
```

---

## 📝 Notes

- Always run inside the **venv** to avoid dependency conflicts.
- Share `requirements.txt` with students for reproducible installs.
- Tested and verified on **Python 3.11** with pinned package versions.
 

--- 
