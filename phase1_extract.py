"""
PHASE 1: PDF to Raw CSV Extraction (FIXED VERSION)
===================================================

Fixes:
1. Removed invalid gpu parameter from EasyOCR
2. Better OCR configuration to avoid € symbol misinterpretation
3. Post-processing to fix common OCR errors
4. Character substitution for known OCR mistakes
"""
# phase1_extract.py
"""
PHASE 1: PDF to Raw CSV with CONTEXT-AWARE OCR correction
==========================================================

Only applies OCR corrections in appropriate contexts:
- USN codes: O→0, I→1, B→8, etc.
- Names/Text: Keep letters as letters
- Marks/Numbers: Convert OCR errors in numbers
"""

import os
import sys
import re
import pandas as pd
from img2table.document import PDF
from img2table.ocr import EasyOCR
import warnings
warnings.filterwarnings('ignore')


def is_likely_usn(text):
    """Check if text looks like a USN (1BM21CS001 pattern)"""
    if not isinstance(text, str):
        return False
    
    text = text.upper().strip()
    
    # USN pattern: 1BM + 7 digits/letters
    usn_pattern = r'^1[BM][A-Z0-9]{7,9}$'
    
    # Clean common OCR errors first for checking
    temp = text.replace('O', '0').replace('I', '1').replace('B', '8')
    
    return bool(re.match(usn_pattern, temp))


def is_likely_marks(text):
    """Check if text looks like marks (numbers, AB, etc.)"""
    if not isinstance(text, str):
        return False
    
    text = text.strip().upper()
    
    # Marks can be: numbers, "AB", "F", "P", etc.
    marks_patterns = [
        r'^\d+$',           # Just numbers
        r'^[A-FP]$',        # Single letter grade
        r'^AB$',            # Absent
        r'^\d+\s*\d*$',     # Numbers with spaces
    ]
    
    return any(re.match(pattern, text) for pattern in marks_patterns)


def correct_text_context_aware(text):
    """
    Apply OCR corrections based on context.
    
    Rules:
    1. For USNs: O→0, I→1, B→8, etc.
    2. For marks/numbers: Fix OCR errors in numbers
    3. For text/names: Only fix definitely wrong characters (€→C)
    """
    if not isinstance(text, str):
        return text
    
    original = text
    text = text.strip()
    
    if not text:
        return ""
    
    # ALWAYS fix these (definitely wrong in any context)
    always_fixes = {
        '€': 'C',      # Euro symbol → C
        'Ⓒ': 'C',     # Copyright C → C
        '©': 'C',      # Copyright → C
        '‘': "'",      # Smart quotes
        '’': "'",
        '“': '"',
        '”': '"',
    }
    
    for wrong, right in always_fixes.items():
        text = text.replace(wrong, right)
    
    # Check context and apply appropriate corrections
    if is_likely_usn(text):
        # USN-specific corrections
        usn_corrections = {
            'O': '0',  # Letter O → Zero
            'I': '1',  # Letter I → One
            'l': '1',  # Lowercase L → One
            'B': '8',  # Letter B → Eight
            'Z': '2',  # Letter Z → Two
            'S': '5',  # Letter S → Five
        }
        for wrong, right in usn_corrections.items():
            text = text.replace(wrong, right)
        
        # Special: Fix 1BM2ICS157 → 1BM21CS157
        text = re.sub(r'1BM2(I|1)CS', '1BM21CS', text)
        
        # Remove any non-alphanumeric
        text = re.sub(r'[^A-Z0-9]', '', text)
        
        return text.upper()
    
    elif is_likely_marks(text):
        # Marks/number-specific corrections
        marks_corrections = {
            'O': '0',  # O → 0 in numbers
            'I': '1',  # I → 1 in numbers
            'l': '1',  # l → 1 in numbers
            'B': '8',  # B → 8 in numbers
            'S': '5',  # S → 5 in numbers
        }
        for wrong, right in marks_corrections.items():
            text = text.replace(wrong, right)
        
        # Clean up marks
        text = re.sub(r'\s+', '', text)  # Remove spaces in numbers
        
        return text
    
    else:
        # Text/name context - be conservative!
        # Only fix obviously wrong things, keep letters as letters
        
        # Remove non-ASCII but keep basic punctuation
        text = re.sub(r'[^\x20-\x7E]', ' ', text)
        
        # Clean extra spaces
        text = re.sub(r'\s+', ' ', text)
        
        return text.strip()


def extract_pdf_to_csv_context_aware(pdf_path, output_csv="debug_raw_table_context.csv"):
    """
    Extract with context-aware OCR correction.
    """
    
    print("\n" + "="*60)
    print("PHASE 1: PDF TO CSV (CONTEXT-AWARE OCR)")
    print("="*60)
    print("Applies corrections based on context:")
    print("  • USNs: O→0, I→1, B→8")
    print("  • Marks: Fix number OCR errors")
    print("  • Names/Text: Keep letters as letters")
    print("="*60)
    
    if not os.path.exists(pdf_path):
        print(f"❌ ERROR: PDF not found: {pdf_path}")
        return False
    
    print(f"📄 Input: {pdf_path}")
    print(f"📊 Output: {output_csv}")
    
    try:
        print("🔧 Initializing OCR...")
        ocr = EasyOCR(lang=['en'])
        
        print("📄 Loading PDF...")
        pdf = PDF(pdf_path, detect_rotation=True)
        
        print("🔍 Extracting tables...")
        tables = pdf.extract_tables(
            ocr=ocr,
            borderless_tables=True,
            implicit_rows=True,
            min_confidence=30
        )
        
        page_tables = tables.get(0, [])
        
        if not page_tables:
            print("❌ No tables found")
            return False
        
        print(f"✅ Found {len(page_tables)} table(s)")
        
        if len(page_tables) >= 2:
            main_table = page_tables[1]
            print("📋 Using table 2 (main marks table)")
        else:
            main_table = page_tables[0]
            print("📋 Using table 1 (only table found)")
        
        raw_df = main_table.df
        print(f"📈 Dimensions: {raw_df.shape[0]} rows × {raw_df.shape[1]} columns")
        
        # Apply CONTEXT-AWARE corrections
        print("\n🔧 Applying CONTEXT-AWARE OCR corrections...")
        
        # Process column by column for better context understanding
        for col_idx in range(raw_df.shape[1]):
            column_data = raw_df.iloc[:, col_idx]
            
            # Try to infer column type from first few rows
            sample_cells = column_data.head(5).dropna().astype(str).tolist()
            
            # Apply context-aware correction to each cell
            corrected_column = column_data.apply(
                lambda x: correct_text_context_aware(x) if pd.notna(x) else x
            )
            
            raw_df.iloc[:, col_idx] = corrected_column
        
        # Save
        print(f"\n💾 Saving to {output_csv}...")
        raw_df.to_csv(output_csv, index=False, header=False, encoding='utf-8-sig')
        
        if os.path.exists(output_csv):
            size_kb = os.path.getsize(output_csv) / 1024
            print(f"✅ File saved ({size_kb:.1f} KB)")
            
            # Show preview
            print("\n🔍 PREVIEW (first 3 rows):")
            with open(output_csv, 'r', encoding='utf-8') as f:
                for i, line in enumerate(f):
                    if i < 3:
                        preview = line[:100].strip()
                        if len(line) > 100:
                            preview += "..."
                        print(f"Row {i}: {preview}")
                    else:
                        break
            
            return True
        else:
            print("❌ Output file not created")
            return False
            
    except Exception as e:
        print(f"❌ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


def main():
    if len(sys.argv) < 2:
        print("Usage: python phase1_extract.py input.pdf [output.csv]")
        print("\nContext-aware OCR correction:")
        print("  • USNs: O→0, I→1, B→8 (correctly)")
        print("  • Names: Keep O, I, B as letters")
        print("  • Marks: Fix number OCR errors")
        return 1
    
    pdf_path = sys.argv[1]
    output_csv = sys.argv[2] if len(sys.argv) > 2 else "debug_raw_table_context.csv"
    
    success = extract_pdf_to_csv_context_aware(pdf_path, output_csv)
    
    if success:
        print("\n" + "="*60)
        print("✅ CONTEXT-AWARE EXTRACTION COMPLETE!")
        print("="*60)
        print(f"\nNext: Run Phase 2:")
        print(f"  python phase2_process.py {output_csv}")
        return 0
    else:
        print("\n❌ EXTRACTION FAILED")
        return 1


if __name__ == "__main__":
    print("\n" + "="*60)
    print("PDF → CSV (CONTEXT-AWARE OCR CORRECTION)")
    print("="*60)
    exit_code = main()
    sys.exit(exit_code)