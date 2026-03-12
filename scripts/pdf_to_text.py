#!/usr/bin/env python3
"""
PDF to Text Extractor with OCR Support

Extracts text from PDF documents using PyMuPDF (fitz) with OCR fallback
via pytesseract for image-based or scanned PDFs.

Usage:
    python pdf_to_text.py input.pdf output.txt
    python pdf_to_text.py input.pdf output.txt --ocr-only
"""

import argparse
import sys
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    print("Error: PyMuPDF not installed. Run: pip install pymupdf")
    sys.exit(1)

try:
    import pytesseract
    from PIL import Image
    HAS_TESSERACT = True
except ImportError:
    HAS_TESSERACT = False
    print("Warning: pytesseract not installed. OCR fallback disabled.")


def configure_tesseract():
    """Configure Tesseract path for Windows."""
    import platform
    if platform.system() == "Windows":
        # Common installation paths
        tesseract_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]
        for path in tesseract_paths:
            if Path(path).exists():
                pytesseract.pytesseract.tesseract_cmd = path
                return True
    return True  # Assume it's in PATH for non-Windows


def extract_text_from_page(page, use_ocr=False):
    """
    Extract text from a single PDF page.
    
    Args:
        page: PyMuPDF page object
        use_ocr: Force OCR even if text layer exists
        
    Returns:
        Extracted text as string
    """
    # Try to extract text layer first
    if not use_ocr:
        text = page.get_text("text")
        if text.strip():
            return text
    
    # Fall back to OCR if no text or OCR forced
    if HAS_TESSERACT:
        # Render page to image at 300 DPI for better OCR
        mat = fitz.Matrix(300/72, 300/72)  # 300 DPI
        pix = page.get_pixmap(matrix=mat)
        
        # Convert to PIL Image
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # Run OCR
        text = pytesseract.image_to_string(img, lang='eng')
        return text
    else:
        return ""


def extract_pdf_text(pdf_path, output_path=None, use_ocr=False, verbose=True):
    """
    Extract all text from a PDF document.
    
    Args:
        pdf_path: Path to input PDF file
        output_path: Optional path to save extracted text
        use_ocr: Force OCR for all pages
        verbose: Print progress information
        
    Returns:
        Extracted text as string
    """
    pdf_path = Path(pdf_path)
    
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")
    
    if verbose:
        print(f"Processing: {pdf_path}")
    
    # Configure Tesseract if available
    if HAS_TESSERACT:
        configure_tesseract()
    
    # Open PDF
    doc = fitz.open(pdf_path)
    
    all_text = []
    total_pages = len(doc)
    
    for page_num in range(total_pages):
        if verbose:
            print(f"  Processing page {page_num + 1}/{total_pages}...", end=" ")
        
        page = doc.load_page(page_num)
        text = extract_text_from_page(page, use_ocr=use_ocr)
        
        # Add page marker
        page_header = f"\n{'='*60}\n--- PAGE {page_num + 1} ---\n{'='*60}\n\n"
        all_text.append(page_header + text)
        
        if verbose:
            word_count = len(text.split())
            print(f"({word_count} words)")
    
    doc.close()
    
    # Combine all text
    full_text = "\n".join(all_text)
    
    # Save to file if output path provided
    if output_path:
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(full_text, encoding='utf-8')
        if verbose:
            print(f"\nSaved extracted text to: {output_path}")
    
    return full_text


def main():
    parser = argparse.ArgumentParser(
        description="Extract text from PDF documents using PyMuPDF with OCR fallback"
    )
    parser.add_argument(
        "input",
        help="Path to input PDF file"
    )
    parser.add_argument(
        "output",
        nargs="?",
        help="Path for output text file (optional, prints to stdout if not provided)"
    )
    parser.add_argument(
        "--ocr-only", "-o",
        action="store_true",
        help="Force OCR for all pages (ignore existing text layer)"
    )
    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Suppress progress output"
    )
    
    args = parser.parse_args()
    
    try:
        text = extract_pdf_text(
            args.input,
            args.output,
            use_ocr=args.ocr_only,
            verbose=not args.quiet
        )
        
        # Print to stdout if no output file specified
        if not args.output:
            print("\n" + "="*60)
            print("EXTRACTED TEXT:")
            print("="*60)
            print(text)
            
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error processing PDF: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
