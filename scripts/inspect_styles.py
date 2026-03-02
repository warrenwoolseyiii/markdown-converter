#!/usr/bin/env python3
"""Inspect styles in a Word document."""

from docx import Document
import sys

def main():
    doc_path = sys.argv[1] if len(sys.argv) > 1 else 'example_format.docx'
    doc = Document(doc_path)
    
    print(f"=== Styles in {doc_path} ===\n")
    
    for style in doc.styles:
        try:
            style_type = {
                1: 'PARAGRAPH',
                2: 'CHARACTER', 
                3: 'TABLE',
                4: 'LIST'
            }.get(style.type, f'UNKNOWN({style.type})')
            print(f"  {style.name:40} | Type: {style_type}")
        except Exception as e:
            print(f"  Error: {e}")
    
    print("\n=== Done ===")

if __name__ == '__main__':
    main()
