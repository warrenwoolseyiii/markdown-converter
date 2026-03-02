# Markdown ↔ Word Document Converter

Python scripts for bidirectional conversion between Markdown and Microsoft Word (.docx) format, enabling seamless round-trip document editing.

## Features

### Markdown to Word (`md_to_docx.py`)

Converts Markdown files to Word format with support for:
- Headings with automatic numbering (1, 1.1, 1.1.1, etc.)
- Tables with formatted headers
- Code blocks (preserves ASCII art diagrams)
- Numbered and bullet lists
- Bold and italic text formatting
- Hyperlinks
- Horizontal rules
- YAML front matter for document metadata
- Title page generation
- Table of contents

### Word to Markdown (`docx_to_md.py`)

Converts Word documents to Markdown format with support for:
- Heading extraction (from Word styles)
- Tables to markdown table syntax
- Code block detection (via monospace font heuristics)
- Numbered and bullet lists with nesting
- Bold, italic, and inline code formatting
- Hyperlink extraction
- Image extraction to separate directory
- Horizontal rules

## Installation

### Requirements
- Python 3.8 or higher
- pip (Python package installer)

### Setup

1. Navigate to the scripts directory:
   ```bash
   cd scripts
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Markdown to Word

**Basic conversion:**
```bash
python md_to_docx.py input.md output.docx
```

**With template for styling:**
```bash
python md_to_docx.py input.md output.docx --template template.docx
```

### Word to Markdown

**Basic conversion:**
```bash
python docx_to_md.py input.docx output.md
```

**With custom images directory:**
```bash
python docx_to_md.py input.docx output.md --images-dir assets/images
```

### Round-Trip Workflow

The typical workflow for document collaboration:

1. **Author in Markdown** - Write and maintain your document in Markdown
2. **Convert to Word** - Generate a Word document for distribution
   ```bash
   python md_to_docx.py mydoc.md mydoc.docx --template company_template.docx
   ```
3. **Distribute for Review** - Send the Word document to reviewers
4. **Receive Edits** - Get the edited Word document back
5. **Convert Back to Markdown** - Extract changes to Markdown
   ```bash
   python docx_to_md.py mydoc_edited.docx mydoc_updated.md
   ```
6. **Review and Merge** - Compare and merge changes into your source Markdown

## Command-Line Options

### `md_to_docx.py`

| Option | Description |
|--------|-------------|
| `input` | Path to the input Markdown file (required) |
| `output` | Path for the output Word document (required) |
| `--template`, `-t` | Path to a template .docx file for style reference (optional) |

### `docx_to_md.py`

| Option | Description |
|--------|-------------|
| `input` | Path to the input Word document (required) |
| `output` | Path for the output Markdown file (required) |
| `--images-dir`, `-i` | Directory for extracted images, relative to output file (default: `images`) |

## Supported Elements

### Headings
```markdown
# Heading 1
## Heading 2
### Heading 3
```

Word styles mapping:
- `Title` style → `# Heading 1`
- `Heading 1` style → `## Heading 2`
- `Heading 2` style → `### Heading 3`

### Tables
```markdown
| Column 1 | Column 2 | Column 3 |
|----------|----------|----------|
| Data 1   | Data 2   | Data 3   |
```

### Code Blocks
````markdown
```
ASCII art or code content here
Preserves line breaks and monospace formatting
```
````

Code blocks are detected by:
- Triple backtick fences (in Markdown)
- Monospace fonts like Consolas, Courier New (in Word)

### Lists

**Numbered:**
```markdown
1. First item
2. Second item
   1. Nested item
3. Third item
```

**Bullet:**
```markdown
- Item one
- Item two
  - Nested item
- Item three
```

### Inline Formatting
```markdown
**bold text**
*italic text*
`code text`
[link text](https://example.com)
```

### Images
```markdown
![image description](images/image_001.png)
```

Images are extracted from Word documents and saved to the specified images directory.

### Horizontal Rules
```markdown
---
```

## Output Formatting

### Markdown to Word Default Styles

| Element | Font | Size | Additional |
|---------|------|------|------------|
| Heading 1 | Arial | 16pt | Bold |
| Heading 2 | Arial | 14pt | Bold |
| Heading 3 | Arial | 12pt | Bold |
| Body text | Arial | 11pt | Normal |
| Code block | Consolas | 9pt | Monospace |
| Table header | Arial | 10pt | Bold, gray background |
| Table cell | Arial | 10pt | Normal |

## Template Support

When providing a `--template` option to `md_to_docx.py`, the script will:

- Use the template's heading styles (Heading 1, 2, 3)
- Apply the Normal style for body text
- Preserve headers, footers, and page setup from the template

If style extraction fails, the script falls back to default styles.

## Troubleshooting

### "Module not found" error
Make sure you've installed the dependencies:
```bash
pip install -r requirements.txt
```

**Note:** Use `python` (not `python3`) when running in a virtual environment on Windows.

### Tables not displaying correctly
Ensure your markdown tables have the separator row:
```markdown
| Header 1 | Header 2 |
|----------|----------|  <-- This line is required
| Data 1   | Data 2   |
```

### Code blocks not detected in Word
The converter detects code by looking for monospace fonts. Ensure code blocks in Word use:
- Consolas
- Courier New
- Other monospace fonts

### Images not extracted
Images must be embedded in the Word document (not linked). The converter extracts:
- PNG, JPEG, GIF, BMP images
- EMF and WMF vector graphics

## Development

### Project Structure
```
markdown-converter/
├── scripts/
│   ├── md_to_docx.py        # Markdown → Word converter
│   ├── docx_to_md.py        # Word → Markdown converter
│   ├── inspect_styles.py    # Utility to inspect Word styles
│   ├── requirements.txt     # Python dependencies
│   └── README.md            # This file
├── input/                   # Input documents
├── output/                  # Generated documents
│   └── images/              # Extracted images
└── template/                # Style templates
```

### Class Architecture

#### `md_to_docx.py`
1. **MarkdownParser**: Parses markdown into structured elements
2. **MarkdownToDocxConverter**: Converts elements to Word format

#### `docx_to_md.py`
1. **DocxParser**: Extracts elements from Word documents
2. **MarkdownGenerator**: Produces markdown from parsed elements
3. **DocxToMarkdownConverter**: Orchestrates the conversion

### Extending the Converters

**To add new element types:**

1. Add parsing logic in the respective parser class
2. Add generation logic in the converter/generator class
3. Update the element type handlers

**Example: Adding support for a new element**
```python
# In DocxParser
def _parse_new_element(self, element):
    # Extract content
    return ParsedElement(type='new_type', content=content)

# In MarkdownGenerator
def _new_element_to_md(self, element):
    return f"<!-- {element.content} -->"
```
