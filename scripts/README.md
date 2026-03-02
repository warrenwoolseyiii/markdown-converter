# Markdown to Word Document Converter

A Python script that converts Markdown files to Microsoft Word (.docx) format, with support for:
- Headings with automatic numbering (1, 1.1, 1.1.1, etc.)
- Tables with formatted headers
- Code blocks (preserves ASCII art diagrams)
- Numbered and bullet lists
- Bold and italic text formatting
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

### Basic Usage

Convert a markdown file to Word:

```bash
python md_to_docx.py input.md output.docx
```

### With Template Reference

Use a template document to match styles:

```bash
python md_to_docx.py input.md output.docx --template example_format.docx
```

### Examples

Convert a markdown document to a word document.

```bash
# From the project root directory
python scripts/md_to_docx.py path/to/mydocument.md output/mydocument.docx

# With template
python scripts/md_to_docx.py path/to/mydocument.md output/mydocument.docx --template example_format.docx
```

## Command-Line Options

| Option | Description |
|--------|-------------|
| `input` | Path to the input Markdown file (required) |
| `output` | Path for the output Word document (required) |
| `--template`, `-t` | Path to a template .docx file for style reference (optional) |

## Supported Markdown Elements

### Headings
```markdown
# Heading 1
## Heading 2
### Heading 3
```

Output: Numbered headings (1, 1.1, 1.1.1)

### Tables
```markdown
| Column 1 | Column 2 | Column 3 |
|----------|----------|----------|
| Data 1   | Data 2   | Data 3   |
```

Output: Formatted table with header row styling

### Code Blocks
````markdown
```
ASCII art or code content here
Preserves line breaks and monospace formatting
```
````

Output: Consolas font, preserved whitespace

### Lists

Numbered:
```markdown
1. First item
2. Second item
3. Third item
```

Bullet:
```markdown
- Item one
- Item two
- Item three
```

### Inline Formatting
```markdown
**bold text**
*italic text*
`code text`
```

### Horizontal Rules
```markdown
---
```

## Output Formatting

The converter applies the following default styles:

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

When providing a `--template` option, the script will attempt to extract styles from the template document and apply them to the output. This includes:

- Font names and sizes from Heading 1, 2, 3 styles
- Font settings from the Normal style

If style extraction fails, the script falls back to default styles.

## Troubleshooting

### "Module not found" error
Make sure you've installed the dependencies:
```bash
pip install -r requirements.txt
```

### Tables not displaying correctly
Ensure your markdown tables have the separator row:
```markdown
| Header 1 | Header 2 |
|----------|----------|  <-- This line is required
| Data 1   | Data 2   |
```

### Code blocks appear inline
Make sure you're using triple backticks (```) for code blocks, not single backticks.

### ASCII art alignment issues
The script uses Consolas font for code blocks. For best results with ASCII art:
- Use consistent character widths
- Avoid very wide diagrams (may exceed page width)

## Development

### Project Structure
```
markdown-converter/
├── example_format.docx      # Style reference template
├── scripts/
│   ├── md_to_docx.py        # Main converter script
│   ├── requirements.txt     # Python dependencies
│   └── README.md            # This file
└── output/                  # Generated documents
```

### Extending the Converter

The converter is built with three main classes:

1. **MarkdownParser**: Parses markdown into structured elements
2. **DocxStyleManager**: Manages document styles
3. **MarkdownToDocxConverter**: Orchestrates the conversion

To add support for new markdown elements, extend the `MarkdownParser.parse()` method and add a corresponding handler in `MarkdownToDocxConverter._process_element()`.
