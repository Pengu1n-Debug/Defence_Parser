# DSGL Document Parser

A Python tool that directly parses Defence and Strategic Goods List (DSGL) documents and converts them to hierarchical JSON format - **no AI API required!**

## Overview

This parser replaces the previous AI-based approach (using OpenAI API) with direct document parsing logic. It reads DSGL documents in various formats (.docx, .pdf, .epub) and produces structured JSON output matching the `usml.json` format.

## Features

- **Direct Parsing**: No external AI API calls - pure Python parsing logic
- **Multiple Format Support**: Handles .docx, .pdf, and .epub files
- **Hierarchical Structure**: Properly nests categories, items, sub-items, and notes
- **DSGL-Specific**: Understands ML (Munitions List) category structure
- **Configurable**: Can extract just the munitions list or parse the entire document

## Installation

### Required Dependencies

```bash
pip install python-docx PyPDF2 ebooklib beautifulsoup4
```

## Usage

### Basic Usage

```python
from dsgl_parser_v2 import process_dsgl_document

# Parse DSGL document
result = process_dsgl_document(
    file_path="path/to/DSGL_document.docx",
    output_file="output.json",
    ml_only=True  # Extract munitions list only
)
```

### Command Line Usage

```bash
python dsgl_parser_v2.py
```

This will parse the default document specified in the script and save to `dsgl_munitions_list.json`.

### Customization

Edit the `__main__` section in `dsgl_parser_v2.py`:

```python
if __name__ == "__main__":
    # Change these paths as needed
    dsgl_file = r"c:\path\to\your\DSGL_document.docx"

    # Parse munitions list only
    result = process_dsgl_document(dsgl_file, "output.json", ml_only=True)

    # Or parse entire document
    result = process_dsgl_document(dsgl_file, "output_full.json", ml_only=False)
```

## Output Format

The parser generates JSON in the following hierarchical structure:

```json
[
  {
    "Label": "ML1",
    "Description": "Smooth-bore weapons with a calibre of less than 20 mm...",
    "SubStructures": [
      {
        "Label": "a.",
        "Description": "Firearms specially designed for...",
        "SubStructures": [
          {
            "Label": "1.",
            "Description": "Specific item description...",
            "SubStructures": []
          }
        ]
      },
      {
        "Label": "Note",
        "Description": "ML1. does not apply to...",
        "SubStructures": []
      }
    ]
  }
]
```

### Structure Elements

- **Label**: The identifier (e.g., "ML1", "a.", "1.", "Note", "Technical Note")
- **Description**: The text content describing the item
- **SubStructures**: Array of nested items following the same structure

### Hierarchy Levels

1. **Level 0**: ML Categories (ML1, ML2, etc.)
2. **Level 1**: Letter items (a., b., c., etc.)
3. **Level 2**: Number items (1., 2., 3., etc.)
4. **Level 3**: Nested letter items (a., b., c. under numbers)
5. **Special**: Notes can appear at any level

## Parser Architecture

### Key Components

1. **Document Loaders** (`load_docx_text`, `load_pdf_text`, `load_epub_text`)
   - Extract text from different document formats
   - Return list of lines for parsing

2. **DSGLParserV2 Class**
   - `classify_line()`: Determines what type each line is (category, item, note, etc.)
   - `build_hierarchy()`: Constructs the nested JSON structure
   - `extract_munitions_list()`: Finds and extracts ML section from document
   - `parse_document()`: Main entry point for parsing

3. **Pattern Recognition**
   - Regex patterns identify:
     - ML categories (ML1., ML2., etc.)
     - ML sub-items (ML1. a., ML1. b., etc.)
     - Letter items (a., b., c.)
     - Number items (1., 2., 3.)
     - Notes (Note:, Technical Note:, N.B.:)

## Comparison with Previous Approach

| Feature | AI-Based (old) | Direct Parser (new) |
|---------|---------------|---------------------|
| API Costs | Yes (OpenAI) | No |
| Internet Required | Yes | No |
| Speed | Slow (API calls) | Fast (local parsing) |
| Consistency | Variable | Deterministic |
| Accuracy | Depends on AI | Rule-based, predictable |
| Customization | Hard to control | Easy to modify rules |

## Files

- **`dsgl_parser_v2.py`**: Main parser implementation (recommended)
- **`dsgl_parser.py`**: Earlier version (kept for reference)
- **`DSGL_PROCESSOR.py`**: Original AI-based approach (deprecated)
- **`usml.json`**: Example output format from USML list
- **`dsgl_munitions_list.json`**: Parsed output from DSGL document

## Example Output Statistics

When parsing the F2024L01024.docx document:

- **Total ML Categories**: 30 (ML1 through approximately ML30)
- **Total Items**: 840+ across all categories
- **Format**: Matches usml.json hierarchical structure

## Extending the Parser

### Adding New Pattern Types

Edit the `patterns` dictionary in `__init__`:

```python
self.patterns = {
    'custom_pattern': re.compile(r'^PATTERN_HERE$'),
    # ... other patterns
}
```

### Modifying Classification Logic

Edit `classify_line()` method to add new classification rules:

```python
# Check for custom pattern
match = self.patterns['custom_pattern'].match(line)
if match:
    return {
        'type': 'custom',
        'label': match.group(1),
        'description': match.group(2),
        'level': 1,
        'raw_text': line
    }
```

### Parsing Different Sections

Modify `extract_munitions_list()` to target different sections:

```python
# Find start of different section
section_start = next((i for i, line in enumerate(lines)
                      if re.match(r'^YOUR_PATTERN', line)), None)
```

## Troubleshooting

### Issue: Categories Not Detected

- Check that your document uses standard DSGL format (ML1., ML2., etc.)
- Verify regex patterns in `patterns` dictionary match your format

### Issue: Items Not Properly Nested

- Review the `build_hierarchy()` logic
- Check that level assignments in `classify_line()` are correct

### Issue: Missing Content

- Ensure document loads correctly (check line count)
- Verify start/end indices for section extraction

## Contributing

Feel free to modify and extend this parser for your specific needs. The code is structured to be easily customizable.

## License

This is a custom tool. Modify and use as needed for your projects.

## Credits

Built to replace AI-based DSGL processing with deterministic, rule-based parsing.
