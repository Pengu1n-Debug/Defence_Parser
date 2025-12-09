# DSGL Document Parser V2 - Improved parser for DSGL Munitions List format
# Converts DSGL documents (.docx, .pdf, .epub) to hierarchical JSON structure matching usml.json format

import json
import re
from pathlib import Path
from typing import List, Dict, Any, Optional

# Document readers for different formats
def load_docx_text(docx_file):
    """Extract text from .docx file, excluding italicized paragraphs"""
    try:
        from docx import Document
        doc = Document(docx_file)
        lines = []
        for p in doc.paragraphs:
            if not p.text.strip():
                continue

            # Check if paragraph is entirely italicized (exclude these)
            # A paragraph is considered italicized if any run is italic
            is_italicized = any(
                run.italic or (run.font.italic if run.font else False)
                for run in p.runs
                if run.text.strip()  # Only check non-empty runs
            )

            if not is_italicized:
                lines.append(p.text)

        return lines
    except ImportError:
        raise ImportError("python-docx not installed. Run: pip install python-docx")

def load_pdf_text(pdf_file):
    """Extract text from .pdf file"""
    try:
        import PyPDF2
        lines = []
        with open(pdf_file, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page in pdf_reader.pages:
                text = page.extract_text()
                lines.extend([line.strip() for line in text.split('\n') if line.strip()])
        return lines
    except ImportError:
        raise ImportError("PyPDF2 not installed. Run: pip install PyPDF2")

def load_epub_text(epub_file):
    """Extract text from .epub file"""
    try:
        import ebooklib
        from ebooklib import epub
        from bs4 import BeautifulSoup

        book = epub.read_epub(epub_file)
        lines = []

        for item in book.get_items():
            if item.get_type() == ebooklib.ITEM_DOCUMENT:
                soup = BeautifulSoup(item.get_content(), 'html.parser')
                text = soup.get_text()
                lines.extend([line.strip() for line in text.split('\n') if line.strip()])

        return lines
    except ImportError:
        raise ImportError("ebooklib and beautifulsoup4 not installed. Run: pip install ebooklib beautifulsoup4")

def load_document_lines(file_path):
    """Load lines from document based on file extension"""
    file_path = Path(file_path)

    if file_path.suffix.lower() == '.docx':
        return load_docx_text(file_path)
    elif file_path.suffix.lower() == '.pdf':
        return load_pdf_text(file_path)
    elif file_path.suffix.lower() == '.epub':
        return load_epub_text(file_path)
    else:
        raise ValueError(f"Unsupported file format: {file_path.suffix}")

class DSGLParserV2:
    """Improved parser for DSGL documents matching usml.json structure"""

    def __init__(self):
        # Regex patterns for different hierarchical levels
        self.patterns = {
            # Main ML categories: "ML1.", "ML2.", etc. (but not "ML1.a." or "ML1. a.")
            'ml_category': re.compile(r'^(ML\d+)\.\s+([^a-z].*)$'),

            # ML sub-items like "ML1. a.", "ML1. b." (with optional space before letter)
            'ml_subitem': re.compile(r'^ML\d+\.\s+([a-z])\.\s+(.+)$'),

            # Category headers like "Category I", "Category 0"
            'category': re.compile(r'^(Category\s+[IVX0-9]+)\s*[:\-\u2013]?\s*(.*)$', re.IGNORECASE),

            # Top-level items: "a.", "b.", "c.", etc.
            'item_letter': re.compile(r'^([a-z])\.\s+(.+)$'),

            # Numbered sub-items: "1.", "2.", "3.", etc.
            'item_number': re.compile(r'^(\d+)\.\s+(.+)$'),

            # Notes: "Note:", "Note 1:", "Note to ML1.a.", "Technical Note:", "N.B.:"
            'note': re.compile(r'^((?:Note|Technical Note|N\.B\.)[^:]*)[:\-\u2013]?\s*(.*)$', re.IGNORECASE),
        }

    def determine_indentation_level(self, line: str) -> int:
        """Determine indentation level based on leading whitespace"""
        # Count leading tabs and spaces
        stripped = line.lstrip()
        indent = len(line) - len(stripped)
        return indent // 4  # Approximate indentation level

    def classify_line(self, line: str, prev_context: Dict = None) -> Dict[str, Any]:
        """Classify a line and extract its label and description"""
        line = line.strip()

        if not line:
            return {'type': 'empty', 'label': '', 'description': '', 'level': 0}

        # Check for ML sub-items first (ML1. a., ML1. b., etc.)
        match = self.patterns['ml_subitem'].match(line)
        if match:
            return {
                'type': 'item',
                'label': match.group(1) + '.',
                'label_type': 'letter',
                'description': match.group(2).strip(),
                'level': 1,
                'raw_text': line
            }

        # Check for ML category headers (ML1., ML2., etc.)
        match = self.patterns['ml_category'].match(line)
        if match:
            return {
                'type': 'category',
                'label': match.group(1),
                'description': match.group(2).strip(),
                'level': 0,
                'raw_text': line
            }

        # Check for Category headers
        match = self.patterns['category'].match(line)
        if match:
            return {
                'type': 'category',
                'label': match.group(1),
                'description': match.group(2).strip(),
                'level': 0,
                'raw_text': line
            }

        # Check for notes - these should be excluded from the structure
        match = self.patterns['note'].match(line)
        if match:
            return {
                'type': 'excluded',  # Mark as excluded instead of 'note'
                'label': match.group(1).strip(),
                'description': match.group(2).strip(),
                'level': -1,
                'raw_text': line
            }

        # Check for letter items (a., b., c., etc.)
        match = self.patterns['item_letter'].match(line)
        if match:
            # Determine if this is a top-level or nested letter
            if prev_context and prev_context.get('type') == 'item' and prev_context.get('label_type') == 'number':
                # This is a nested letter under a number
                level = 3
            else:
                # This is a top-level letter
                level = 1

            return {
                'type': 'item',
                'label': match.group(1) + '.',
                'label_type': 'letter',
                'description': match.group(2).strip(),
                'level': level,
                'raw_text': line
            }

        # Check for number items (1., 2., 3., etc.)
        match = self.patterns['item_number'].match(line)
        if match:
            return {
                'type': 'item',
                'label': match.group(1) + '.',
                'label_type': 'number',
                'description': match.group(2).strip(),
                'level': 2,
                'raw_text': line
            }

        # Default: continuation text
        return {
            'type': 'continuation',
            'label': '',
            'description': line,
            'level': -1,
            'raw_text': line
        }

    def build_hierarchy(self, lines: List[str], start_idx: int = 0, end_idx: int = None) -> List[Dict[str, Any]]:
        """Build hierarchical structure from classified lines"""
        if end_idx is None:
            end_idx = len(lines)

        result = []
        current_category = None
        stack = []  # Stack: [(node, level), ...]
        prev_context = None
        in_exclusion_section = False  # Track if we're in a "does not apply" section
        exclusion_ended_by_note = False  # Track if exclusion section was ended by another note
        level_before_exclusion = None  # Track the level of the last item before entering exclusion

        for i in range(start_idx, end_idx):
            line = lines[i]
            classified = self.classify_line(line, prev_context)

            if classified['type'] == 'empty':
                continue

            # Check if this is any kind of note
            if classified['type'] == 'excluded':
                # Check if this is a "does not apply/control" exclusion note
                if 'does not apply' in classified['description'].lower() or 'does not control' in classified['description'].lower():
                    # Check if this note indicates a list follows:
                    # - contains "following" or "as follows"
                    # - OR the full line (label + description) ends with a colon
                    note_text = classified['raw_text'].strip()
                    desc_text = classified['description'].strip()

                    has_list_indicator = (
                        'following' in desc_text.lower() or
                        'as follows' in desc_text.lower() or
                        note_text.endswith(':') or
                        desc_text.endswith(':')
                    )

                    if has_list_indicator:
                        # This starts an exclusion list - items after this are excluded
                        in_exclusion_section = True
                        exclusion_ended_by_note = False
                        # Don't set level_before_exclusion - we want to exclude all following items
                        level_before_exclusion = None
                    # else: standalone note without a list, don't enter exclusion mode
                elif in_exclusion_section:
                    # We're in an exclusion section and hit another note (Technical Note, Note 1, etc.)
                    # This signals the end of the exclusion list - next item is valid
                    exclusion_ended_by_note = True
                else:
                    # This is a note but NOT in an exclusion context
                    # Items after informational notes (like "Note 1: ... include:") should be excluded
                    # until we hit an ML sub-item or new category
                    if 'include' in classified['description'].lower():
                        # Notes that say "include:" list items that should be excluded from structure
                        in_exclusion_section = True
                        exclusion_ended_by_note = False
                        # Remember the level of the last item before entering exclusion
                        if prev_context and prev_context.get('type') == 'item':
                            level_before_exclusion = prev_context.get('level', None)
                # Skip all notes and excluded content
                continue

            # Reset exclusion section when we hit:
            # 1. A new ML category (ML1., ML2., etc.)
            # 2. An ML sub-item (ML1. a., ML1. b., etc.) - detected by checking if raw_text contains "ML\d+\."
            if classified['type'] == 'category':
                in_exclusion_section = False
                exclusion_ended_by_note = False
            elif classified['type'] == 'item':
                # Check if this is an ML sub-item (e.g., "ML1. a.", "ML2. c.")
                import re
                if re.match(r'^ML\d+\.\s+[a-z]\.', classified.get('raw_text', '')):
                    # ML sub-items always end exclusion sections
                    in_exclusion_section = False
                    exclusion_ended_by_note = False

            # Handle items in or after exclusion sections
            if classified['type'] == 'item':
                # If exclusion was ended by a note, the next item is valid - exit exclusion mode
                if exclusion_ended_by_note:
                    in_exclusion_section = False
                    exclusion_ended_by_note = False
                    level_before_exclusion = None
                    # Don't skip this item - fall through to add it
                elif in_exclusion_section:
                    # Check if this item is at the same level or higher than the item before exclusion
                    # If so, it's a sibling/parent, not part of the exclusion list
                    if level_before_exclusion is not None and classified.get('level', 999) <= level_before_exclusion:
                        # This is a sibling or parent level item - exit exclusion mode
                        in_exclusion_section = False
                        level_before_exclusion = None
                        # Don't skip this item - fall through to add it
                    else:
                        # We're still in an exclusion section - skip this item
                        continue

            if classified['type'] == 'category':
                # New ML category (e.g., ML1., ML2.)
                current_category = {
                    'Label': classified['label'],
                    'Description': classified['description'],
                    'SubStructures': []
                }
                result.append(current_category)
                stack = [(current_category, 0)]
                prev_context = classified

            elif classified['type'] == 'item':
                # Hierarchical item (a., 1., etc.)
                item = {
                    'Label': classified['label'],
                    'Description': classified['description'],
                    'SubStructures': []
                }

                level = classified['level']

                # Find appropriate parent based on level
                # Level 1: top-level letter items (a., b., c.)
                # Level 2: numbered items (1., 2., 3.)
                # Level 3: nested letter items (a., b., c.) under numbers

                # Pop stack to appropriate level
                while stack and len(stack) > level:
                    stack.pop()

                # Add item to parent
                if stack:
                    stack[-1][0]['SubStructures'].append(item)
                elif current_category:
                    current_category['SubStructures'].append(item)
                else:
                    result.append(item)

                # Push item onto stack
                stack.append((item, level))
                prev_context = classified

            elif classified['type'] == 'continuation':
                # Append continuation text to most recent item
                if stack and stack[-1][0].get('Description'):
                    stack[-1][0]['Description'] += ' ' + classified['description']
                elif stack:
                    stack[-1][0]['Description'] = classified['description']

                # Don't update prev_context for continuations

        return result

    def extract_munitions_list(self, lines: List[str]) -> List[Dict[str, Any]]:
        """Extract the munitions list section starting from ML1"""
        # Find start of munitions list (ML1)
        ml_start = None
        for i, line in enumerate(lines):
            if re.match(r'^ML1\.\s+', line):
                ml_start = i
                break

        if ml_start is None:
            print("Warning: Could not find ML1 in document. Parsing entire document...")
            return self.build_hierarchy(lines)

        print(f"Found ML1 at line {ml_start}")

        # Find end of munitions list (usually before "Category 0" or other sections)
        ml_end = None
        for i in range(ml_start + 1, len(lines)):
            if re.match(r'^(Category\s+0|Part\s+2|Dual-use list)', lines[i], re.IGNORECASE):
                ml_end = i
                break

        if ml_end is None:
            ml_end = len(lines)

        print(f"Munitions list spans lines {ml_start} to {ml_end}")

        # Parse the munitions list section
        return self.build_hierarchy(lines, ml_start, ml_end)

    def parse_document(self, file_path: str, extract_ml_only: bool = True) -> List[Dict[str, Any]]:
        """Parse a DSGL document and return hierarchical JSON structure"""
        print(f"Loading document: {file_path}")
        lines = load_document_lines(file_path)
        print(f"Loaded {len(lines)} lines")

        if extract_ml_only:
            print("Extracting munitions list (ML) section...")
            hierarchy = self.extract_munitions_list(lines)
        else:
            print("Parsing entire document...")
            hierarchy = self.build_hierarchy(lines)

        print(f"Created {len(hierarchy)} categories")
        return hierarchy

def process_dsgl_document(file_path: str, output_file: str = "dsgl_munitions_list.json", ml_only: bool = True):
    """Process a DSGL document and save to JSON file"""
    parser = DSGLParserV2()

    try:
        # Parse document
        hierarchy = parser.parse_document(file_path, extract_ml_only=ml_only)

        # Save to JSON file
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(hierarchy, f, indent=2, ensure_ascii=False)

        print(f"\nProcessing complete!")
        print(f"Output saved to: {output_file}")
        print(f"Total ML categories: {len(hierarchy)}")

        # Count total items
        def count_items(structures):
            count = len(structures)
            for item in structures:
                count += count_items(item.get('SubStructures', []))
            return count

        total_items = count_items(hierarchy)
        print(f"Total items across all categories: {total_items}")

        # Show sample
        if hierarchy:
            print("\nFirst category structure:")
            print(json.dumps(hierarchy[0], indent=2)[:800])

        return hierarchy

    except Exception as e:
        print(f"Error processing document: {e}")
        import traceback
        traceback.print_exc()
        return None

# Example usage
if __name__ == "__main__":
    dsgl_file = r"c:\Users\Macla\Desktop\AI\DSGL Docs\F2024L01024.docx"

    if Path(dsgl_file).exists():
        print("DSGL Document Parser V2")
        print("=" * 70)

        # Parse munitions list only
        result = process_dsgl_document(dsgl_file, "dsgl_munitions_list.json", ml_only=True)

        print("\n" + "=" * 70)
        print("You can also parse the entire document by setting ml_only=False")
    else:
        print(f"File not found: {dsgl_file}")
        print("Please update the file path in the script.")
