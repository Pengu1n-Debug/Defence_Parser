# DSGL Document Parser V3 - Clean two-pass approach
# Converts DSGL documents (.docx, .pdf, .epub) to hierarchical JSON structure

import json
import re
from pathlib import Path
from typing import List, Dict, Any, Optional, Set

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

            # Check if paragraph is italicized (exclude these)
            is_italicized = any(
                run.italic or (run.font.italic if run.font else False)
                for run in p.runs
                if run.text.strip()
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

class DSGLParserV3:
    """Clean two-pass parser for DSGL documents"""

    def __init__(self):
        # Regex patterns
        self.patterns = {
            'ml_category': re.compile(r'^(ML\d+)\.\s+(.*)$'),
            'ml_subitem': re.compile(r'^ML\d+\.\s+([a-z])\.\s+(.+)$'),
            'item_letter': re.compile(r'^([a-z])\.\s+(.+)$'),
            'item_number': re.compile(r'^(\d+)\.\s+(.+)$'),
            'note': re.compile(r'^((?:Note|Technical Note|N\.B\.)[^:]*)[:\-\u2013]?\s*(.*)$', re.IGNORECASE),
        }

    def is_note(self, line: str) -> bool:
        """Check if line is a note"""
        return bool(self.patterns['note'].match(line.strip()))

    def is_ml_category(self, line: str) -> bool:
        """Check if line is an ML category (ML1., ML2., etc.)"""
        return bool(self.patterns['ml_category'].match(line.strip()))

    def is_ml_subitem(self, line: str) -> bool:
        """Check if line is an ML sub-item (ML1. a., ML2. b., etc.)"""
        return bool(self.patterns['ml_subitem'].match(line.strip()))

    def is_item(self, line: str) -> bool:
        """Check if line is an item (a., b., 1., 2., etc.)"""
        line = line.strip()
        return bool(self.patterns['item_letter'].match(line) or self.patterns['item_number'].match(line))

    def identify_exclusions(self, lines: List[str]) -> Set[int]:
        """
        First pass: Identify line indices that should be excluded.
        Returns a set of line indices to skip during parsing.
        """
        excluded_indices = set()
        i = 0

        while i < len(lines):
            line = lines[i].strip()

            # Check if this is a note
            if self.is_note(line):
                # Always exclude the note itself
                excluded_indices.add(i)

                # Check if this note starts an exclusion list
                # Patterns that indicate a list follows:
                # 1. "does not apply to:" or "does not apply to the following:"
                # 2. "include:" or "includes:"
                starts_exclusion_list = (
                    ('does not apply' in line.lower() or 'does not control' in line.lower()) and
                    (line.endswith(':') or 'following' in line.lower() or 'as follows' in line.lower())
                ) or (
                    ('include' in line.lower() or 'includes' in line.lower()) and line.endswith(':')
                )

                if starts_exclusion_list:
                    # Exclude all following items until we hit an ML sub-item or new ML category
                    j = i + 1
                    while j < len(lines):
                        next_line = lines[j].strip()

                        # Stop if we hit a new ML category or ML sub-item
                        if self.is_ml_category(next_line) or self.is_ml_subitem(next_line):
                            break

                        # Stop if we hit another note (marks end of exclusion list)
                        if self.is_note(next_line):
                            break

                        # If it's an item, exclude it
                        if self.is_item(next_line):
                            excluded_indices.add(j)

                        j += 1

            i += 1

        return excluded_indices

    def classify_line(self, line: str, prev_context: Dict = None) -> Dict[str, Any]:
        """Classify a line and extract its structure"""
        line = line.strip()

        if not line:
            return {'type': 'empty'}

        # Check for ML sub-items (ML1. a., ML2. b.)
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

        # Check for ML category (ML1., ML2.)
        match = self.patterns['ml_category'].match(line)
        if match:
            return {
                'type': 'category',
                'label': match.group(1),
                'description': match.group(2).strip(),
                'level': 0,
                'raw_text': line
            }

        # Check for letter items (a., b., c.)
        match = self.patterns['item_letter'].match(line)
        if match:
            # Determine level based on context
            if prev_context and prev_context.get('label_type') == 'number':
                level = 3  # Nested under a number
            else:
                level = 1  # Top-level

            return {
                'type': 'item',
                'label': match.group(1) + '.',
                'label_type': 'letter',
                'description': match.group(2).strip(),
                'level': level,
                'raw_text': line
            }

        # Check for number items (1., 2., 3.)
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

        # Continuation text
        return {
            'type': 'continuation',
            'description': line,
            'raw_text': line
        }

    def build_hierarchy(self, lines: List[str], excluded_indices: Set[int], start_idx: int = 0, end_idx: int = None) -> List[Dict[str, Any]]:
        """Build hierarchical structure, skipping excluded indices"""
        if end_idx is None:
            end_idx = len(lines)

        result = []
        current_category = None
        stack = []
        prev_context = None
        last_ml_subitem_level = None  # Track if we just processed an ML sub-item

        for i in range(start_idx, end_idx):
            # Skip excluded lines
            if i in excluded_indices:
                continue

            line = lines[i]

            # Check if this is an ML sub-item BEFORE classification
            is_ml_sub = self.is_ml_subitem(line)

            classified = self.classify_line(line, prev_context)

            if classified['type'] == 'empty':
                continue

            if classified['type'] == 'category':
                # New ML category
                current_category = {
                    'Label': classified['label'],
                    'Description': classified['description'],
                    'SubStructures': []
                }
                result.append(current_category)
                stack = [(current_category, 0)]
                prev_context = classified
                last_ml_subitem_level = None

            elif classified['type'] == 'item':
                # Create item
                item = {
                    'Label': classified['label'],
                    'Description': classified['description'],
                    'SubStructures': []
                }

                level = classified['level']

                # Special handling: if previous was an ML sub-item and current is level-1 letter,
                # treat it as a sibling (both are direct children of ML category)
                if last_ml_subitem_level == 1 and level == 1 and not is_ml_sub:
                    # Reset stack to category level (depth 1)
                    while stack and len(stack) > 1:
                        stack.pop()
                else:
                    # Normal case: adjust stack to appropriate level
                    while stack and len(stack) > level:
                        stack.pop()

                # Add to parent
                if stack:
                    stack[-1][0]['SubStructures'].append(item)
                elif current_category:
                    current_category['SubStructures'].append(item)
                else:
                    result.append(item)

                # Push onto stack
                stack.append((item, level))
                prev_context = classified

                # Track if this was an ML sub-item
                # Don't reset if we're processing children of an ML sub-item
                if is_ml_sub:
                    last_ml_subitem_level = level
                elif level > 1:
                    # This is a child item (numbered or nested letter), keep the ML sub-item marker
                    pass
                else:
                    # This is a level-1 item that's not an ML sub-item, reset the marker
                    last_ml_subitem_level = None

            elif classified['type'] == 'continuation':
                # Append to last item's description
                if stack:
                    if stack[-1][0].get('Description'):
                        stack[-1][0]['Description'] += ' ' + classified['description']
                    else:
                        stack[-1][0]['Description'] = classified['description']

        return result

    def extract_munitions_list(self, lines: List[str], excluded_indices: Set[int]) -> List[Dict[str, Any]]:
        """Extract munitions list section"""
        # Find ML1
        ml_start = None
        for i, line in enumerate(lines):
            if re.match(r'^ML1\.\s+', line):
                ml_start = i
                break

        if ml_start is None:
            print("Warning: ML1 not found")
            return self.build_hierarchy(lines, excluded_indices)

        # Find end (usually before Category 0 or Part 2)
        ml_end = len(lines)
        for i in range(ml_start + 1, len(lines)):
            if re.match(r'^(Category\s+0|Part\s+2)', lines[i], re.IGNORECASE):
                ml_end = i
                break

        print(f"Munitions list: lines {ml_start} to {ml_end}")
        return self.build_hierarchy(lines, excluded_indices, ml_start, ml_end)

    def parse_document(self, file_path: str, extract_ml_only: bool = True) -> List[Dict[str, Any]]:
        """Parse DSGL document"""
        print(f"Loading: {file_path}")
        lines = load_document_lines(file_path)
        print(f"Loaded {len(lines)} lines")

        # First pass: identify exclusions
        print("Identifying exclusions...")
        excluded_indices = self.identify_exclusions(lines)
        print(f"Excluded {len(excluded_indices)} lines")

        # Second pass: build hierarchy
        if extract_ml_only:
            hierarchy = self.extract_munitions_list(lines, excluded_indices)
        else:
            hierarchy = self.build_hierarchy(lines, excluded_indices)

        print(f"Created {len(hierarchy)} categories")
        return hierarchy

def process_dsgl_document(file_path: str, output_file: str = "dsgl_parsed.json", ml_only: bool = True):
    """Process DSGL document and save to JSON"""
    parser = DSGLParserV3()

    try:
        hierarchy = parser.parse_document(file_path, extract_ml_only=ml_only)

        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(hierarchy, f, indent=2, ensure_ascii=False)

        print(f"\nOutput saved to: {output_file}")

        # Count items
        def count_items(structures):
            count = len(structures)
            for item in structures:
                count += count_items(item.get('SubStructures', []))
            return count

        total_items = count_items(hierarchy)
        print(f"Total items: {total_items}")

        if hierarchy:
            print("\nFirst category:")
            print(json.dumps(hierarchy[0], indent=2, ensure_ascii=False)[:600])

        return hierarchy

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    dsgl_file = r"c:\Users\Macla\Desktop\AI\DSGL Docs\F2024L01024.docx"

    if Path(dsgl_file).exists():
        print("DSGL Parser V3 - Two-Pass Approach")
        print("=" * 70)
        result = process_dsgl_document(dsgl_file, "dsgl_munitions_list.json", ml_only=True)
    else:
        print(f"File not found: {dsgl_file}")
