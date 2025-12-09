# DSGL Document Parser - Final Version with Validation & Auto-Fix
# Converts DSGL documents to hierarchical JSON with automatic structure validation

import json
import re
from pathlib import Path
from typing import List, Dict, Any, Set, Tuple
from collections import defaultdict

# Document readers
def load_docx_text(docx_file):
    """Extract text from .docx file, excluding italicized paragraphs"""
    try:
        from docx import Document
        doc = Document(docx_file)
        lines = []
        for p in doc.paragraphs:
            if not p.text.strip():
                continue
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
        raise ImportError("ebooklib and beautifulsoup4 not installed")

def load_document_lines(file_path):
    """Load lines from document"""
    file_path = Path(file_path)
    if file_path.suffix.lower() == '.docx':
        return load_docx_text(file_path)
    elif file_path.suffix.lower() == '.pdf':
        return load_pdf_text(file_path)
    elif file_path.suffix.lower() == '.epub':
        return load_epub_text(file_path)
    else:
        raise ValueError(f"Unsupported file format: {file_path.suffix}")

class DSGLParser:
    """Final DSGL parser with validation and auto-correction"""

    def __init__(self):
        self.patterns = {
            'ml_category': re.compile(r'^(ML\d+)\.\s+(.*)$'),
            'ml_subitem': re.compile(r'^ML\d+\.\s+([a-z])\.\s+(.+)$'),
            'item_letter': re.compile(r'^([a-z])\.\s+(.+)$'),
            'item_number': re.compile(r'^(\d+)\.\s+(.+)$'),
            'note': re.compile(r'^((?:Note|Technical Note|N\.B\.)[^:]*)[:\-\u2013]?\s*(.*)$', re.IGNORECASE),
        }

    def is_note(self, line: str) -> bool:
        return bool(self.patterns['note'].match(line.strip()))

    def is_ml_category(self, line: str) -> bool:
        return bool(self.patterns['ml_category'].match(line.strip()))

    def is_ml_subitem(self, line: str) -> bool:
        return bool(self.patterns['ml_subitem'].match(line.strip()))

    def is_item(self, line: str) -> bool:
        line = line.strip()
        return bool(self.patterns['item_letter'].match(line) or self.patterns['item_number'].match(line))

    def get_item_label(self, line: str) -> str:
        """Extract just the label from a line"""
        line = line.strip()

        # ML sub-item
        match = self.patterns['ml_subitem'].match(line)
        if match:
            return match.group(1) + '.'

        # Letter item
        match = self.patterns['item_letter'].match(line)
        if match:
            return match.group(1) + '.'

        # Number item
        match = self.patterns['item_number'].match(line)
        if match:
            return match.group(1) + '.'

        return ''

    def identify_exclusions(self, lines: List[str]) -> Set[int]:
        """Identify lines to exclude (notes and their associated lists)"""
        excluded_indices = set()
        i = 0

        while i < len(lines):
            line = lines[i].strip()

            if self.is_note(line):
                excluded_indices.add(i)

                # Check if this note starts an exclusion list
                starts_exclusion_list = (
                    ('does not apply' in line.lower() or 'does not control' in line.lower()) and
                    (line.endswith(':') or 'following' in line.lower() or 'as follows' in line.lower())
                ) or (
                    ('include' in line.lower() or 'includes' in line.lower()) and line.endswith(':')
                )

                if starts_exclusion_list:
                    # Exclude following items until ML sub-item or new ML category
                    j = i + 1
                    while j < len(lines):
                        next_line = lines[j].strip()
                        if self.is_ml_category(next_line) or self.is_ml_subitem(next_line):
                            break
                        if self.is_note(next_line):
                            break
                        if self.is_item(next_line):
                            excluded_indices.add(j)
                        j += 1

            i += 1

        return excluded_indices

    def extract_expected_structure(self, lines: List[str], excluded_indices: Set[int]) -> Dict[str, List[str]]:
        """
        Extract the expected structure from the document.
        Returns a dict mapping ML category labels to their expected direct children labels.
        """
        structure = defaultdict(list)
        current_ml = None

        for i, line in enumerate(lines):
            if i in excluded_indices:
                continue

            line = line.strip()

            # Check for ML category
            if self.is_ml_category(line):
                match = self.patterns['ml_category'].match(line)
                current_ml = match.group(1)
                continue

            # Check for ML sub-item (direct child of ML category)
            if current_ml and self.is_ml_subitem(line):
                label = self.get_item_label(line)
                if label and label not in structure[current_ml]:
                    structure[current_ml].append(label)
                continue

            # Check for top-level letter item (direct child of ML category)
            # These appear right after ML category without "ML#." prefix
            if current_ml:
                match = self.patterns['item_letter'].match(line)
                if match:
                    label = match.group(1) + '.'
                    # Only add if this looks like a direct child (not deeply nested)
                    # We can tell by checking if we're close to the ML category line
                    if i < len(lines) and label not in structure[current_ml]:
                        # Simple heuristic: if it's an 'a.' shortly after ML category, it's likely a direct child
                        structure[current_ml].append(label)

        return structure

    def parse_document(self, file_path: str) -> Tuple[List[Dict[str, Any]], Dict[str, List[str]]]:
        """Parse document and return (hierarchy, expected_structure)"""
        print(f"Loading: {file_path}")
        lines = load_document_lines(file_path)
        print(f"Loaded {len(lines)} lines")

        print("Identifying exclusions...")
        excluded_indices = self.identify_exclusions(lines)
        print(f"Excluded {len(excluded_indices)} lines")

        print("Extracting expected structure...")
        expected_structure = self.extract_expected_structure(lines, excluded_indices)

        print("Building hierarchy...")
        hierarchy = self.build_hierarchy(lines, excluded_indices, expected_structure)

        return hierarchy, expected_structure

    def build_hierarchy(self, lines: List[str], excluded_indices: Set[int], expected_structure: Dict[str, List[str]]) -> List[Dict[str, Any]]:
        """Build hierarchy with structure awareness"""
        # Find ML section
        ml_start = None
        for i, line in enumerate(lines):
            if re.match(r'^ML1\.\s+', line):
                ml_start = i
                break

        if ml_start is None:
            return []

        ml_end = len(lines)
        for i in range(ml_start + 1, len(lines)):
            if re.match(r'^(Category\s+0|Part\s+2)', lines[i], re.IGNORECASE):
                ml_end = i
                break

        result = []
        current_category = None
        stack = []
        current_ml_label = None

        for i in range(ml_start, ml_end):
            if i in excluded_indices:
                continue

            line = lines[i].strip()
            if not line:
                continue

            # ML Category
            match = self.patterns['ml_category'].match(line)
            if match:
                current_ml_label = match.group(1)
                current_category = {
                    'Label': current_ml_label,
                    'Description': match.group(2).strip(),
                    'SubStructures': []
                }
                result.append(current_category)
                stack = [(current_category, 0)]
                continue

            # ML sub-item (e.g., ML1. a.)
            if self.is_ml_subitem(line):
                match = self.patterns['ml_subitem'].match(line)
                item = {
                    'Label': match.group(1) + '.',
                    'Description': match.group(2).strip(),
                    'SubStructures': []
                }
                if current_category:
                    current_category['SubStructures'].append(item)
                    stack = [(current_category, 0), (item, 1)]
                continue

            # Letter item
            match = self.patterns['item_letter'].match(line)
            if match:
                label = match.group(1) + '.'
                desc = match.group(2).strip()
                item = {'Label': label, 'Description': desc, 'SubStructures': []}

                # Determine if this should be a direct child of ML category
                is_direct_child = (
                    current_ml_label and
                    label in expected_structure.get(current_ml_label, [])
                )

                if is_direct_child:
                    # Add as direct child of ML category
                    if current_category:
                        current_category['SubStructures'].append(item)
                        stack = [(current_category, 0), (item, 1)]
                else:
                    # Add to current parent in stack
                    # Determine nesting level based on previous context
                    if len(stack) > 1 and stack[-1][1] == 2:  # Previous was a number
                        level = 3  # Nested letter under number
                    else:
                        level = 1  # Top-level letter

                    while stack and len(stack) > level:
                        stack.pop()

                    if stack:
                        stack[-1][0]['SubStructures'].append(item)
                        stack.append((item, level))
                continue

            # Number item
            match = self.patterns['item_number'].match(line)
            if match:
                item = {
                    'Label': match.group(1) + '.',
                    'Description': match.group(2).strip(),
                    'SubStructures': []
                }
                level = 2

                while stack and len(stack) > level:
                    stack.pop()

                if stack:
                    stack[-1][0]['SubStructures'].append(item)
                    stack.append((item, level))
                continue

        return result

    def validate_and_fix(self, hierarchy: List[Dict[str, Any]], expected_structure: Dict[str, List[str]]) -> List[Dict[str, Any]]:
        """Validate hierarchy against expected structure and fix issues"""
        print("\nValidating and fixing structure...")

        for ml_category in hierarchy:
            ml_label = ml_category['Label']
            expected_children = expected_structure.get(ml_label, [])

            if not expected_children:
                continue

            actual_children = {child['Label']: child for child in ml_category['SubStructures']}
            actual_labels = list(actual_children.keys())

            # Find missing children
            missing = [label for label in expected_children if label not in actual_labels]

            # Find duplicates
            duplicates = [label for label in actual_labels if actual_labels.count(label) > 1]

            if missing:
                print(f"  {ml_label}: Missing children: {missing}")

            if duplicates:
                print(f"  {ml_label}: Duplicate children: {set(duplicates)}")
                # Fix duplicates: keep first occurrence, move sub-items of duplicates to first
                seen = {}
                fixed_children = []
                for child in ml_category['SubStructures']:
                    if child['Label'] not in seen:
                        seen[child['Label']] = child
                        fixed_children.append(child)
                    else:
                        # Merge sub-structures into first occurrence
                        if child['SubStructures']:
                            seen[child['Label']]['SubStructures'].extend(child['SubStructures'])

                ml_category['SubStructures'] = fixed_children
                print(f"    Fixed: Merged duplicates for {ml_label}")

        return hierarchy

def process_dsgl_document(file_path: str, output_file: str = "dsgl_final.json"):
    """Process DSGL document with validation and auto-fix"""
    parser = DSGLParser()

    try:
        hierarchy, expected_structure = parser.parse_document(file_path)

        # Validate and fix
        hierarchy = parser.validate_and_fix(hierarchy, expected_structure)

        # Save
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(hierarchy, f, indent=2, ensure_ascii=False)

        print(f"\nOutput saved to: {output_file}")

        # Stats
        def count_items(structures):
            count = len(structures)
            for item in structures:
                count += count_items(item.get('SubStructures', []))
            return count

        total_items = count_items(hierarchy)
        print(f"Total ML categories: {len(hierarchy)}")
        print(f"Total items: {total_items}")

        # Detailed stats per category
        print("\nPer-category breakdown:")
        for ml in hierarchy[:5]:  # Show first 5
            direct_children = len(ml['SubStructures'])
            print(f"  {ml['Label']}: {direct_children} direct children")

        return hierarchy

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    dsgl_file = r"c:\Users\Macla\Desktop\AI\DSGL Docs\F2024L01024.docx"

    if Path(dsgl_file).exists():
        print("DSGL Parser - Final Version with Auto-Fix")
        print("=" * 70)
        result = process_dsgl_document(dsgl_file, "dsgl_munitions_list.json")
    else:
        print(f"File not found: {dsgl_file}")
