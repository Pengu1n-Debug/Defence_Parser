# DSGL Parser V4 - Combining best of V3 exclusions with smarter hierarchy building

import json
import re
from pathlib import Path
from typing import List, Dict, Any, Set

# Use V3's document loaders (they work perfectly)
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
        raise ImportError("python-docx not installed")

def load_document_lines(file_path):
    """Load lines from document"""
    file_path = Path(file_path)
    if file_path.suffix.lower() == '.docx':
        return load_docx_text(file_path)
    else:
        raise ValueError(f"Unsupported: {file_path.suffix}")

class DSGLParserV4:
    """Improved parser with proper hierarchy handling"""

    def __init__(self):
        self.patterns = {
            # ML category: ML1. followed by description (NOT followed by a letter and period)
            'ml_category': re.compile(r'^(ML\d+)\.\s+(?![a-z]\.\s)(.*)$'),
            # ML sub-item: ML1. a., ML1. b., etc.
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

    def identify_exclusions(self, lines: List[str]) -> Set[int]:
        """V3's perfect exclusion logic"""
        excluded_indices = set()
        i = 0

        while i < len(lines):
            line = lines[i].strip()

            if self.is_note(line):
                excluded_indices.add(i)

                starts_exclusion_list = (
                    ('does not apply' in line.lower() or 'does not control' in line.lower()) and
                    (line.endswith(':') or 'following' in line.lower() or 'as follows' in line.lower())
                ) or (
                    ('include' in line.lower() or 'includes' in line.lower()) and line.endswith(':')
                )

                if starts_exclusion_list:
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

    def build_hierarchy_smart(self, lines: List[str], excluded_indices: Set[int], start_idx: int, end_idx: int) -> List[Dict[str, Any]]:
        """
        Smarter hierarchy building:
        - Track when we're inside an ML sub-item's children
        - Properly handle transitions between ML sub-items
        """
        result = []
        current_category = None
        stack = []  # [(node, level, is_ml_subitem)]

        for i in range(start_idx, end_idx):
            if i in excluded_indices:
                continue

            line = lines[i].strip()
            if not line:
                continue

            # Check what type of line this is
            is_ml_sub = self.is_ml_subitem(line)
            is_ml_cat = self.is_ml_category(line)

            # ML Category
            if is_ml_cat:
                match = self.patterns['ml_category'].match(line)
                current_category = {
                    'Label': match.group(1),
                    'Description': match.group(2).strip(),
                    'SubStructures': []
                }
                result.append(current_category)
                stack = [(current_category, 0, False)]
                continue

            # ML sub-item (e.g., ML2. b.)
            if is_ml_sub:
                match = self.patterns['ml_subitem'].match(line)
                item = {
                    'Label': match.group(1) + '.',
                    'Description': match.group(2).strip(),
                    'SubStructures': []
                }
                # Reset stack to category level, add ML sub-item
                if current_category:
                    stack = [(current_category, 0, False)]
                    current_category['SubStructures'].append(item)
                    stack.append((item, 1, True))  # Mark as ML sub-item
                continue

            # Letter item (a., b., c.)
            match = self.patterns['item_letter'].match(line)
            if match:
                item = {
                    'Label': match.group(1) + '.',
                    'Description': match.group(2).strip(),
                    'SubStructures': []
                }

                # Determine level
                # If we're inside an ML sub-item's children and see a numbered item, this is nested
                # If we see another letter after an ML sub-item's numbered children, it's a new ML sub-item sibling
                if len(stack) >= 2 and stack[-1][2]:  # Parent is ML sub-item
                    level = 1  # Direct child of ML sub-item
                elif len(stack) >= 1 and stack[-1][1] == 2:  # Parent is a number
                    level = 3  # Nested under number
                else:
                    # Check if we should pop back to category level
                    # This happens when we finish an ML sub-item's tree and start a new one
                    if len(stack) > 1:
                        # Pop back to category
                        stack = [stack[0]]
                    level = 1

                # Adjust stack
                while len(stack) > level:
                    stack.pop()

                # Add item
                if stack:
                    stack[-1][0]['SubStructures'].append(item)
                    stack.append((item, level, False))
                continue

            # Number item (1., 2., 3.)
            match = self.patterns['item_number'].match(line)
            if match:
                item = {
                    'Label': match.group(1) + '.',
                    'Description': match.group(2).strip(),
                    'SubStructures': []
                }
                level = 2

                while len(stack) > level:
                    stack.pop()

                if stack:
                    stack[-1][0]['SubStructures'].append(item)
                    stack.append((item, level, False))
                continue

        return result

    def parse_document(self, file_path: str) -> List[Dict[str, Any]]:
        """Main parsing function"""
        print(f"Loading: {file_path}")
        lines = load_document_lines(file_path)
        print(f"Loaded {len(lines)} lines")

        print("Identifying exclusions...")
        excluded_indices = self.identify_exclusions(lines)
        print(f"Excluded {len(excluded_indices)} lines")

        # Find ML section
        ml_start = None
        for i, line in enumerate(lines):
            if re.match(r'^ML1\.\s+', line):
                ml_start = i
                break

        if ml_start is None:
            print("Error: ML1 not found")
            return []

        ml_end = len(lines)
        for i in range(ml_start + 1, len(lines)):
            if re.match(r'^(Category\s+0|Part\s+2)', lines[i], re.IGNORECASE):
                ml_end = i
                break

        print(f"Parsing ML section: lines {ml_start}-{ml_end}")
        hierarchy = self.build_hierarchy_smart(lines, excluded_indices, ml_start, ml_end)
        print(f"Created {len(hierarchy)} ML categories")

        return hierarchy

def process_dsgl_document(file_path: str, output_file: str = "dsgl_munitions_list.json"):
    """Process and save"""
    parser = DSGLParserV4()

    try:
        hierarchy = parser.parse_document(file_path)

        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(hierarchy, f, indent=2, ensure_ascii=False)

        print(f"\nSaved to: {output_file}")

        # Stats
        def count_items(structures):
            count = len(structures)
            for item in structures:
                count += count_items(item.get('SubStructures', []))
            return count

        print(f"Total ML categories: {len(hierarchy)}")
        print(f"Total items: {count_items(hierarchy)}")

        # Show first few categories
        print("\nFirst 3 categories:")
        for i, ml in enumerate(hierarchy[:3]):
            children = [x['Label'] for x in ml['SubStructures']]
            print(f"  {ml['Label']}: {children}")

        return hierarchy

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    dsgl_file = r"c:\Users\Macla\Desktop\AI\DSGL Docs\F2024L01024.docx"

    if Path(dsgl_file).exists():
        print("DSGL Parser V4 - Smart Hierarchy Building")
        print("=" * 70)
        result = process_dsgl_document(dsgl_file)
    else:
        print(f"File not found: {dsgl_file}")
