# DSGL Parser - Context-Aware Version
# Uses context clues to determine nesting levels

import json
import re
from pathlib import Path
from typing import List, Dict, Any, Set, Optional

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
    file_path = Path(file_path)
    if file_path.suffix.lower() == '.docx':
        return load_docx_text(file_path)
    else:
        raise ValueError(f"Unsupported: {file_path.suffix}")

class DSGLContextParser:
    """Parser that uses context to determine nesting"""

    def __init__(self):
        self.patterns = {
            'ml_category': re.compile(r'^(ML\d+)\.\s+(?![a-z]\.\s)(.*)$'),
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

    def is_item_letter(self, line: str) -> bool:
        return bool(self.patterns['item_letter'].match(line.strip()))

    def is_item_number(self, line: str) -> bool:
        return bool(self.patterns['item_number'].match(line.strip()))

    def is_item(self, line: str) -> bool:
        return self.is_item_letter(line) or self.is_item_number(line)

    def identify_exclusions(self, lines: List[str]) -> Set[int]:
        """Identify lines to exclude"""
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

    def build_hierarchy(self, lines: List[str], excluded_indices: Set[int], start_idx: int, end_idx: int) -> List[Dict[str, Any]]:
        """
        Build hierarchy with context-aware nesting.

        Key rules:
        1. ML categories are top level
        2. ML sub-items (ML#. letter.) are always direct children of ML category
        3. After ML category (before any ML sub-item), letter items are direct children
        4. After ML sub-item, letter items can be:
           - Direct children of ML sub-item (level 1)
           - Or children of numbers (level 3)
        5. Number items are always level 2 (children of letter items)
        """
        result = []
        current_category = None
        current_ml_label = None
        stack = []

        # Track what we've added as direct ML children
        ml_direct_children = set()

        for i in range(start_idx, end_idx):
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
                ml_direct_children = set()
                continue

            # ML sub-item (explicitly marked: ML1. b.)
            if self.is_ml_subitem(line):
                match = self.patterns['ml_subitem'].match(line)
                label = match.group(1) + '.'
                item = {
                    'Label': label,
                    'Description': match.group(2).strip(),
                    'SubStructures': []
                }
                if current_category:
                    # Check if we already have this label
                    if label not in ml_direct_children:
                        current_category['SubStructures'].append(item)
                        ml_direct_children.add(label)
                        stack = [(current_category, 0), (item, 1)]
                    else:
                        # Duplicate - find the existing one and update stack to point to it
                        existing = next(c for c in current_category['SubStructures'] if c['Label'] == label)
                        stack = [(current_category, 0), (existing, 1)]
                continue

            # Letter item
            if self.is_item_letter(line):
                match = self.patterns['item_letter'].match(line)
                label = match.group(1) + '.'
                desc = match.group(2).strip()
                item = {'Label': label, 'Description': desc, 'SubStructures': []}

                # Decision logic based on context
                if len(stack) == 1:
                    # We're right after ML category - this is a direct child
                    if label not in ml_direct_children:
                        current_category['SubStructures'].append(item)
                        ml_direct_children.add(label)
                        stack.append((item, 1))
                    else:
                        # Duplicate at top level - shouldn't happen, but handle it
                        existing = next(c for c in current_category['SubStructures'] if c['Label'] == label)
                        stack = [(current_category, 0), (existing, 1)]
                elif len(stack) >= 2:
                    # We're nested - need to determine correct parent
                    parent_level = stack[-1][1]

                    if parent_level == 2:
                        # Parent is a number - we're deeply nested (level 3)
                        stack[-1][0]['SubStructures'].append(item)
                        stack.append((item, 3))
                    elif parent_level == 3:
                        # Parent is a deeply nested letter - this is a sibling (also level 3)
                        # Pop to parent's parent (the number at level 2)
                        stack.pop()
                        stack[-1][0]['SubStructures'].append(item)
                        stack.append((item, 3))
                    elif parent_level == 1:
                        # Parent is a letter (ML direct child)
                        # This could be a child of the letter, or a sibling
                        # If this letter follows alphabetically and we haven't seen any numbers yet,
                        # it's likely a sibling (another direct child of ML category)
                        parent_label = stack[-1][0]['Label']
                        parent_letter = parent_label[0] if len(parent_label) > 0 else ''
                        current_letter = label[0]

                        # Check if this looks like the next letter in sequence
                        is_next_letter = (ord(current_letter) == ord(parent_letter) + 1)

                        # Check if parent has any numbered children
                        parent_has_numbers = any(
                            child['Label'][0].isdigit()
                            for child in stack[-1][0].get('SubStructures', [])
                        )

                        if is_next_letter and not parent_has_numbers:
                            # This is a sibling - pop back to ML category level
                            stack.pop()
                            if current_category and label not in ml_direct_children:
                                current_category['SubStructures'].append(item)
                                ml_direct_children.add(label)
                                stack.append((item, 1))
                        else:
                            # This is a child of the current letter item
                            stack[-1][0]['SubStructures'].append(item)
                            stack.append((item, 1))
                    else:
                        # Shouldn't get here, but add to parent
                        stack[-1][0]['SubStructures'].append(item)
                        stack.append((item, parent_level + 1))
                continue

            # Number item
            if self.is_item_number(line):
                match = self.patterns['item_number'].match(line)
                label = match.group(1) + '.'
                desc = match.group(2).strip()
                item = {'Label': label, 'Description': desc, 'SubStructures': []}

                # Numbers are always children of letters (level 2)
                # Pop stack until we find a letter parent (level 1)
                while len(stack) > 2:
                    stack.pop()

                if len(stack) >= 2:
                    stack[-1][0]['SubStructures'].append(item)
                    stack.append((item, 2))
                continue

        return result

    def parse_document(self, file_path: str) -> List[Dict[str, Any]]:
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
        hierarchy = self.build_hierarchy(lines, excluded_indices, ml_start, ml_end)
        print(f"Created {len(hierarchy)} ML categories")

        return hierarchy

def process_dsgl_document(file_path: str, output_file: str = "dsgl_munitions_list.json"):
    parser = DSGLContextParser()

    try:
        hierarchy = parser.parse_document(file_path)

        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(hierarchy, f, indent=2, ensure_ascii=False)

        print(f"\nSaved to: {output_file}")

        def count_items(structures):
            count = len(structures)
            for item in structures:
                count += count_items(item.get('SubStructures', []))
            return count

        print(f"Total ML categories: {len(hierarchy)}")
        print(f"Total items: {count_items(hierarchy)}")

        print("\nFirst 5 categories:")
        for ml in hierarchy[:5]:
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
        print("DSGL Parser - Context-Aware Version")
        print("=" * 70)
        result = process_dsgl_document(dsgl_file)
    else:
        print(f"File not found: {dsgl_file}")
