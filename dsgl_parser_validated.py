# DSGL Parser - Structure-Aware Version
# Extracts expected structure first, then uses it to guide parsing

import json
import re
from pathlib import Path
from typing import List, Dict, Any, Set
from collections import defaultdict

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

class DSGLStructureAwareParser:
    """Parser that learns document structure before building hierarchy"""

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

    def is_item(self, line: str) -> bool:
        line = line.strip()
        return bool(self.patterns['item_letter'].match(line) or self.patterns['item_number'].match(line))

    def identify_exclusions(self, lines: List[str]) -> Set[int]:
        """Identify lines to exclude (notes and exclusion lists)"""
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

    def extract_ml_structure(self, lines: List[str], excluded_indices: Set[int]) -> Dict[str, Set[str]]:
        """
        First pass: Extract what items should be direct children of each ML category.
        This is done by finding all ML sub-items (ML1. a., ML2. b., etc.)
        """
        structure = defaultdict(set)

        for i, line in enumerate(lines):
            if i in excluded_indices:
                continue

            line = line.strip()
            if not line:
                continue

            # ML sub-items are explicitly marked (e.g., "ML2. b.")
            if self.is_ml_subitem(line):
                match = self.patterns['ml_subitem'].match(line)
                # Extract ML number from full line
                ml_match = re.match(r'^(ML\d+)\.', line)
                if ml_match:
                    ml_label = ml_match.group(1)
                    letter = match.group(1) + '.'
                    structure[ml_label].add(letter)

        return structure

    def build_hierarchy(self, lines: List[str], excluded_indices: Set[int], ml_structure: Dict[str, Set[str]], start_idx: int, end_idx: int) -> List[Dict[str, Any]]:
        """
        Build hierarchy using known ML structure to guide decisions.
        Key insight: Only items in ml_structure[ML#] should be direct children.
        """
        result = []
        current_category = None
        current_ml_label = None
        stack = []

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
                continue

            # ML sub-item (e.g., ML2. b.)
            if self.is_ml_subitem(line):
                match = self.patterns['ml_subitem'].match(line)
                label = match.group(1) + '.'
                item = {
                    'Label': label,
                    'Description': match.group(2).strip(),
                    'SubStructures': []
                }
                if current_category:
                    current_category['SubStructures'].append(item)
                    stack = [(current_category, 0), (item, 1)]
                continue

            # Letter item (a., b., c.)
            match = self.patterns['item_letter'].match(line)
            if match:
                label = match.group(1) + '.'
                desc = match.group(2).strip()
                item = {'Label': label, 'Description': desc, 'SubStructures': []}

                # Decision: Is this a direct child of ML category?
                # It is ONLY if it's in the ml_structure for this category
                is_direct_ml_child = (
                    current_ml_label and
                    label in ml_structure.get(current_ml_label, set())
                )

                if is_direct_ml_child:
                    # This is actually an ML sub-item without the "ML#." prefix
                    # Add it as direct child and reset stack
                    if current_category:
                        current_category['SubStructures'].append(item)
                        stack = [(current_category, 0), (item, 1)]
                else:
                    # Regular nesting logic
                    if len(stack) > 1 and stack[-1][1] == 2:  # Parent is number
                        level = 3
                    else:
                        level = 1

                    while len(stack) > level:
                        stack.pop()

                    if stack:
                        stack[-1][0]['SubStructures'].append(item)
                        stack.append((item, level))
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
                    stack.append((item, level))
                continue

        return result

    def parse_document(self, file_path: str) -> List[Dict[str, Any]]:
        print(f"Loading: {file_path}")
        lines = load_document_lines(file_path)
        print(f"Loaded {len(lines)} lines")

        print("Identifying exclusions...")
        excluded_indices = self.identify_exclusions(lines)
        print(f"Excluded {len(excluded_indices)} lines")

        print("Extracting ML structure...")
        ml_structure = self.extract_ml_structure(lines, excluded_indices)

        # Print structure for debugging
        print("\nExpected ML structure (direct children):")
        for ml_label in sorted(ml_structure.keys(), key=lambda x: int(x[2:])):
            children = sorted(ml_structure[ml_label])
            print(f"  {ml_label}: {children}")

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

        print(f"\nParsing ML section: lines {ml_start}-{ml_end}")
        hierarchy = self.build_hierarchy(lines, excluded_indices, ml_structure, ml_start, ml_end)
        print(f"Created {len(hierarchy)} ML categories")

        return hierarchy

def process_dsgl_document(file_path: str, output_file: str = "dsgl_munitions_list.json"):
    parser = DSGLStructureAwareParser()

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
        print("DSGL Parser - Structure-Aware Version")
        print("=" * 70)
        result = process_dsgl_document(dsgl_file)
    else:
        print(f"File not found: {dsgl_file}")
