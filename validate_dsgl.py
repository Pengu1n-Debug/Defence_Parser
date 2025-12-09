# DSGL Validation and Issue Reporter
# Identifies structural issues in parsed DSGL output

import json
from collections import Counter

def validate_dsgl_output(json_file):
    """Validate DSGL output and report issues"""
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    print("DSGL Output Validation Report")
    print("=" * 70)
    print(f"Total ML categories: {len(data)}\n")

    issues = []

    for ml in data:
        ml_label = ml['Label']
        children = ml.get('SubStructures', [])
        child_labels = [c['Label'] for c in children]

        # Check for duplicates
        label_counts = Counter(child_labels)
        duplicates = [label for label, count in label_counts.items() if count > 1]

        if duplicates:
            issues.append({
                'category': ml_label,
                'type': 'duplicate',
                'labels': duplicates,
                'count': len(children)
            })

        # Check for missing expected sequences
        # If we have a., c. but no b., that's suspicious
        letter_labels = [l for l in child_labels if len(l) == 2 and l[0].isalpha()]
        if letter_labels:
            letters = [l[0] for l in letter_labels]
            if letters:
                first_letter = min(letters)
                last_letter = max(letters)
                expected = [chr(i) + '.' for i in range(ord(first_letter), ord(last_letter) + 1)]
                missing = [l for l in expected if l not in letter_labels]

                if missing:
                    issues.append({
                        'category': ml_label,
                        'type': 'missing_sequence',
                        'missing': missing,
                        'has': letter_labels
                    })

    # Print issues
    if issues:
        print("ISSUES FOUND:\n")
        for issue in issues:
            if issue['type'] == 'duplicate':
                print(f"[!] {issue['category']}: Duplicate children {issue['labels']}")
                print(f"   Total children: {issue['count']}\n")
            elif issue['type'] == 'missing_sequence':
                print(f"[?] {issue['category']}: Possible missing items {issue['missing']}")
                print(f"   Has: {issue['has']}\n")
    else:
        print("[OK] No structural issues found!")

    # Summary stats
    print("\nSUMMARY:")
    print(f"Categories with issues: {len(set(i['category'] for i in issues))}")
    print(f"Duplicate issues: {sum(1 for i in issues if i['type'] == 'duplicate')}")
    print(f"Missing sequence issues: {sum(1 for i in issues if i['type'] == 'missing_sequence')}")

    return issues

if __name__ == "__main__":
    issues = validate_dsgl_output("dsgl_munitions_list.json")
