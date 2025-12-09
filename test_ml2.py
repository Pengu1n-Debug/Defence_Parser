from dsgl_parser_v3 import DSGLParserV3, load_document_lines

lines = load_document_lines(r"c:\Users\Macla\Desktop\AI\DSGL Docs\F2024L01024.docx")
parser = DSGLParserV3()
excluded = parser.identify_exclusions(lines)

print("Tracing ML2 section:")
for i in range(487, 515):
    if i in excluded:
        continue

    line = lines[i]
    is_ml_sub = parser.is_ml_subitem(line)
    classified = parser.classify_line(line, None)

    if classified['type'] != 'empty':
        ctype = classified['type']
        level = classified.get('level', '?')
        label = classified.get('label', '')
        print(f"{i}: ML_SUB={is_ml_sub:5} | type={ctype:12s} level={level} label={label:5s} | {line[:50]}")

    if 'ML3' in line:
        break
