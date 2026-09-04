from pathlib import Path

path = Path('tools/_post_review_batch_v3.py')
text = path.read_text(encoding='utf-8')
replacements = {
    '    start = text.index("    catalogue = placeholders.get(\\\"catalogue\\\")\\n", fn_start)\n':
        '    start = text.index("        catalogue = placeholders.get(\\\"catalogue\\\")\\n", fn_start)\n',
    '    end = text.index("    block_markers = placeholders.get(\\\"block_markers\\\")\\n", start)\n':
        '    end = text.index("        block_markers = placeholders.get(\\\"block_markers\\\")\\n", start)\n',
    '    text = text[:start] + "    _validate_placeholder_catalogue(placeholders, failures)\\n" + text[end:]\n':
        '    text = text[:start] + "        _validate_placeholder_catalogue(placeholders, failures)\\n" + text[end:]\n',
}
for before, after in replacements.items():
    if before not in text:
        raise SystemExit(f'post-review patch anchor missing: {before!r}')
    text = text.replace(before, after, 1)
path.write_text(text, encoding='utf-8', newline='\n')
print('placeholder extraction indentation patched')
