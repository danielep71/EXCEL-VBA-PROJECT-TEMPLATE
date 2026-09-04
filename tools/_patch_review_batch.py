from pathlib import Path

path = Path('tools/_apply_review_batch.py')
text = path.read_text(encoding='utf-8')
before = "    return_marker = '    return document, rule_result(\\n'\n"
after = "    return_marker = '    return (\\n'\n"
if before not in text:
    raise SystemExit('review-batch return marker patch anchor missing')
path.write_text(text.replace(before, after, 1), encoding='utf-8', newline='\n')
print('review-batch migration anchor patched')
