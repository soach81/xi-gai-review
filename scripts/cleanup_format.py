"""Comprehensive formatting cleanup for outline.html."""
import re

with open(r"C:\Users\99114\Desktop\习概复习\outline.html", encoding='utf-8') as f:
    html = f.read()

fixes = 0

# 1. Chapter title: 第X章  标题 → 第X章 标题 (double space to single)
# Only in h2.ch-title
def fix_chapter_spacing(m):
    global fixes
    fixes += 1
    return m.group(0).replace('  ', ' ')
html = re.sub(r'(?<=章)\s{2,}(?=<)', ' ', html)

# 2. Section title: 第X节标题 → 第X节 标题 (add missing space)
def fix_section_spacing(m):
    global fixes
    # Check if space already exists
    num = m.group(1)
    text = m.group(2)
    if not text.startswith(' '):
        fixes += 1
        return num + ' ' + text.lstrip()
    return m.group(0)
html = re.sub(r'(第[一二三四五六七八九十]+节)(.+)', fix_section_spacing, html)

# 3. Fix extra spaces in section headings (三、  标题)
html = re.sub(r'([一二三四五六七八九十])、\s{2,}', r'\1、 ', html)

# 4. Fix number spacing: "N.    text" → "N. text" (extra spaces)
html = re.sub(r'(\d+)\.\s{3,}', r'\1. ', html)

# 5. Fix number spacing: "N.<text" → "N. <text" (no space after dot + Chinese)
html = re.sub(r'(\d+)\.([一-鿿<])', r'\1. \2', html)

# 6. Fix number spacing: "N.  text" → "N. text" (double space)
html = re.sub(r'(\d+)\.\s{2}([^\s])', r'\1. \2', html)

# 7. Fix "N    text" → "N. text" (tab-like spacess, already partially handled but some remain)
html = re.sub(r'(\d{1,2})\s{4,}([一-鿿<A-Za-z])', r'\1. \2', html)

# 8. Fix "10月" → "10 月" consistency (already OK in most places)

# 9. Fix stray double spaces in body text
html = re.sub(r'(\S)\s{2,}(\S)', r'\1 \2', html)

# 10. Fix "三、  坚定" type patterns
html = re.sub(r'([一二三四五六七八九十])、\s{2,}([^\s])', r'\1、 \2', html)

print(f"Applied {fixes} formatting fixes")

with open(r"C:\Users\99114\Desktop\习概复习\outline.html", 'w', encoding='utf-8') as f:
    f.write(html)

# Verify
counts = {}
for pattern, label in [
    (r'第[一二三四五六七八九十]+章\s{2,}', 'Chapter double space'),
    (r'第[一二三四五六七八九十]+节\S', 'Section no space'),
    (r'\d+\.\s{3,}', 'Extra spaces after dot'),
    (r'\d+\.[一-鿿]', 'No space after dot+CJK'),
    (r'[一二三四五六七八九十]、\s{2,}', 'Heading double space'),
]:
    cnt = len(re.findall(pattern, html))
    counts[label] = cnt

for label, cnt in counts.items():
    status = '✅' if cnt == 0 else f'⚠ {cnt} remaining'
    print(f'  {label}: {status}')

print(f'\nTotal size: {len(html):,} chars')
print(f'Sections: {html.count("<section class=")}')
