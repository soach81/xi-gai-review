"""
Convert 钱学组复习提纲 docx → HTML with proper hierarchy.
"""
from docx import Document
from pathlib import Path
import re

DOCX_PATH = Path(r"C:\Users\99114\Desktop\习概\01_提纲\【钱学组】《习思想》复习提纲（第六版）.docx (1)(1).docx")
OUT_PATH = Path(r"C:\Users\99114\Desktop\习概复习\outline.html")

CH_NUMS = ["一","二","三","四","五","六","七","八","九","十",
           "十一","十二","十三","十四","十五","十六","十七"]


def get_run_info(para):
    for r in para.runs:
        if r.text.strip():
            sz = round(r.font.size / 12700) if r.font.size else 11
            return sz, r.font.bold or False, r.font.underline or False
    return 11, False, False


def get_indent(para):
    pf = para.paragraph_format
    return round(pf.left_indent / 12700) if pf.left_indent else 0


def clean_runs(para):
    """Extract text from runs, merging runs properly. Returns (html_string, is_kaiti).
    Handles: trim whitespace-only bold runs, merge adjacent bold spans, replace 'n ' marker."""
    is_kaiti = False
    runs = []
    for r in para.runs:
        t = r.text
        if not t:
            continue
        fn = (r.font.name or '').lower()
        if 'kaiti' in fn or '楷体' in fn:
            is_kaiti = True
        bold = r.font.bold or False
        ul = r.font.underline or False
        runs.append({'text': t, 'bold': bold, 'ul': ul})

    if not runs:
        return "", is_kaiti

    # Step 1: trim leading/trailing whitespace from individual runs
    # but preserve internal spaces
    for i, r in enumerate(runs):
        if i == 0:
            r['text'] = r['text'].lstrip()
        if i == len(runs) - 1:
            r['text'] = r['text'].rstrip()

    # Step 2: handle leading "n " marker
    # Check if the combined text (first few chars) starts with "n " or "n"
    combined_start = ''.join(r['text'] for r in runs)[:10]
    n_match = re.match(r'^n\s*', combined_start)
    has_n = bool(n_match)
    if has_n:
        # Remove "n" and following whitespace from the first run(s)
        n_end = n_match.end()
        for r in runs:
            if n_end <= 0:
                break
            t = r['text']
            if len(t) <= n_end:
                n_end -= len(t)
                r['text'] = ''
            else:
                r['text'] = t[n_end:]
                n_end = 0

    # Step 3: remove empty/whitespace-only runs
    runs = [r for r in runs if r['text'].strip()]

    # Step 4: merge consecutive runs with same formatting
    merged = []
    for r in runs:
        if merged and merged[-1]['bold'] == r['bold'] and merged[-1]['ul'] == r['ul']:
            merged[-1]['text'] += r['text']
        else:
            merged.append(r)

    # Step 5: build HTML, normalizing number formats
    parts = []
    for i, r in enumerate(merged):
        t = r['text']
        # Normalize knowledge-point numbering: "1     text" or "1.xxx" → "1. text"
        # Only for 1-2 digit numbers (avoids matching dates like "2025 年")
        if i == 0:
            t = re.sub(r'^(\d{1,2})\s[．\.]', r'\1. ', t)    # "1 ．" or "1 ." → "1. "
            t = re.sub(r'^(\d{1,2})\.(\S)', r'\1. \2', t)    # "1.xxx" → "1. xxx"
            t = re.sub(r'^(\d{1,2})\s{2,}', r'\1. ', t)      # "1    text" → "1. text"
            t = re.sub(r'^(\d{1,2})([一-鿿])', r'\1. \2', t)  # "1从" → "1. 从" (spaces lost in run merge)
        # Escape HTML
        t = t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
        if r['bold'] and r['ul']:
            t = f"<strong><u>{t}</u></strong>"
        elif r['bold']:
            t = f"<strong>{t}</strong>"
        elif r['ul']:
            t = f"<u>{t}</u>"
        parts.append(t)

    result = "".join(parts)
    # Prepend n-marker if needed
    if has_n:
        result = '<span class="n-marker">▪</span> ' + result
    return result, is_kaiti


def classify(text, left, sz, bold, is_kaiti=False):
    """Classify paragraph type."""
    if sz >= 14 and bold:
        return 'chapter'
    if sz == 12 and bold:
        return 'jie'
    # KaiTi paragraphs → commentary / 幕提要
    if is_kaiti:
        return 'mutiyao'
    if sz >= 11 and bold and left <= 5 and text.endswith("："):
        return 'h1'
    # Tightened h2: must be short and end with ：, not a full explanatory sentence
    if sz >= 11 and bold and 20 <= left <= 55 and text.endswith("：") and len(text) < 60:
        return 'h2'
    # Bold content at mid-indent → emphasized point (was incorrectly h2)
    if sz >= 11 and bold and 20 <= left <= 55:
        return 'em-point'
    # Numbered knowledge points (bold)
    if sz >= 9 and bold and re.match(r'^\d+[\.\s]', text):
        return 'point'
    # Numbered sub-points (non-bold, like "十个明确")
    if sz >= 9 and re.match(r'^\d+[\.\s]', text):
        return 'sub-point'
    if re.match(r'^n\s', text):
        return 'note'
    # Short colon-ending items at mid-indent → structural break
    if left >= 20 and text.endswith('：') and len(text) < 50:
        return 'sub-head'
    # Memory/understanding tips (checked last — only items not matching any structural type)
    if re.search(r'如何理解|如何记忆|理解和记忆|补充[：:]', text):
        return 'tip'
    return 'body'


def get_level(left, typ):
    """Get indent level (1-3) based on left indent."""
    if left <= 35:
        return 1
    elif left <= 58:
        return 2
    else:
        return 3


def is_chapter(text, sz, bold):
    """True if this paragraph is a chapter title."""
    if sz < 14 or not bold:
        return False
    if text == "导论":
        return True
    for num in CH_NUMS:
        if text == f"第{num}章" or text.startswith(f"第{num}章 "):
            return True
    return False


def load_outline_refs():
    """Load questions.json and build set of outlineRef anchors."""
    import json
    qpath = Path(r"C:\Users\99114\Desktop\习概复习\data\questions.json")
    refs = set()
    if qpath.exists():
        try:
            data = json.loads(qpath.read_text(encoding='utf-8'))
            for ch in data.get('chapters', []):
                for q in ch.get('questions', []):
                    ref = q.get('outlineRef', '')
                    if ref:
                        for r in ref.split(','):
                            r = r.strip()
                            if r:
                                refs.add(r)
        except Exception:
            pass
    return refs


def build_html():
    doc = Document(DOCX_PATH)
    qrefs = load_outline_refs()
    paragraphs = [p for p in doc.paragraphs if p.text.strip()]
    html = []
    ch_id = None
    sec_counter = 0
    pt_counter = 0
    in_preamble = True

    TERMINAL = set('。！？》"）;；')
    # Track consecutive point group for level normalization
    pt_group_levels = []  # list of (html_index, level)
    pt_group_section = None
    pt_group_raw_lv = None  # raw level (from get_level) for items in this group

    def flush_pt_group():
        """Normalize levels in the current point group to the minimum."""
        nonlocal pt_group_levels, pt_group_section, pt_group_raw_lv
        if len(pt_group_levels) > 1:
            min_lv = min(lv for _, lv in pt_group_levels)
            for idx, lv in pt_group_levels:
                if lv != min_lv:
                    html[idx] = html[idx].replace(f'lv{lv}"', f'lv{min_lv}"', 1)
        pt_group_levels = []
        pt_group_section = None
        pt_group_raw_lv = None

    for para in paragraphs:
        text = para.text.strip()
        left = get_indent(para)
        sz, bold, _ = get_run_info(para)

        if is_chapter(text, sz, bold):
            if ch_id:
                html.append('</section>')
            ch_id = "guide" if text == "导论" else f"ch{CH_NUMS.index(text[1:text.index('章')])+1:02d}"
            sec_counter = 0
            pt_counter = 0
            in_preamble = False
            rendered, _ = clean_runs(para)
            html.append(f'\n<section class="chapter" id="{ch_id}">')
            html.append(f'<h2 class="ch-title">{rendered}</h2>')
            continue

        if in_preamble or not ch_id:
            continue

        rendered, is_kaiti = clean_runs(para)
        if not rendered.strip():
            continue

        # Check if this is a continuation fragment of the previous element
        if html and html[-1] and not html[-1].startswith('<h') and not html[-1].startswith('<details') and not html[-1].startswith('</section'):
            # Get previous paragraph text from docx to check ending
            prev_idx = paragraphs.index(para) - 1
            while prev_idx >= 0:
                prev_text = paragraphs[prev_idx].text.strip()
                if prev_text:
                    break
                prev_idx -= 1
            if prev_idx >= 0 and prev_text and prev_text[-1] not in TERMINAL:
                prev_left = get_indent(paragraphs[prev_idx])
                if (len(text) < 60 and left > 5 and
                    not re.match(r'^\d+[\.\s]', text) and
                    not re.match(r'^n\s', text) and
                    not (sz >= 14 and bold) and
                    not (sz == 12 and bold)):
                    # Merge continuation into previous element
                    last = html[-1]
                    # Extract existing 📋 button so it ends up at the very end
                    btn_match = re.search(r'<button class="kp-btn"[^>]*>[^<]*</button>', last)
                    if btn_match:
                        last = last.replace(btn_match.group(0), '')
                        close_pos = last.rfind('</')
                        html[-1] = last[:close_pos] + rendered + btn_match.group(0) + last[close_pos:]
                    else:
                        close_pos = last.rfind('</')
                        if close_pos >= 0:
                            html[-1] = last[:close_pos] + rendered + last[close_pos:]
                        else:
                            html[-1] = last + rendered
                    continue

        typ = classify(text, left, sz, bold, is_kaiti)
        level = get_level(left, typ)

        if typ == 'jie':
            flush_pt_group()
            sec_counter += 1; pt_counter = 0
            html.append(f'<h3 class="sec-jie">{rendered}</h3>')
        elif typ == 'h1':
            flush_pt_group()
            sec_counter += 1; pt_counter = 0
            html.append(f'<h4 class="sec-h1">{rendered}</h4>')
        elif typ == 'h2':
            flush_pt_group()
            sec_counter += 1; pt_counter = 0
            html.append(f'<h5 class="sec-h2">{rendered}</h5>')
        elif typ == 'point':
            pt_counter += 1
            anchor = f"{ch_id}_s{sec_counter}_p{pt_counter}"
            btn = f'<button class="kp-btn" data-kp-id="{anchor}">📋</button>' if anchor in qrefs else ''
            html.append(f'<p class="kp-point lv{level}" id="{anchor}">{rendered}{btn}</p>')
            raw_lv = get_level(left, typ)
            if pt_group_section == sec_counter and pt_group_raw_lv is not None and abs(pt_group_raw_lv - raw_lv) <= 1:
                pt_group_levels.append((len(html)-1, level))
            else:
                flush_pt_group()
                pt_group_section = sec_counter
                pt_group_raw_lv = raw_lv
                pt_group_levels.append((len(html)-1, level))
        elif typ == 'sub-point':
            pt_counter += 1
            anchor = f"{ch_id}_s{sec_counter}_p{pt_counter}"
            btn = f'<button class="kp-btn" data-kp-id="{anchor}">📋</button>' if anchor in qrefs else ''
            html.append(f'<p class="kp-sub lv{level}" id="{anchor}">{rendered}{btn}</p>')
            raw_lv = get_level(left, typ)
            if pt_group_section == sec_counter and pt_group_raw_lv is not None and abs(pt_group_raw_lv - raw_lv) <= 1:
                pt_group_levels.append((len(html)-1, level))
            else:
                flush_pt_group()
                pt_group_section = sec_counter
                pt_group_raw_lv = raw_lv
                pt_group_levels.append((len(html)-1, level))
        elif typ == 'em-point':
            flush_pt_group()
            html.append(f'<p class="kp-em lv{level}">{rendered}</p>')
        elif typ == 'sub-head':
            flush_pt_group()
            html.append(f'<p class="kp-subhead lv{level}">{rendered}</p>')
        elif typ == 'mutiyao':
            flush_pt_group()
            html.append(f'<p class="mutiyao lv{level}">{rendered}</p>')
        elif typ == 'note':
            pt_counter += 1
            anchor = f"{ch_id}_s{sec_counter}_p{pt_counter}"
            btn = f'<button class="kp-btn" data-kp-id="{anchor}">📋</button>' if anchor in qrefs else ''
            html.append(f'<p class="kp-note lv{level}" id="{anchor}">{rendered}{btn}</p>')
            raw_lv = get_level(left, typ)
            if pt_group_section == sec_counter and pt_group_raw_lv is not None and abs(pt_group_raw_lv - raw_lv) <= 1:
                pt_group_levels.append((len(html)-1, level))
            else:
                flush_pt_group()
                pt_group_section = sec_counter
                pt_group_raw_lv = raw_lv
                pt_group_levels.append((len(html)-1, level))
        elif typ == 'tip':
            html.append(f'<details class="tip-box lv{level}" open><summary><strong>记忆提示</strong></summary>{rendered}</details>')
        elif typ == 'tip':
            flush_pt_group()
            html.append(f'<details class="tip-box lv{level}" open><summary><strong>记忆提示</strong></summary>{rendered}</details>')
        else:
            flush_pt_group()
            html.append(f'<p class="body-text lv{level}">{rendered}</p>')

    flush_pt_group()
    if ch_id:
        html.append('</section>')
    return '\n'.join(html)


CSS = """
:root{--bg:#fff;--text:#1a1a2e;--text2:#555;--text3:#777;--ch:#1a3a5c;--sec:#2c5282;
  --hover:#f0f4ff;--border:#e2e8f0;--accent:#4a6cf7;--accent-light:#eef1ff;
  --note-bg:#fffbeb;--note-border:#f59e0b}
[data-theme="dark"]{--bg:#0f0f1a;--text:#ddd;--text2:#999;--text3:#888;--ch:#8ab4f8;
  --sec:#a0c4f8;--hover:#1a2040;--border:#2a2a45;--accent:#6d8af7;--accent-light:#181830;
  --note-bg:#1e1a0a;--note-border:#b45309}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:"PingFang SC","Microsoft YaHei",sans-serif;background:var(--bg);
  color:var(--text);line-height:1.9;padding:24px 32px 80px;max-width:820px;margin:0 auto}

h2.ch-title{font-size:20px;font-weight:800;color:var(--ch);text-align:center;
  padding:20px 0 12px;margin:40px 0 20px;border-bottom:3px solid var(--accent);letter-spacing:1px}
h3.sec-jie{font-size:16px;font-weight:700;color:var(--sec);margin:30px 0 10px;
  padding:8px 0;border-bottom:2px solid var(--border)}
h4.sec-h1{font-size:14px;font-weight:700;color:var(--sec);margin:22px 0 6px;padding-left:2px}
h5.sec-h2{font-size:13px;font-weight:700;color:var(--text);margin:14px 0 4px;
  padding:4px 0 4px 10px;border-left:3px solid var(--accent)}

/* === Knowledge point (with button) === */
.kp-point{position:relative;padding:3px 8px;margin:2px 0;border-radius:4px;
  cursor:default;transition:background .15s;line-height:1.85;font-size:13px}
.kp-point:hover{background:var(--hover)}
.kp-point.lv1{padding-left:4px}
.kp-point.lv2{padding-left:28px}
.kp-point.lv3{padding-left:52px}

/* === Sub knowledge point (non-bold numbered items, lighter) === */
.kp-sub{position:relative;padding:3px 8px;margin:2px 0;border-radius:4px;
  cursor:default;transition:background .15s;line-height:1.85;font-size:13px;
  color:var(--text2)}
.kp-sub:hover{background:var(--hover)}
.kp-sub.lv1{padding-left:4px}
.kp-sub.lv2{padding-left:28px}
.kp-sub.lv3{padding-left:52px}

/* === Note / supplementary (subtle, only 📝 icon distinguishes) === */
.kp-note{position:relative;padding:3px 8px;margin:2px 0;border-radius:4px;
  cursor:default;transition:background .15s;line-height:1.85;font-size:13px}
.kp-note:hover{background:var(--hover)}
.kp-note.lv1{padding-left:4px}
.kp-note.lv2{padding-left:28px}
.kp-note.lv3{padding-left:52px}

/* === Emphasized content (bold, mid-indent, was incorrectly h2) === */
.kp-em{padding:3px 8px;margin:2px 0;font-size:13px;font-weight:600;
  color:var(--sec);line-height:1.85}
.kp-em.lv1{padding-left:4px}
.kp-em.lv2{padding-left:28px}
.kp-em.lv3{padding-left:52px}

/* === Sub heading (short colon break, was body-text) === */
.kp-subhead{padding:6px 8px 2px;margin:6px 0 0;font-size:13px;font-weight:600;
  color:var(--text);line-height:1.85}
.kp-subhead.lv1{padding-left:4px}
.kp-subhead.lv2{padding-left:28px}
.kp-subhead.lv3{padding-left:52px}

/* === 幕提要 (KaiTi commentary) === */
.mutiyao{padding:2px 8px;margin:2px 0;font-size:12px;font-family:KaiTi,STKaiti,KaiTi SC,serif;
  color:var(--text3);line-height:1.85;font-style:italic}
.mutiyao.lv2{padding-left:28px}
.mutiyao.lv3{padding-left:52px}

.n-marker{font-size:14px;margin-right:2px;display:inline}

.kp-btn{display:inline-block;margin-left:6px;vertical-align:middle;
  background:var(--accent);color:#fff;border:none;border-radius:50%;
  width:24px;height:24px;cursor:pointer;font-size:12px;line-height:24px;
  text-align:center;opacity:0;transition:opacity .2s}
.kp-point:hover .kp-btn,.kp-note:hover .kp-btn,.kp-sub:hover .kp-btn{opacity:1}
.kp-btn:hover{transform:scale(1.15)}

.body-text{font-size:12px;color:var(--text2);margin:2px 0;line-height:1.85;text-indent:2em}
.body-text.lv1{padding-left:4px}
.body-text.lv2{padding-left:28px}
.body-text.lv3{padding-left:52px}

/* === Tip / memory aid (collapsible) === */
.tip-box{margin:10px 0;padding:8px 16px;border-radius:8px;
  background:linear-gradient(135deg,#f0fdf4,#ecfdf5);border:1px solid #86efac;
  font-size:12.5px;line-height:1.85;color:#166534;cursor:default}
.tip-box summary{font-weight:700;cursor:pointer;color:#166534;font-size:13px;padding:2px 0}
.tip-box summary:hover{text-decoration:underline}
.tip-box.lv1{padding:8px 16px}
.tip-box.lv2{padding:8px 16px 8px 28px}
.tip-box.lv3{padding:8px 16px 8px 52px}
[data-theme="dark"] .tip-box{background:linear-gradient(135deg,#052e16,#064e24);border-color:#166534;color:#86efac}
[data-theme="dark"] .tip-box summary{color:#86efac}

/* Super important (bold+underline) */
strong>u,u>strong,strong u,u strong{font-weight:700;text-decoration:underline;
  text-underline-offset:3px;text-decoration-color:var(--accent);
  text-decoration-thickness:2px}

#toc-nav{position:fixed;left:0;top:0;width:230px;height:100vh;
  background:var(--bg);border-right:1px solid var(--border);
  overflow-y:auto;padding:16px 12px;z-index:100;font-size:13px;display:none}
#toc-nav .toc-link{display:block;padding:5px 0;color:var(--text2);text-decoration:none;
  border-bottom:1px solid var(--border);font-size:12.5px;transition:all .15s}
#toc-nav .toc-link:hover{color:var(--accent);padding-left:6px}
#toc-toggle{position:fixed;left:8px;top:72px;z-index:101;width:34px;height:34px;
  border-radius:50%;background:var(--accent);color:#fff;border:none;font-size:15px;
  cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.15)}
.back-top{position:fixed;bottom:80px;right:24px;width:40px;height:40px;
  border-radius:50%;background:var(--accent);color:#fff;border:none;
  font-size:18px;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.2);z-index:50;
  display:none;align-items:center;justify-content:center}
.back-top:hover{transform:translateY(-2px)}

@media(max-width:768px){body{padding:16px 10px 80px}h2.ch-title{font-size:17px}}
@media(min-width:1025px){#toc-nav{display:block}#content{margin-left:230px}}
"""

TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>复习提纲</title>
<style>{css}</style>
</head>
<body data-theme="light">
<nav id="toc-nav">
  <div style="font-weight:700;font-size:15px;margin-bottom:12px;color:var(--ch)">📑 目录</div>
  <div id="toc-links"></div>
</nav>
<button id="toc-toggle" title="目录">☰</button>
<div id="content">
{body}
</div>
<button class="back-top" id="back-top">↑</button>
<script>
(function() {{
var tocNav = document.getElementById('toc-nav');
var tocToggle = document.getElementById('toc-toggle');
var tocVisible = window.innerWidth > 1024;

document.querySelectorAll('h2.ch-title').forEach(function(h2) {{
  var a = document.createElement('a');
  a.href = '#' + h2.parentElement.id;
  a.className = 'toc-link';
  a.textContent = h2.textContent.trim();
  a.addEventListener('click', function(e) {{
    e.preventDefault();
    document.getElementById(h2.parentElement.id).scrollIntoView({{behavior:'smooth'}});
    if (window.innerWidth < 1025) tocNav.style.display='none';
  }});
  document.getElementById('toc-links').appendChild(a);
}});

tocToggle.addEventListener('click', function() {{
  if (tocNav.style.display === 'none' || tocNav.style.display === '') {{
    tocNav.style.display = 'block';
    if (window.innerWidth < 1025) document.getElementById('content').style.marginLeft = '230px';
    else tocNav.style.display = 'block';
  }} else {{
    tocNav.style.display = 'none';
    document.getElementById('content').style.marginLeft = '0';
  }}
}});

var bt = document.getElementById('back-top');
window.addEventListener('scroll', function() {{ bt.style.display = window.scrollY > 400 ? 'flex' : 'none'; }});
bt.addEventListener('click', function() {{ window.scrollTo({{top:0,behavior:'smooth'}}); }});

document.querySelectorAll('.kp-btn').forEach(function(btn) {{
  btn.addEventListener('click', function(e) {{
    e.stopPropagation();
    var txt = btn.parentElement.textContent || '';
    txt = txt.replace(/[📋]/g,'').trim().slice(0,60);
    window.parent.postMessage({{type:'openQuiz',kpId:btn.dataset.kpId,label:txt}}, '*');
  }});
}});

window.addEventListener('message', function(e) {{
  if (e.data && e.data.type === 'setTheme') document.body.setAttribute('data-theme', e.data.theme);
}});
}})();
</script>
</body>
</html>"""


if __name__ == "__main__":
    body = build_html()
    html = TEMPLATE.format(css=CSS, body=body)
    OUT_PATH.write_text(html, encoding="utf-8")
    print(f"Generated {OUT_PATH} — {len(html):,} chars")
