"""
Convert textbook docx → HTML with proper hierarchy.
教材结构: 导论 + 第1-17章, 每章有若干节(第X节), 节下有目(一、二、...)
"""
from docx import Document
from pathlib import Path
import re

DOCX_PATH = Path(r"C:\Users\99114\Desktop\习概\06_教材参考\习近平新时代中国特色社会主义思想概论.docx")
OUT_PATH = Path(r"C:\Users\99114\Desktop\习概复习\textbook.html")

CH_NUMS = ["一","二","三","四","五","六","七","八","九","十",
           "十一","十二","十三","十四","十五","十六","十七","十八"]


def get_run_info(para):
    """Get font size (pt), bold status from first non-empty run."""
    for r in para.runs:
        if r.text.strip():
            sz = round(r.font.size / 12700) if r.font.size else None
            return sz, r.font.bold or False
    return None, False


def get_indent(para):
    """Get left indent in pt."""
    pf = para.paragraph_format
    return round(pf.left_indent / 12700) if pf.left_indent else 0


def clean_runs(para):
    """Extract and merge runs into HTML string."""
    runs = []
    for r in para.runs:
        t = r.text
        if not t:
            continue
        bold = r.font.bold or False
        ul = r.font.underline or False
        runs.append({'text': t, 'bold': bold, 'ul': ul})

    if not runs:
        return ""

    # Merge consecutive runs with same formatting
    merged = []
    for r in runs:
        if merged and merged[-1]['bold'] == r['bold'] and merged[-1]['ul'] == r['ul']:
            merged[-1]['text'] += r['text']
        else:
            merged.append(r)

    # Build HTML
    parts = []
    for r in merged:
        t = r['text']
        t = t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
        if r['bold'] and r['ul']:
            t = f"<strong><u>{t}</u></strong>"
        elif r['bold']:
            t = f"<strong>{t}</strong>"
        elif r['ul']:
            t = f"<u>{t}</u>"
        parts.append(t)

    return "".join(parts)


def is_page_number(text, sz):
    """Detect page number paragraphs."""
    if not text:
        return True
    if sz is not None and sz <= 10:
        cleaned = re.sub(r'[\s　]', '', text)
        if re.match(r'^[\d\-－—第页章节导论记后\W]*$', cleaned) and len(cleaned) <= 15:
            return True
    # Catch page footers with mixed chapter/section names + numbers
    cleaned = re.sub(r'[\s　]', '', text)
    if re.match(r'^\d+', cleaned) and any(w in cleaned for w in ['结','章','节']):
        if len(cleaned) <= 10:
            return True
    # Catch orphan page numbers
    if re.match(r'^\d{1,4}\s*$', text) and (sz is None or sz <= 12):
        return True
    return False


def is_citation(text, bold):
    """Detect citation/reference lines."""
    if bold:
        return False
    if re.search(r'[《《]', text) and ('页' in text or '年版' in text):
        return True
    return False


def is_chapter_title(text, sz, bold):
    """True if this paragraph is a chapter title."""
    if not bold:
        return False
    if sz is not None and sz >= 17:
        return True
    # 导论 and 后记 are also chapter-level
    if '导' in text and '论' in text and len(text) < 12:
        return True
    if '后' in text and '记' in text and len(text) < 12:
        return True
    if ('结' in text and ('语' in text or '束' in text)) and len(text) < 12:
        return True
    for num in CH_NUMS:
        if f"第{num}章" in text:
            return True
    return False


def is_section_title(text, sz, bold):
    """True if this paragraph is a section title (第X节)."""
    if not bold:
        return False
    if sz is not None and 14 <= sz < 17:
        return True
    for num in CH_NUMS:
        if f"第{num}节" in text:
            return True
    return False


def is_subsection_title(text, sz, bold):
    """True if this is a 目标题 (一、/二、...)."""
    if not bold:
        return False
    if sz is not None and sz >= 14:
        return False
    for num in CH_NUMS:
        if text.startswith(f"{num}、") or text.startswith(f"{num}  ") or text.startswith(f"{num}."):
            return True
    return False


def is_title_like(sz, bold):
    """True if paragraph looks like a title (bold + large font)."""
    return bold and sz is not None and sz >= 14


def build_html():
    doc = Document(DOCX_PATH)
    all_paras = [p for p in doc.paragraphs if p.text.strip()]

    # Step 1: Build typed entries, skipping noise
    entries = []  # list of dicts: type, html, raw_text
    for para in all_paras:
        text = para.text.strip()
        sz, bold = get_run_info(para)
        left = get_indent(para)

        # Skip noise
        if is_page_number(text, sz):
            continue
        if is_citation(text, bold):
            continue

        rendered = clean_runs(para)
        if not rendered.strip():
            continue

        if is_chapter_title(text, sz, bold):
            entries.append({'type': 'chapter', 'html': rendered, 'raw': text, 'sz': sz, 'bold': bold})
        elif is_section_title(text, sz, bold):
            entries.append({'type': 'section', 'html': rendered, 'raw': text, 'sz': sz, 'bold': bold})
        elif is_subsection_title(text, sz, bold):
            entries.append({'type': 'subsection', 'html': rendered, 'raw': text, 'sz': sz, 'bold': bold})
        else:
            entries.append({'type': 'body', 'html': rendered, 'raw': text, 'sz': sz, 'bold': bold})

    # Step 2: Merge multi-line titles (only merge adjacent entries of the SAME title type)
    merged = []
    for entry in entries:
        if entry['type'] in ('chapter', 'section'):
            if merged and merged[-1]['type'] == entry['type'] and is_title_like(entry['sz'], entry['bold']):
                # Same title type, adjacent → merge
                merged[-1]['html'] += entry['html']
                merged[-1]['raw'] += entry['raw']
                continue
        merged.append(entry)

    # Step 3: Merge orphan continuation lines into previous title
    # (e.g., "section" line followed by another bold large-font line that's the same type)
    # We already handled this in step 2 for same-type merges.
    # Now handle: chapter line -> section line (not a merge case)
    # and: title line -> body line (also not a merge)

    # Step 4: Render HTML
    html_parts = []
    ch_id = None

    for entry in merged:
        typ = entry['type']
        h = entry['html']
        raw = entry['raw']

        if typ == 'chapter':
            if ch_id:
                html_parts.append('</section>')

            # Determine chapter ID
            for idx, num in enumerate(CH_NUMS):
                if f"第{num}章" in raw:
                    ch_id = f"ch{idx+1:02d}"
                    break
            else:
                if '导' in raw and '论' in raw:
                    ch_id = "guide"
                elif '后' in raw and '记' in raw:
                    ch_id = "afterword"
                else:
                    ch_id = "unknown"

            html_parts.append(f'\n<section class="chapter" id="{ch_id}">')
            html_parts.append(f'<h2 class="ch-title">{h}</h2>')

        elif typ == 'section':
            html_parts.append(f'<h3 class="sec-jie">{h}</h3>')

        elif typ == 'subsection':
            html_parts.append(f'<h4 class="sec-mu">{h}</h4>')

        elif typ == 'body':
            html_parts.append(f'<p class="body-text">{h}</p>')

    if ch_id:
        html_parts.append('</section>')

    return '\n'.join(html_parts)


CSS = """
:root{--bg:#fff;--text:#1a1a2e;--text2:#555;--text3:#777;--ch:#1a3a5c;
  --sec:#2c5282;--hover:#f0f4ff;--border:#e2e8f0;--accent:#4a6cf7;
  --accent-light:#eef1ff;--note-bg:#fffbeb;--note-border:#f59e0b}
[data-theme="dark"]{--bg:#0f0f1a;--text:#ddd;--text2:#999;--text3:#888;
  --ch:#8ab4f8;--sec:#a0c4f8;--hover:#1a2040;--border:#2a2a45;
  --accent:#6d8af7;--accent-light:#181830;--note-bg:#1e1a0a;--note-border:#b45309}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:"PingFang SC","Microsoft YaHei","Noto Serif SC",serif;
  background:var(--bg);color:var(--text);line-height:1.9;
  padding:24px 32px 80px;max-width:860px;margin:0 auto}

/* Chapter title */
h2.ch-title{font-size:22px;font-weight:800;color:var(--ch);text-align:center;
  padding:24px 0 16px;margin:48px 0 24px;
  background:linear-gradient(to right,transparent,var(--accent),transparent);
  background-size:100% 3px;background-position:bottom;background-repeat:no-repeat;
  letter-spacing:2px}
h2.ch-title::before{content:"";display:block;width:60px;height:4px;
  background:var(--accent);margin:0 auto 12px;border-radius:2px}
h2.ch-title::after{content:"";display:block;width:40px;height:4px;
  background:var(--accent);margin:12px auto 0;border-radius:2px}

/* Section title (节) */
h3.sec-jie{font-size:17px;font-weight:700;color:var(--sec);
  margin:36px 0 14px;padding:10px 0 10px 16px;
  border-left:4px solid var(--accent);
  background:linear-gradient(90deg,var(--accent-light),transparent);
  border-radius:0 8px 8px 0}

/* Subsection title (目) */
h4.sec-mu{font-size:15px;font-weight:700;color:var(--sec);
  margin:24px 0 8px;padding-left:12px;
  border-bottom:2px dashed var(--border);padding-bottom:6px;position:relative}
h4.sec-mu::before{content:"";position:absolute;left:0;top:8px;bottom:8px;
  width:3px;background:var(--accent);border-radius:2px}

/* Body text */
.body-text{font-size:14px;color:var(--text);margin:6px 0;
  line-height:1.95;text-indent:2em;padding:4px 8px;text-align:justify}
.body-text strong{color:var(--ch)}

/* Super important (bold+underline) */
strong>u,u>strong,strong u,u strong{font-weight:700;text-decoration:underline;
  text-underline-offset:3px;text-decoration-color:var(--accent);
  text-decoration-thickness:2px}

/* TOC Navigation */
#toc-nav{position:fixed;left:0;top:0;width:240px;height:100vh;
  background:var(--bg);border-right:1px solid var(--border);
  overflow-y:auto;padding:20px 16px;z-index:100;font-size:13px;display:none;
  box-shadow:2px 0 10px rgba(0,0,0,0.08)}
#toc-nav .toc-ch{display:block;padding:6px 8px;color:var(--ch);
  text-decoration:none;font-weight:700;font-size:13px;margin-top:4px;
  border-radius:4px;transition:all .15s}
#toc-nav .toc-ch:hover{background:var(--accent-light);color:var(--accent)}
#toc-nav .toc-sec{display:block;padding:4px 8px 4px 20px;color:var(--text2);
  text-decoration:none;font-size:12px;border-radius:4px;transition:all .15s}
#toc-nav .toc-sec:hover{background:var(--hover);color:var(--accent)}
#toc-nav .toc-title{font-weight:700;font-size:15px;margin-bottom:12px;
  color:var(--ch);border-bottom:2px solid var(--accent);padding-bottom:8px}

#toc-toggle{position:fixed;left:8px;top:72px;z-index:101;
  width:36px;height:36px;border-radius:50%;background:var(--accent);
  color:#fff;border:none;font-size:16px;cursor:pointer;
  box-shadow:0 2px 8px rgba(74,108,247,0.3);transition:transform .2s}
#toc-toggle:hover{transform:scale(1.05)}

/* Back to top */
.back-top{position:fixed;bottom:80px;right:24px;width:44px;height:44px;
  border-radius:50%;background:var(--accent);color:#fff;border:none;
  font-size:20px;cursor:pointer;box-shadow:0 4px 12px rgba(74,108,247,0.3);
  z-index:50;display:none;align-items:center;justify-content:center;
  transition:transform .2s}
.back-top:hover{transform:translateY(-2px)}

/* Chapter section separator */
section.chapter{padding-bottom:40px;margin-bottom:40px;
  border-bottom:2px solid var(--border)}
section.chapter:last-child{border-bottom:none}

/* Smooth reveal */
section.chapter{animation:fadeInUp 0.5s ease}
@keyframes fadeInUp{from{opacity:0;transform:translateY(20px)}
  to{opacity:1;transform:translateY(0)}}

@media(max-width:768px){body{padding:16px 10px 80px}h2.ch-title{font-size:18px}
  h3.sec-jie{font-size:15px}h4.sec-mu{font-size:14px}}
@media(min-width:1025px){#toc-nav{display:block}#content{margin-left:240px}}
@media print{body{padding:20px}
  #toc-nav,#toc-toggle,.back-top{display:none!important}
  section.chapter{break-before:page}}
"""

TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>教材 — 习近平新时代中国特色社会主义思想概论</title>
<style>{css}</style>
</head>
<body data-theme="light">
<nav id="toc-nav">
  <div class="toc-title">📑 教材目录</div>
  <div id="toc-links"></div>
</nav>
<button id="toc-toggle" title="目录">☰</button>
<div id="content">
{body}
</div>
<button class="back-top" id="back-top" title="回到顶部">↑</button>
<script>
(function() {{
  var tocNav = document.getElementById('toc-nav');
  var tocToggle = document.getElementById('toc-toggle');

  // Build TOC from chapter and section headings
  var chapters = document.querySelectorAll('section.chapter');
  var tocContainer = document.getElementById('toc-links');
  chapters.forEach(function(ch) {{
    var chTitle = ch.querySelector('h2.ch-title');
    if (!chTitle) return;
    var chLink = document.createElement('a');
    chLink.href = '#' + ch.id;
    chLink.className = 'toc-ch';
    chLink.textContent = chTitle.textContent.trim();
    chLink.addEventListener('click', function(e) {{
      e.preventDefault();
      ch.scrollIntoView({{behavior:'smooth'}});
      if (window.innerWidth < 1025) tocNav.style.display = 'none';
    }});
    tocContainer.appendChild(chLink);

    // Add section links
    var sections = ch.querySelectorAll('h3.sec-jie');
    sections.forEach(function(sec) {{
      var secLink = document.createElement('a');
      secLink.href = '#' + ch.id;
      secLink.className = 'toc-sec';
      secLink.textContent = sec.textContent.trim();
      secLink.addEventListener('click', function(e) {{
        e.preventDefault();
        sec.scrollIntoView({{behavior:'smooth'}});
        if (window.innerWidth < 1025) tocNav.style.display = 'none';
      }});
      tocContainer.appendChild(secLink);
    }});
  }});

  // TOC toggle
  tocToggle.addEventListener('click', function() {{
    if (tocNav.style.display === 'none' || tocNav.style.display === '') {{
      tocNav.style.display = 'block';
    }} else {{
      tocNav.style.display = 'none';
    }}
  }});

  // Back to top
  var bt = document.getElementById('back-top');
  window.addEventListener('scroll', function() {{
    bt.style.display = window.scrollY > 400 ? 'flex' : 'none';
  }});
  bt.addEventListener('click', function() {{
    window.scrollTo({{top:0,behavior:'smooth'}});
  }});

  // Listen for theme changes from parent (iframe embedding)
  window.addEventListener('message', function(e) {{
    if (e.data && e.data.type === 'setTheme') {{
      document.body.setAttribute('data-theme', e.data.theme);
    }}
  }});

  // Tell parent we're ready
  if (window.parent !== window) {{
    window.parent.postMessage({{type:'ready'}}, '*');
  }}
}})();
</script>
</body>
</html>"""


if __name__ == "__main__":
    body = build_html()
    html = TEMPLATE.format(css=CSS, body=body)
    OUT_PATH.write_text(html, encoding="utf-8")
    print(f"Generated {OUT_PATH} — {len(html):,} chars")
