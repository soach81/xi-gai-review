"""
从提纲和题库中提取速览题目。
- 提纲：提取所有知识点作为问答题
- 题库：高频考点作为填空题，冷门知识点也保留
"""
import json
import re
from html.parser import HTMLParser

# === 解析提纲 HTML ===
class OutlineParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.chapters = []
        self.current_chapter = None
        self.current_section = None
        self.current_subsection = None
        self.in_tag = None
        self.tag_stack = []

    def handle_starttag(self, tag, attrs):
        attrs_dict = dict(attrs)
        cls = attrs_dict.get('class', '')
        if tag == 'section' and 'chapter' in cls:
            self.current_chapter = attrs_dict.get('id', '')
        elif tag == 'h2' and 'ch-title' in cls:
            self.in_tag = 'ch_title'
        elif tag == 'h3' and 'sec-jie' in cls:
            self.in_tag = 'sec_title'
        elif tag == 'h4' and 'sec-mu' in cls:
            self.in_tag = 'subsec_title'
        elif tag == 'p' and 'body-text' in cls:
            self.in_tag = 'body'
        elif tag == 'strong':
            self.tag_stack.append('strong')

    def handle_endtag(self, tag):
        if self.in_tag and tag in ('h2', 'h3', 'h4', 'p'):
            self.in_tag = None
        if tag == 'strong' and self.tag_stack:
            self.tag_stack.pop()

    def handle_data(self, data):
        text = data.strip()
        if not text or len(text) < 5:
            return

        if self.in_tag == 'body' and self.current_chapter:
            # Skip page numbers and navigation
            if re.match(r'^[\d\s\-]+$', text):
                return
            if '第' in text and '节' in text and len(text) < 20:
                return
            if self.current_chapter:
                ch_id = self.current_chapter
                self.chapters.append({
                    'chapterId': ch_id,
                    'text': text,
                    'type': 'body'
                })

# === 章节名称映射 ===
CH_NAMES = {
    'guide': '导论', 'ch01': '第一章', 'ch02': '第二章', 'ch03': '第三章',
    'ch04': '第四章', 'ch05': '第五章', 'ch06': '第六章', 'ch07': '第七章',
    'ch08': '第八章', 'ch09': '第九章', 'ch10': '第十章', 'ch11': '第十一章',
    'ch12': '第十二章', 'ch13': '第十三章', 'ch14': '第十四章', 'ch15': '第十五章',
    'ch16': '第十六章', 'ch17': '第十七章'
}

# === 从提纲提取问答题 ===
def extract_qa_from_outline(parser):
    """从提纲正文中提取可以作为问答题的知识点"""
    slides = []
    item_id = 1

    # Group by chapter
    by_ch = {}
    for item in parser.chapters:
        ch = item['chapterId']
        if ch not in by_ch:
            by_ch[ch] = []
        by_ch[ch].append(item['text'])

    for ch_id, texts in by_ch.items():
        ch_name = CH_NAMES.get(ch_id, ch_id)
        for text in texts:
            # 寻找可以转化为问答题的知识点
            # 模式1：包含"是"的定义句
            if '是' in text and len(text) > 20 and len(text) < 150:
                # 找到"是"前面的部分作为问题
                parts = text.split('是', 1)
                if len(parts) == 2 and len(parts[0]) > 5:
                    question = parts[0].strip() + '是什么？'
                    answer = parts[1].strip()
                    if len(question) < 50 and len(answer) > 5 and len(answer) < 150:
                        slides.append({
                            'id': f'slide_{item_id:03d}',
                            'question': question,
                            'answer': answer,
                            'chapter': ch_name,
                            'chapterId': ch_id,
                            'type': '问答',
                            'difficulty': '基础',
                            'source': '提纲'
                        })
                        item_id += 1
                        continue

            # 模式2：包含"包括"/"包含"的列举句
            if ('包括' in text or '包含' in text) and len(text) > 20:
                parts = re.split(r'(包括|包含)', text, 1)
                if len(parts) >= 2:
                    question = parts[0].strip() + '包括哪些内容？'
                    answer = parts[2].strip() if len(parts) > 2 else ''
                    if len(question) < 50 and len(answer) > 5:
                        slides.append({
                            'id': f'slide_{item_id:03d}',
                            'question': question,
                            'answer': answer[:150],
                            'chapter': ch_name,
                            'chapterId': ch_id,
                            'type': '问答',
                            'difficulty': '基础',
                            'source': '提纲'
                        })
                        item_id += 1
                        continue

            # 模式3：包含"必须"/"需要"的实践句
            if ('必须' in text or '需要' in text or '应当' in text) and len(text) > 20:
                question = text[:30] + '...？'
                slides.append({
                    'id': f'slide_{item_id:03d}',
                    'question': question,
                    'answer': text[:150],
                    'chapter': ch_name,
                    'chapterId': ch_id,
                    'type': '问答',
                    'difficulty': '进阶',
                    'source': '提纲'
                })
                item_id += 1
                continue

            # 模式4：包含数字+核心词
            if re.search(r'[一二三四五六七八九十]+个|第[一二三四五六七八九十]+', text):
                if len(text) > 15 and len(text) < 150:
                    slides.append({
                        'id': f'slide_{item_id:03d}',
                        'question': text[:40] + '...？',
                        'answer': text[:150],
                        'chapter': ch_name,
                        'chapterId': ch_id,
                        'type': '问答',
                        'difficulty': '基础',
                        'source': '提纲'
                    })
                    item_id += 1

    return slides

# === 从题库提取填空题 ===
def extract_fill_from_qbank(qbank):
    """从选择题库中提取高频考点作为填空题"""
    slides = []
    item_id = 1000

    for ch in qbank['chapters']:
        ch_id = ch['id']
        ch_name = CH_NAMES.get(ch_id, ch.get('name', ''))

        for q in ch['questions']:
            if q['type'] != 'single':
                continue
            if not q.get('options') or not q.get('answer'):
                continue

            correct_idx = q['answer'][0] if isinstance(q['answer'], list) else q['answer']
            if correct_idx >= len(q['options']):
                continue

            correct_answer = q['options'][correct_idx]
            question = q['question']

            if len(question) < 10 or len(correct_answer) < 3:
                continue

            # 创建填空题：把正确答案挖空
            fill_q = question
            if correct_answer in fill_q:
                fill_q = fill_q.replace(correct_answer, '___（' + correct_answer + '）', 1)
            else:
                fill_q = fill_q + '___（' + correct_answer + '）'

            if len(fill_q) < 120:
                slides.append({
                    'id': f'slide_{item_id:03d}',
                    'question': fill_q,
                    'answer': correct_answer,
                    'chapter': ch_name,
                    'chapterId': ch_id,
                    'type': '填空',
                    'difficulty': '基础',
                    'source': '题库'
                })
                item_id += 1

    return slides

# === 主程序 ===
if __name__ == '__main__':
    # 解析提纲
    parser = OutlineParser()
    with open('outline.html', 'r', encoding='utf-8') as f:
        html = f.read()
    parser.feed(html)
    print(f'Parsed outline: {len(parser.chapters)} text blocks')

    # 从提纲提取问答题
    outline_slides = extract_qa_from_outline(parser)
    print(f'Outline Q&A: {len(outline_slides)} slides')

    # 从题库提取填空题
    with open('data/questions.json', 'r', encoding='utf-8') as f:
        qbank = json.load(f)
    qbank_slides = extract_fill_from_qbank(qbank)
    print(f'QBank fill-blank: {len(qbank_slides)} slides')

    # 合并
    all_slides = outline_slides + qbank_slides

    # 去重
    seen = set()
    unique = []
    for s in all_slides:
        q = s['question']
        if q not in seen:
            seen.add(q)
            unique.append(s)

    print(f'Total unique: {len(unique)} slides')

    # 统计
    ch_counts = {}
    for s in unique:
        ch = s['chapterId']
        ch_counts[ch] = ch_counts.get(ch, 0) + 1
    print('\nPer chapter:')
    for ch, cnt in sorted(ch_counts.items(), key=lambda x: -x[1]):
        print(f'  {ch}: {cnt}')

    fill = sum(1 for s in unique if s['type'] == '填空')
    qa = sum(1 for s in unique if s['type'] == '问答')
    print(f'\nFill-blank: {fill}, Q&A: {qa}')

    # 写入文件
    result = {
        'meta': {
            'subject': '习近平新时代中国特色社会主义思想概论',
            'version': '5.0',
            'totalSlides': len(unique)
        },
        'items': unique
    }

    with open('data/slides.json', 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f'\nWritten to data/slides.json')
