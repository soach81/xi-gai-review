"""
Parse 23版题库 docx files → questions.json
Format: chapter → 一、单选题 (Q+A+答案) → 二、多选题 (Q+A+答案)
"""
import re
import json
from pathlib import Path
from docx import Document

SRC_DIR = Path(r"C:\Users\99114\Desktop\习概\02_题库_主要")
OUT_PATH = Path(r"C:\Users\99114\Desktop\习概复习\data\questions.json")

CHAPTER_NAME_TO_ID = {
    "导论": ("ch00", "导论"),
    "第一章": ("ch01", "第一章 新时代坚持和发展中国特色社会主义"),
    "第二章": ("ch02", "第二章 以中国式现代化全面推进中华民族伟大复兴"),
    "第三章": ("ch03", "第三章 坚持党的全面领导"),
    "第四章": ("ch04", "第四章 坚持以人民为中心"),
    "第五章": ("ch05", "第五章 全面深化改革开放"),
    "第六章": ("ch06", "第六章 推动高质量发展"),
    "第七章": ("ch07", "第七章 社会主义现代化建设的教育科技人才战略"),
    "第八章": ("ch08", "第八章 发展全过程人民民主"),
    "第九章": ("ch09", "第九章 全面依法治国"),
    "第十章": ("ch10", "第十章 建设社会主义文化强国"),
    "第十一章": ("ch11", "第十一章 以保障和改善民生为重点加强社会建设"),
    "第十二章": ("ch12", "第十二章 建设社会主义生态文明"),
    "第十三章": ("ch13", "第十三章 维护和塑造国家安全"),
    "第十四章": ("ch14", "第十四章 建设巩固国防和强大人民军队"),
    "第十五章": ("ch15", "第十五章 坚持一国两制推进祖国完全统一"),
    "第十六章": ("ch16", "第十六章 中国特色社会主义外交和推动构建人类命运共同体"),
    "第十七章": ("ch17", "第十七章 全面从严治党"),
}


def detect_chapter(text):
    for name in CHAPTER_NAME_TO_ID:
        if text.startswith(name):
            return name
    return None


def parse_answer(answer_text):
    """Parse '正确答案：D' or '正确答案：ABC' → ['D'] or ['A','B','C']"""
    match = re.search(r'正确答案[：:]\s*([A-G\s]+)', answer_text)
    if match:
        letters = re.findall(r'[A-G]', match.group(1))
        # Filter out standalone letters that might be content
        if letters:
            return letters
    return []


def option_to_index(letter):
    """A→0, B→1, etc."""
    return ord(letter.upper()) - ord('A')


def parse_file(filepath):
    """Parse a single 题库 docx and return list of chapters with questions."""
    doc = Document(filepath)
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    chapters = []
    current_chapter = None
    current_qtype = None  # 'single' or 'multi'
    current_question = None
    qid_counter = 0

    i = 0
    while i < len(paragraphs):
        text = paragraphs[i]

        # Detect chapter
        ch_name = detect_chapter(text)
        if ch_name and (current_chapter is None or current_chapter["name"] != ch_name):
            current_chapter = {
                "id": CHAPTER_NAME_TO_ID[ch_name][0],
                "name": CHAPTER_NAME_TO_ID[ch_name][1],
                "questions": []
            }
            chapters.append(current_chapter)
            current_qtype = None
            i += 1
            continue

        # Detect question type section
        if "单选题" in text and "一、" in text[:10]:
            current_qtype = "single"
            i += 1
            continue
        if "多选题" in text and "二、" in text[:10]:
            current_qtype = "multi"
            i += 1
            continue
        # Skip other section headers
        if re.match(r'^[三四五六七八九十]、', text) and ("题" in text or "简答" in text or "判断" in text or "论述" in text):
            current_qtype = None
            i += 1
            continue

        # Parse question
        if current_qtype and current_chapter:
            q_match = re.match(r'^(\d+)[\.\、\s\)]+(.+)', text)
            if q_match:
                q_num = q_match.group(1)
                q_text = q_match.group(2).strip()

                # Look ahead for options and answer
                options = []
                answer = []
                j = i + 1
                while j < len(paragraphs):
                    next_text = paragraphs[j]
                    # Check if this is an option line (A、 or A. or A）
                    opt_match = re.match(r'^([A-G])[\.\、\）\)]\s*(.+)', next_text)
                    if opt_match:
                        options.append(opt_match.group(2).strip())
                        j += 1
                        continue
                    # Check if this is the answer line
                    if "正确答案" in next_text:
                        answer = parse_answer(next_text)
                        j += 1
                        break
                    # Check if this starts a new question
                    if re.match(r'^\d+[\.\、\s\)]', next_text) and "正确答案" not in next_text:
                        break
                    # Could be continuation of question text
                    if not re.match(r'^[A-G][\.\、\）\)]', next_text) and "正确答案" not in next_text:
                        q_text += " " + next_text
                        j += 1
                    else:
                        break

                if options and answer:
                    qid_counter += 1
                    answer_indices = [option_to_index(a) for a in answer]
                    current_chapter["questions"].append({
                        "id": f"{current_chapter['id']}_q{qid_counter:04d}",
                        "type": current_qtype,
                        "question": q_text[:500],
                        "options": options,
                        "answer": answer_indices,
                        "source": "23版题库",
                        "priority": 2
                    })

                i = j
                continue

        i += 1

    return chapters


def main():
    all_chapters = {}

    for filepath in sorted(SRC_DIR.glob("23版题库*.docx")):
        print(f"Parsing: {filepath.name}")
        chapters = parse_file(filepath)
        for ch in chapters:
            ch_id = ch["id"]
            if ch_id in all_chapters:
                all_chapters[ch_id]["questions"].extend(ch["questions"])
            else:
                all_chapters[ch_id] = ch

    # Sort by chapter ID
    result = {
        "meta": {
            "subject": "习近平新时代中国特色社会主义思想概论",
            "shortName": "习概",
            "version": "1.0",
            "totalQuestions": 0
        },
        "chapters": []
    }

    total_q = 0
    for ch_id in sorted(all_chapters.keys()):
        ch = all_chapters[ch_id]
        total_q += len(ch["questions"])
        single_count = sum(1 for q in ch["questions"] if q["type"] == "single")
        multi_count = sum(1 for q in ch["questions"] if q["type"] == "multi")
        result["chapters"].append(ch)
        print(f"  {ch['name']}: {len(ch['questions'])} questions ({single_count} single, {multi_count} multi)")

    result["meta"]["totalQuestions"] = total_q
    print(f"\nTotal: {total_q} questions across {len(result['chapters'])} chapters")

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Written to: {OUT_PATH}")


if __name__ == "__main__":
    main()
