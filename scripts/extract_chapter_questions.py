#!/usr/bin/env python3
"""
从章节练习题 doc/docx 文件中提取选择题，追加到 questions.json。

支持三种格式：
A. （一）单选题 + 末尾集中答案（参考答案✦ 1.B 2.D ...）
B. 一、单选题 + 每题后答案（参考答案：B）
C. 单选题（无编号）+ 末尾集中答案
"""

import json
import os
import re
import subprocess
import sys
from pathlib import Path

# === 配置 ===
CHAPTER_DIR = Path(r"C:\Users\99114\Desktop\习概\习题\最新新思想各章节题库")
QUESTIONS_FILE = Path(r"C:\Users\99114\Desktop\习概复习\data\questions.json")

CHAPTER_MAP = {
    "导论": "ch00",
    "第一章": "ch01",
    "第二章": "ch02",
    "第三章": "ch03",
    "第四章": "ch04",
    "第五章": "ch05",
    "第六章": "ch06",
    "第七章": "ch07",
    "第八章": "ch08",
    "第九章": "ch09",
    "第十章": "ch10",
    "第十一章": "ch11",
    "第十二章": "ch12",
    "第十三章": "ch13",
    "第十四章": "ch14",
    "第十五章": "ch15",
    "第十六章": "ch16",
    "第十七章": "ch17",
}


def read_doc(filepath):
    result = subprocess.run(
        ["antiword", "-m", "UTF-8", str(filepath)],
        capture_output=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"antiword failed for {filepath}: {result.stderr}")
    return result.stdout.decode("utf-8", errors="replace")


def read_docx(filepath):
    from docx import Document
    doc = Document(str(filepath))
    return "\n".join(p.text for p in doc.paragraphs)


def read_file(filepath):
    filepath = Path(filepath)
    if filepath.suffix == ".doc":
        return read_doc(filepath)
    elif filepath.suffix == ".docx":
        return read_docx(filepath)
    else:
        raise ValueError(f"Unsupported file type: {filepath.suffix}")


def is_single_header(line):
    """判断是否为单选题 section header"""
    if "单选题" not in line:
        return False
    # 匹配: （一）单选题, 一、单选题, 单选题
    if re.search(r"[（(]一[）)]", line):
        return True
    if re.search(r"一[、．.]", line):
        return True
    # 纯粹的 "单选题" (无编号前缀)
    if line.strip() == "单选题":
        return True
    return False


def is_multi_header(line):
    """判断是否为多选题 section header"""
    if "多选题" not in line:
        return False
    if re.search(r"[（(]二[）)]", line):
        return True
    if re.search(r"二[、．.]", line):
        return True
    if line.strip() == "多选题":
        return True
    return False


def parse_questions(text):
    """
    解析题目文本，返回 (single_questions, multi_questions)。
    """
    lines = text.split("\n")

    section = None  # 'single' or 'multi'
    current_question = None
    current_options = []

    single_questions = []
    multi_questions = []

    # 格式A/B 共用：收集题目和答案
    pending_questions = []  # [(section, q_num, q_text, options)]
    answers_map = {}  # q_num -> answer_str (格式A: 末尾集中答案)

    # 格式B: 每题后答案
    per_question_answer = None

    # 格式A: 集中答案在下一行
    expecting_answers = False

    for i, line in enumerate(lines):
        line_stripped = line.strip()
        if not line_stripped:
            continue

        # 检测 section header
        if is_single_header(line_stripped):
            _flush_pending(pending_questions, answers_map, single_questions, multi_questions)
            section = "single"
            pending_questions = []
            answers_map = {}
            current_question = None
            current_options = []
            expecting_answers = False
            continue
        if is_multi_header(line_stripped):
            _flush_pending(pending_questions, answers_map, single_questions, multi_questions)
            section = "multi"
            pending_questions = []
            answers_map = {}
            current_question = None
            current_options = []
            expecting_answers = False
            continue

        # 格式A: 集中答案行（参考答案✦ 或 参考答案）
        if "参考答案" in line_stripped and section is not None:
            # 检查是否为集中答案格式（如 "1.B  2.D  3.B"）
            answer_matches = re.findall(r"(\d+)\s*[.．、]\s*([A-F]+)", line_stripped)
            if answer_matches:
                # 格式A: 集中答案
                for q_num_str, ans_str in answer_matches:
                    answers_map[int(q_num_str)] = ans_str.upper()
                continue
            else:
                # 格式B: 可能是 "参考答案：B" 或 "参考答案✦"（无具体答案）
                per_match = re.search(r"参考答案[：:✦]\s*([A-F]+)", line_stripped)
                if per_match:
                    per_question_answer = per_match.group(1).upper()
                    # 保存当前题目
                    if current_question is not None and current_options:
                        pending_questions.append(
                            (section, current_question["num"], current_question["text"], list(current_options), per_question_answer)
                        )
                        current_question = None
                        current_options = []
                        per_question_answer = None
                    continue
                else:
                    # "参考答案✦" 但没有答案在这一行，可能是格式A的答案在下一行
                    expecting_answers = True
                    continue

        # 格式B: 独立的答案行（如 "参考答案：B"）
        per_match = re.match(r"参考答案[：:]\s*([A-F]+)", line_stripped)
        if per_match and section is not None:
            per_question_answer = per_match.group(1).upper()
            if current_question is not None and current_options:
                pending_questions.append(
                    (section, current_question["num"], current_question["text"], list(current_options), per_question_answer)
                )
                current_question = None
                current_options = []
                per_question_answer = None
            continue

        # 跳过知识点/解析行
        if "知识点" in line_stripped and "：" in line_stripped:
            continue
        if "答案要点" in line_stripped:
            continue

        if section is None:
            continue

        # 格式A: 集中答案在下一行（如 "1.A    2.B  3.B"）
        if expecting_answers:
            answer_matches = re.findall(r"(\d+)\s*[.．、]\s*([A-F]+)", line_stripped)
            if answer_matches:
                for q_num_str, ans_str in answer_matches:
                    answers_map[int(q_num_str)] = ans_str.upper()
                expecting_answers = False
                continue
            else:
                expecting_answers = False
                # 不是答案行，继续正常处理

        # 匹配题号
        q_match = re.match(r"^(\d+)[.．、]\s*(.+)", line_stripped)
        if q_match:
            # 保存上一题（如果有 per_question_answer 就已经在上面保存了）
            if current_question is not None and current_options and per_question_answer is not None:
                pending_questions.append(
                    (section, current_question["num"], current_question["text"], list(current_options), per_question_answer)
                )
                per_question_answer = None
            elif current_question is not None and current_options:
                # 格式A: 答案稍后统一解析
                pending_questions.append(
                    (section, current_question["num"], current_question["text"], list(current_options), None)
                )
            q_num = int(q_match.group(1))
            q_text = q_match.group(2).strip()
            current_question = {"num": q_num, "text": q_text}
            current_options = []
            continue

        # 匹配选项
        opt_match = re.match(r"^([A-F])[.．、]\s*(.+)", line_stripped)
        if opt_match and current_question is not None:
            current_options.append(opt_match.group(2).strip())
            continue

        # 题目跨行续行
        if current_question is not None and not current_options:
            current_question["text"] += line_stripped

    # 保存最后一题
    if current_question is not None and current_options:
        if per_question_answer is not None:
            pending_questions.append(
                (section, current_question["num"], current_question["text"], list(current_options), per_question_answer)
            )
        else:
            pending_questions.append(
                (section, current_question["num"], current_question["text"], list(current_options), None)
            )

    # 最后 flush
    _flush_pending(pending_questions, answers_map, single_questions, multi_questions)

    return single_questions, multi_questions


def _flush_pending(pending, answers_map, single_list, multi_list):
    """将 pending 题目与答案组装成最终结果"""
    for item in pending:
        if len(item) == 5:
            section, q_num, q_text, options, per_answer = item
        else:
            section, q_num, q_text, options = item
            per_answer = None

        # 确定答案
        answer_str = ""
        if per_answer:
            answer_str = per_answer
        elif q_num in answers_map:
            answer_str = answers_map[q_num]

        if not answer_str:
            continue  # 没有答案，跳过

        # 转换答案字母为索引
        answer_indices = []
        for ch in answer_str:
            idx = ord(ch) - ord("A")
            if 0 <= idx < len(options):
                answer_indices.append(idx)

        if not answer_indices:
            continue

        q_type = "single" if section == "single" else "multi"

        question_entry = {
            "type": q_type,
            "question": q_text.strip(),
            "options": options,
            "answer": sorted(answer_indices),
        }

        if section == "single":
            single_list.append(question_entry)
        else:
            multi_list.append(question_entry)


def get_chapter_id(filename):
    for prefix, ch_id in CHAPTER_MAP.items():
        if filename.startswith(prefix):
            return ch_id
    return None


def get_chapter_name(text, ch_id):
    lines = text.strip().split("\n")
    for line in lines:
        line = line.strip()
        if line and len(line) > 3:
            name = re.sub(r"\s+", " ", line)
            return name
    return ch_id


def main():
    with open(QUESTIONS_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    existing_questions = set()
    existing_ids = set()
    for ch in data.get("chapters", []):
        for q in ch.get("questions", []):
            existing_questions.add(q["question"].strip())
            existing_ids.add(q["id"])

    print(f"现有题目数: {len(existing_questions)}")

    new_chapters = {}

    files = sorted(CHAPTER_DIR.iterdir())
    for fpath in files:
        if fpath.suffix not in (".doc", ".docx"):
            continue

        ch_id = get_chapter_id(fpath.name)
        if ch_id is None:
            print(f"  跳过无法识别的文件: {fpath.name}")
            continue

        print(f"处理: {fpath.name} -> {ch_id}")

        try:
            text = read_file(fpath)
        except Exception as e:
            print(f"  读取失败: {e}")
            continue

        single_qs, multi_qs = parse_questions(text)
        all_new = single_qs + multi_qs
        print(f"  提取到 {len(single_qs)} 单选 + {len(multi_qs)} 多选 = {len(all_new)} 题")

        if ch_id not in new_chapters:
            ch_name = get_chapter_name(text, ch_id)
            new_chapters[ch_id] = {"name": ch_name, "questions": []}

        for q in all_new:
            if q["question"] in existing_questions:
                continue
            existing_questions.add(q["question"])

            idx = len(new_chapters[ch_id]["questions"]) + 1
            while True:
                new_id = f"{ch_id}_q{idx:04d}"
                if new_id not in existing_ids:
                    break
                idx += 1
            existing_ids.add(new_id)

            q["id"] = new_id
            q["source"] = "章节练习"
            q["priority"] = 2
            new_chapters[ch_id]["questions"].append(q)

    total_new = 0
    for ch_id, ch_data in sorted(new_chapters.items()):
        if not ch_data["questions"]:
            continue

        found = False
        for ch in data["chapters"]:
            if ch["id"] == ch_id:
                ch["questions"].extend(ch_data["questions"])
                found = True
                break

        if not found:
            data["chapters"].append(
                {
                    "id": ch_id,
                    "name": ch_data["name"],
                    "questions": ch_data["questions"],
                }
            )

        total_new += len(ch_data["questions"])
        print(f"  {ch_id}: 新增 {len(ch_data['questions'])} 题")

    total = sum(len(ch["questions"]) for ch in data["chapters"])
    data["meta"]["totalQuestions"] = total
    data["chapters"].sort(key=lambda ch: ch["id"])

    with open(QUESTIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"\n完成! 新增 {total_new} 题, 总计 {total} 题")


if __name__ == "__main__":
    main()
