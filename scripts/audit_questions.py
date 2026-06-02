#!/usr/bin/env python3
"""审计 questions.json 选择题数据质量，找出有问题的题目。"""

import json
import re
import sys
import io
from collections import defaultdict

# 强制 UTF-8 输出，避免 Windows GBK 编码问题
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

DATA_PATH = r"C:\Users\99114\Desktop\习概复习\data\questions.json"

# 合并选项中常见的字母标记（中文顿号、英文句点、中文冒号等）
MERGED_PATTERN = re.compile(
    r'^(A\s*[、.．：:]\s*|B\s*[、.．：:]\s*|C\s*[、.．：:]\s*|D\s*[、.．：:]\s*)',
    re.IGNORECASE
)


def check_question(q, chapter_id):
    """检查单个题目，返回问题列表。"""
    problems = []
    q_id = q.get("id", "unknown")
    q_type = q.get("type", "unknown")
    q_text = q.get("question", "")[:50]
    options = q.get("options", [])
    answer = q.get("answer", [])

    # 1. 选项数量不足：单选题应有4个选项，多选题应有>=4个选项
    if q_type == "single" and len(options) < 4:
        problems.append(f"选项不足: 单选题只有{len(options)}个选项（应有4个）")
    elif q_type == "multi" and len(options) < 4:
        problems.append(f"选项不足: 多选题只有{len(options)}个选项（应至少4个）")

    # 2. 选项合并错误：选项文本以 "A、" "A." 等开头，说明多个选项被合并到一个里
    for i, opt in enumerate(options):
        if MERGED_PATTERN.match(opt.strip()):
            problems.append(f"选项合并错误: 第{i+1}个选项以标记字母开头（'{opt[:30]}'）")

    # 3. 答案越界
    for ans_idx in answer:
        if ans_idx >= len(options) or ans_idx < 0:
            problems.append(f"答案越界: answer索引{ans_idx}超出options范围（共{len(options)}个选项）")

    # 4. 空选项
    for i, opt in enumerate(options):
        if not opt or not opt.strip():
            problems.append(f"空选项: 第{i+1}个选项为空")

    # 5. 重复选项
    seen = {}
    for i, opt in enumerate(options):
        opt_clean = opt.strip()
        if opt_clean in seen:
            problems.append(f"重复选项: 第{i+1}个选项与第{seen[opt_clean]+1}个选项相同")
        else:
            seen[opt_clean] = i

    return problems


def main():
    with open(DATA_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)

    total = 0
    problem_count = 0
    problems_by_type = defaultdict(int)  # 问题类型 -> 数量
    problems_by_chapter = defaultdict(int)  # 章节 -> 数量
    merged_by_chapter = defaultdict(int)  # 章节 -> 合并错误数量

    all_problems = []

    for ch in data.get("chapters", []):
        ch_id = ch.get("id", "unknown")
        ch_name = ch.get("name", "unknown")
        for q in ch.get("questions", []):
            total += 1
            q_type = q.get("type", "unknown")
            issues = check_question(q, ch_id)
            if issues:
                problem_count += 1
                problems_by_chapter[ch_id] += 1
                for issue in issues:
                    # 分类问题类型
                    if "选项合并错误" in issue:
                        problems_by_type["选项合并错误"] += 1
                        merged_by_chapter[ch_id] += 1
                    elif "选项不足" in issue:
                        problems_by_type["选项不足"] += 1
                    elif "答案越界" in issue:
                        problems_by_type["答案越界"] += 1
                    elif "空选项" in issue:
                        problems_by_type["空选项"] += 1
                    elif "重复选项" in issue:
                        problems_by_type["重复选项"] += 1

                all_problems.append({
                    "id": q.get("id"),
                    "chapter": ch_id,
                    "chapter_name": ch_name,
                    "type": q_type,
                    "text": q.get("question", "")[:50],
                    "options_count": len(q.get("options", [])),
                    "issues": issues,
                })

    # ========== 输出报告 ==========
    print("=" * 80)
    print("questions.json 选择题数据质量审计报告")
    print("=" * 80)
    print(f"\n总题目数: {total}")
    print(f"有问题题目数: {problem_count} ({problem_count/total*100:.1f}%)")

    print(f"\n{'─' * 60}")
    print("各类问题统计:")
    print(f"{'─' * 60}")
    for issue_type, count in sorted(problems_by_type.items(), key=lambda x: -x[1]):
        print(f"  {issue_type}: {count}")

    print(f"\n{'─' * 60}")
    print("各章节问题分布:")
    print(f"{'─' * 60}")
    for ch_id, count in sorted(problems_by_chapter.items(), key=lambda x: -x[1]):
        ch_name = next((ch["name"] for ch in data["chapters"] if ch["id"] == ch_id), ch_id)
        merged_info = f"（其中合并错误: {merged_by_chapter[ch_id]}）" if merged_by_chapter.get(ch_id, 0) > 0 else ""
        print(f"  {ch_id} ({ch_name}): {count}题有问题 {merged_info}")

    # 输出所有有问题的题目详情
    print(f"\n{'=' * 80}")
    print("有问题题目详情:")
    print(f"{'=' * 80}")

    for p in all_problems:
        print(f"\n  ID: {p['id']}")
        print(f"  类型: {'单选' if p['type'] == 'single' else '多选' if p['type'] == 'multi' else p['type']}")
        print(f"  章节: {p['chapter']} ({p['chapter_name']})")
        print(f"  选项数: {p['options_count']}")
        print(f"  题目: {p['text']}...")
        for issue in p["issues"]:
            print(f"    >> {issue}")

    # 特别关注：选项合并错误最严重的章节
    print(f"\n{'=' * 80}")
    print("特别关注：选项合并错误（完全不可用的题目）")
    print(f"{'=' * 80}")
    merged_questions = [p for p in all_problems if any("合并错误" in i for i in p["issues"])]
    if merged_questions:
        print(f"共 {len(merged_questions)} 道题目存在选项合并错误:")
        for p in merged_questions:
            print(f"  {p['id']} ({p['chapter']}) - {p['text']}...")
            for issue in p["issues"]:
                if "合并错误" in issue:
                    print(f"    >> {issue}")
    else:
        print("未发现选项合并错误。")

    print(f"\n{'=' * 80}")
    print("审计完成。未修改 questions.json。")
    print(f"{'=' * 80}")


if __name__ == "__main__":
    main()
