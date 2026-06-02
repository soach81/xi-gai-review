#!/usr/bin/env python3
"""
从多个源 docx 文件中提取高质量"速览"题目，输出 slides.json。
用于暗色全屏演示模式的快速复习。
"""

import json
import re
import sys
import hashlib
from pathlib import Path
from docx import Document

# ── 配置 ──────────────────────────────────────────────────────────────────
SOURCES = [
    {
        "path": r"C:\Users\99114\Desktop\习概\04_重点预测\重点常考总结.docx",
        "name": "重点常考总结",
        "priority": 1,
    },
    {
        "path": r"C:\Users\99114\Desktop\习概\04_重点预测\重点.docx",
        "name": "重点",
        "priority": 2,
    },
    {
        "path": r"C:\Users\99114\Desktop\习概\04_重点预测\单选知识点.docx",
        "name": "单选知识点",
        "priority": 3,
    },
    {
        "path": r"C:\Users\99114\Desktop\习概\06_教材参考\习思想 知识小点 2021(OCR).docx",
        "name": "知识小点",
        "priority": 4,
    },
    {
        "path": r"C:\Users\99114\Desktop\习概\03_题库_辅助\课后题详解.docx",
        "name": "课后题详解",
        "priority": 5,
    },
]

OUTPUT_PATH = r"C:\Users\99114\Desktop\习概复习\data\slides.json"

# 章节映射
CHAPTER_MAP = {
    "导论": ("导论", "guide"),
    "绪论": ("导论", "guide"),
    "第一章": ("坚持党的领导", "ch01"),
    "第二章": ("以中国式现代化全面推进中华民族伟大复兴", "ch02"),
    "第三章": ("坚持党的全面领导", "ch01"),
    "第四章": ("坚持以人民为中心", "ch04"),
    "第五章": ("全面深化改革开放", "ch05"),
    "第六章": ("推动高质量发展", "ch06"),
    "第七章": ("教育科技人才", "ch07"),
    "第八章": ("发展全过程人民民主", "ch08"),
    "第九章": ("全面依法治国", "ch09"),
    "第十章": ("建设社会主义文化强国", "ch10"),
    "第十一章": ("增进民生福祉", "ch11"),
    "第十二章": ("建设美丽中国", "ch12"),
    "第十三章": ("推进国家安全体系和能力现代化", "ch13"),
    "第十四章": ("实现建军一百年奋斗目标", "ch14"),
    "第十五章": ("坚持一国两制和推进祖国统一", "ch15"),
    "第十六章": ("推动构建人类命运共同体", "ch16"),
    "第十七章": ("全面从严治党", "ch17"),
    "第一讲": ("导论", "guide"),
    "第二讲": ("总任务与中国式现代化", "ch02"),
    "第三讲": ("坚持党的全面领导", "ch01"),
    "第四讲": ("坚持以人民为中心", "ch04"),
    "第五讲": ("全面深化改革开放", "ch05"),
    "第六讲": ("推动高质量发展", "ch06"),
    "第七讲": ("教育科技人才", "ch07"),
    "第八讲": ("发展全过程人民民主", "ch08"),
    "第九讲": ("全面依法治国", "ch09"),
    "第十讲": ("建设社会主义文化强国", "ch10"),
    "第十一讲": ("增进民生福祉", "ch11"),
    "第十二讲": ("建设美丽中国", "ch12"),
    "第十三讲": ("实现建军一百年奋斗目标", "ch14"),
    "第十四讲": ("推进国家安全体系和能力现代化", "ch13"),
    "第十五讲": ("坚持一国两制和推进祖国统一", "ch15"),
    "第十六讲": ("推动构建人类命运共同体", "ch16"),
    "第十七讲": ("全面从严治党", "ch17"),
}

# 题型分类关键词
TYPE_KEYWORDS = {
    "定义": ["什么是", "是什么", "的定义", "的内涵", "的概念", "指什么"],
    "原因": ["为什么", "为何", "原因", "依据"],
    "理解": ["如何理解", "怎样理解", "怎么理解", "怎样认识"],
    "措施": ["如何", "怎样", "怎么做", "路径", "方法"],
    "内容": ["包括哪些", "有哪些", "内容是什么", "包括什么", "包含"],
    "意义": ["意义", "作用", "价值", "重要性"],
    "关系": ["关系", "联系", "区别"],
    "原则": ["原则", "方针", "政策"],
}

# 无意义答案黑名单
BAD_ANSWERS = {
    "不会考具体是啥", "材料分析", "记小点其他编一下", "编一下",
    "看一下就行", "有点重要", "结合其他的答一下",
    "不会考", "编一下", "看一下就行", "有点重要",
    "（不会考具体是啥）", "（记小点其他编一下）",
    "（编一下）", "（看一下就行）", "（有点重要）",
}

# 无意义答案模式
BAD_ANSWER_PATTERNS = [
    r"^第[一二三四五六七八九十\d]+(?:章|讲)\s*$",
    r"^[（(]?\d+[）)]?\s*$",
    r"^(?:必考|重点|了解|掌握)\s*$",
    r"^(?:二十大报告|二十大报告重点)\s*$",
    r"^[（(](?:不会考|编一下|看一下就行|有点重要|材料分析)[）)]?\s*$",
]


def detect_chapter(text):
    """从文本中检测章节信息"""
    # 匹配 "第X章" 或 "第X讲"
    m = re.search(r"第([一二三四五六七八九十]+)(?:章|讲)", text)
    if m:
        key = m.group(0)
        if key in CHAPTER_MAP:
            return CHAPTER_MAP[key]

    # 匹配数字章节 "第8章" 等
    m = re.search(r"第(\d+)章", text)
    if m:
        num = int(m.group(1))
        key_cn = _num_to_cn(num)
        key = f"第{key_cn}章"
        if key in CHAPTER_MAP:
            return CHAPTER_MAP[key]

    return None


def _num_to_cn(n):
    cn = "零一二三四五六七八九十"
    if n <= 10:
        return cn[n]
    if n < 20:
        return f"十{cn[n-10]}" if n > 10 else "十"
    return str(n)


def classify_type(question):
    """根据问题内容分类题型"""
    for qtype, keywords in TYPE_KEYWORDS.items():
        for kw in keywords:
            if kw in question:
                return qtype
    return "知识点"


def shorten_answer(answer, max_len=100):
    """将长答案精简到 max_len 字以内"""
    if len(answer) <= max_len:
        return answer

    # 先尝试按句号分割，取前几句
    sentences = re.split(r"[。；;]", answer)
    result = ""
    for s in sentences:
        s = s.strip()
        if not s:
            continue
        candidate = result + ("。" if result else "") + s
        if len(candidate) <= max_len:
            result = candidate
        else:
            break
    if result and len(result) <= max_len:
        return result + "。"

    # 如果还是太长，直接截断
    return answer[: max_len - 1] + "。"


def shorten_question(question, max_len=30):
    """将长问题精简到 max_len 字以内"""
    if len(question) <= max_len:
        return question

    # 尝试去掉冗余前缀
    prefixes = [
        "请问", "请回答", "请简要回答", "请简述", "请说明",
        "简要回答", "简述", "试述", "论述",
    ]
    for p in prefixes:
        if question.startswith(p):
            question = question[len(p):]
            break

    # 去掉开头的标点
    question = question.lstrip("，,。.、：:")

    if len(question) <= max_len:
        return question

    # 截断
    return question[:max_len]


def is_valid_pair(question, answer):
    """质量检查：判断问答对是否有效"""
    if len(question) < 5 or len(answer) < 5:
        return False

    if question.strip() == answer.strip():
        return False

    # 不是纯数字或页码
    if re.match(r"^[\d\s.]+$", question):
        return False
    if re.match(r"^[\d\s.]+$", answer):
        return False

    # 不是章节标题
    if re.match(r"^第[一二三四五六七八九十\d]+(?:章|讲)$", question):
        return False

    # 不是页码
    if re.match(r"^\d+\s*页$", question):
        return False

    # 不是空内容或纯标点
    if re.match(r"^[，,。.、；;：:\s]+$", question):
        return False
    if re.match(r"^[，,。.、；;：:\s]+$", answer):
        return False

    # 答案不能太碎
    if len(answer) < 6:
        return False

    # 问题不能包含答案（自引用）
    if question in answer and len(question) > len(answer) * 0.8:
        return False

    # 无意义答案
    if answer.strip() in BAD_ANSWERS:
        return False

    # 无意义答案模式
    for pattern in BAD_ANSWER_PATTERNS:
        if re.match(pattern, answer.strip()):
            return False

    # 答案不能以"答"开头（说明提取错误）
    if answer.strip().startswith("答") and len(answer.strip()) < 10:
        return False

    # 答案不能是章节标题
    if re.match(r"^第[一二三四五六七八九十\d]+(?:章|讲)\s*$", answer.strip()):
        return False

    # 答案不能包含"讲"作为主要内容（如"第十    讲"）
    if re.match(r"^第[一二三四五六七八九十\d]+\s*讲\s*$", answer.strip()):
        return False

    # 问题包含明显非问题内容
    bad_q_patterns = ["不会考", "编一下", "看一下就行"]
    for bp in bad_q_patterns:
        if bp in question:
            return False

    # 问题中包含括号标注（如"（掌握）"）→ 清理
    # 但不在问题中间

    return True


def text_hash(text):
    """生成文本指纹用于去重"""
    cleaned = re.sub(r"\s+", "", text)
    return hashlib.md5(cleaned.encode("utf-8")).hexdigest()


def clean_question(q):
    """清理问题文本：去掉前缀编号、括号标注等"""
    # 去掉 "X、（重点）" 前缀
    q = re.sub(r"^\d+\s*[、.]\s*(?:（[了解掌握重点选简论多]{1,4}）\s*)?", "", q)
    # 去掉 "（选）" 等前缀
    q = re.sub(r"^[（(]([了解掌握重点选简论多]{1,4})[）)]\s*", "", q)
    # 去掉开头的 "一 、" 等
    q = re.sub(r"^[一二三四五六七八九十]+\s*[、.]\s*(?:（[了解掌握重点]{2,4}）\s*)?", "", q)
    # 去掉问题末尾的括号标注（如"（掌握）"）
    q = re.sub(r"[（(](?:掌握|了解|重点|选|简|论|多选)[）)]\s*$", "", q)
    # 去掉问题末尾的多余冒号和空格
    q = q.rstrip("：: ")
    return q.strip()


def extract_from_keystrokes(paras, name, current_chapter):
    """从重点常考总结.docx 格式提取（带编号+答：格式）"""
    items = []
    for idx, para in enumerate(paras):
        if not para:
            continue

        ch = detect_chapter(para)
        if ch:
            current_chapter = ch
            continue

        # ── 模式 A: "X、（重点）问题？" + 下一行 "答：答案" ──
        m = re.match(
            r"^(?:[一二三四五六七八九十]+\s*[、.]\s*|\d+\s*[、.]\s*)"
            r"(?:（[了解掌握重点]{2,4}）\s*)?"
            r"(.{5,50}[？?])\s*$",
            para,
        )
        if m:
            q = clean_question(m.group(1).strip())
            # 检查下一行是否有答案
            for fwd in range(1, min(5, len(paras) - idx)):
                nxt = paras[idx + fwd].strip()
                if not nxt:
                    continue
                # 跳过章节标记
                if detect_chapter(nxt):
                    break
                # 跳过另一个问题
                if re.match(r"^[一二三四五六七八九十\d]+\s*[、.]", nxt):
                    break
                # 找到答案行
                ans_m = re.match(r"^(?:答[：:]|答案[：:])?\s*(.{5,300})", nxt)
                if ans_m:
                    a = ans_m.group(1).strip()
                    # 如果答案行以编号列表开头，尝试收集更多
                    if re.match(r"^\d+\s*[、.]", a):
                        # 收集列表项
                        parts = [a]
                        for fwd2 in range(fwd + 1, min(fwd + 10, len(paras))):
                            nxt2 = paras[idx + fwd2].strip()
                            if not nxt2:
                                continue
                            if re.match(r"^[一二三四五六七八九十\d]+\s*[、.]", nxt2):
                                break
                            if re.match(r"^(?:答[：:])", nxt2):
                                break
                            if re.match(r"^\d+\s*[、.]", nxt2) or re.match(r"^[（(]\d", nxt2):
                                parts.append(nxt2)
                            else:
                                break
                        a = "；".join(parts)
                    a = shorten_answer(a)
                    if is_valid_pair(q, a):
                        items.append({
                            "question": q,
                            "answer": a,
                            "chapter": current_chapter[0],
                            "chapterId": current_chapter[1],
                            "source": name,
                        })
                break

        # ── 模式 B: "问题？答案" 在同一行 ──
        m = re.match(r"^(.{5,50}[？?])\s*(.{5,200})", para)
        if m:
            q = clean_question(m.group(1).strip())
            a = shorten_answer(m.group(2).strip())
            if is_valid_pair(q, a):
                # 避免重复（已经通过模式A提取）
                already = any(
                    text_hash(q) == text_hash(it["question"])
                    for it in items[-5:]
                )
                if not already:
                    items.append({
                        "question": q,
                        "answer": a,
                        "chapter": current_chapter[0],
                        "chapterId": current_chapter[1],
                        "source": name,
                    })
                continue

    return items


def extract_from_lesson_qa(paras, name, current_chapter):
    """从课后题详解.docx 格式提取（编号问题 + 【答案】+ 回答段落）"""
    items = []
    for idx, para in enumerate(paras):
        if not para:
            continue

        ch = detect_chapter(para)
        if ch:
            current_chapter = ch
            continue

        # 匹配问题行：以数字编号开头，以？结尾
        m = re.match(r"^\d+\.\s*(.{5,60}[？?])\s*$", para)
        if not m:
            continue

        q = clean_question(m.group(1).strip())

        # 向后找【答案】标记
        answer_start = None
        for fwd in range(1, min(5, len(paras) - idx)):
            nxt = paras[idx + fwd].strip()
            if "【答案】" in nxt or "答案" == nxt:
                answer_start = idx + fwd + 1
                break
            if re.match(r"^\d+\.\s*", nxt):
                break  # 下一个问题

        if answer_start is None:
            continue

        # 收集答案段落（直到下一个问题或章节）
        answer_parts = []
        for fwd in range(answer_start, min(answer_start + 15, len(paras))):
            nxt = paras[fwd].strip()
            if not nxt:
                continue
            # 停止条件
            if re.match(r"^\d+\.\s*[一二三四五六七八九十]*\d*\.\s*", nxt):
                break
            if re.match(r"^\d+\.\s*如何", nxt) or re.match(r"^\d+\.\s*怎样", nxt):
                break
            if re.match(r"^\d+\.\s*为什么", nxt) or re.match(r"^\d+\.\s*如何理解", nxt):
                break
            if detect_chapter(nxt):
                break
            # 去掉（1）（2）等前缀，保留内容
            clean = re.sub(r"^[（(]\d+[）)]\s*", "", nxt)
            clean = re.sub(r"^①②③④⑤⑥⑦⑧⑨⑩", "", clean)
            if clean and len(clean) > 3:
                answer_parts.append(clean)

        if not answer_parts:
            continue

        a = "。".join(answer_parts)
        a = shorten_answer(a)

        if is_valid_pair(q, a):
            items.append({
                "question": q,
                "answer": a,
                "chapter": current_chapter[0],
                "chapterId": current_chapter[1],
                "source": name,
            })

    return items


def extract_from_outline(paras, name, current_chapter):
    """从重点.docx 格式提取（（选/简/论）+ 知识点内容）"""
    items = []
    for idx, para in enumerate(paras):
        if not para:
            continue

        ch = detect_chapter(para)
        if ch:
            current_chapter = ch
            continue

        # ── 模式1: "（选/简/论）问题内容" ──
        m = re.match(
            r"^[（(](选|简|论|多选|选或简|选或多)(?:选)?[）)]\s*(.{5,80})",
            para,
        )
        if m:
            content = m.group(2).strip()
            # 如果有问号，拆分为问答
            if "？" in content or "?" in content:
                parts = re.split(r"[：:？?]", content, maxsplit=1)
                if len(parts) == 2:
                    q = parts[0].strip()
                    if not q.endswith("？"):
                        q += "？"
                    a = parts[1].strip().rstrip("。")
                    a = shorten_answer(a)
                    q = clean_question(q)
                    if is_valid_pair(q, a):
                        items.append({
                            "question": q,
                            "answer": a,
                            "chapter": current_chapter[0],
                            "chapterId": current_chapter[1],
                            "source": name,
                        })
                        continue
                # 没有冒号分隔，但有问号 → 向后找答案
                q = content.strip()
                if not q.endswith("？"):
                    q += "？"
                q = clean_question(q)
                # 向后找答案
                for fwd in range(1, min(5, len(paras) - idx)):
                    nxt = paras[idx + fwd].strip()
                    if not nxt:
                        continue
                    if re.match(r"^[（(](选|简|论|多选)", nxt):
                        break
                    if re.match(r"^[一二三四五六七八九十\d]+\s*[、.]", nxt):
                        break
                    a = shorten_answer(nxt)
                    if is_valid_pair(q, a):
                        items.append({
                            "question": q,
                            "answer": a,
                            "chapter": current_chapter[0],
                            "chapterId": current_chapter[1],
                            "source": name,
                        })
                    break
                continue

            # 没有问号 → 如果内容是"XXX是什么"类，转为问答
            if "是什么" in content or "有哪些" in content or "包括" in content:
                q = content.strip()
                if not q.endswith("？"):
                    q += "？"
                q = clean_question(q)
                for fwd in range(1, min(5, len(paras) - idx)):
                    nxt = paras[idx + fwd].strip()
                    if not nxt:
                        continue
                    if re.match(r"^[（(](选|简|论|多选)", nxt):
                        break
                    if re.match(r"^[一二三四五六七八九十\d]+\s*[、.]", nxt):
                        break
                    a = shorten_answer(nxt)
                    if is_valid_pair(q, a):
                        items.append({
                            "question": q,
                            "answer": a,
                            "chapter": current_chapter[0],
                            "chapterId": current_chapter[1],
                            "source": name,
                        })
                    break
                continue

            # 纯知识点标题 → 转为问答（向后取答案）
            q = content.strip().rstrip("。")
            if len(q) > 5 and len(q) < 50:
                # 检查是否包含"是"可以拆分
                if "是" in q:
                    parts = q.split("是", 1)
                    if len(parts) == 2 and len(parts[0]) > 3 and len(parts[1]) > 3:
                        q = parts[0].strip() + "是什么？"
                        a = shorten_answer(parts[1].strip().rstrip("。"))
                        q = clean_question(q)
                        if is_valid_pair(q, a):
                            items.append({
                                "question": q,
                                "answer": a,
                                "chapter": current_chapter[0],
                                "chapterId": current_chapter[1],
                                "source": name,
                            })
                            continue
                # 向后取答案
                for fwd in range(1, min(5, len(paras) - idx)):
                    nxt = paras[idx + fwd].strip()
                    if not nxt:
                        continue
                    if re.match(r"^[（(](选|简|论|多选)", nxt):
                        break
                    if re.match(r"^[一二三四五六七八九十\d]+\s*[、.]", nxt):
                        break
                    a = shorten_answer(nxt)
                    if is_valid_pair(q + "？", a):
                        items.append({
                            "question": q + "？",
                            "answer": a,
                            "chapter": current_chapter[0],
                            "chapterId": current_chapter[1],
                            "source": name,
                        })
                    break

        # ── 模式2: "X、（重点）问题？" + 下一行答案 ──
        m = re.match(
            r"^(?:[一二三四五六七八九十]+\s*[、.]\s*|\d+\s*[、.]\s*)"
            r"(?:（[了解掌握重点选简论]{1,4}）\s*)?"
            r"(.{5,60}[？?])\s*$",
            para,
        )
        if m:
            q = clean_question(m.group(1).strip())
            # 向后找答案
            for fwd in range(1, min(5, len(paras) - idx)):
                nxt = paras[idx + fwd].strip()
                if not nxt:
                    continue
                if re.match(r"^[一二三四五六七八九十\d]+\s*[、.]", nxt):
                    break
                if re.match(r"^[（(](选|简|论|多选)", nxt):
                    break
                ans_m = re.match(r"^(?:答[：:]|答案[：:])?\s*(.{5,300})", nxt)
                if ans_m:
                    a = ans_m.group(1).strip()
                    a = shorten_answer(a)
                    if is_valid_pair(q, a):
                        items.append({
                            "question": q,
                            "answer": a,
                            "chapter": current_chapter[0],
                            "chapterId": current_chapter[1],
                            "source": name,
                        })
                    break

        # ── 模式3: "问题？答案" 在同一行 ──
        m = re.match(r"^(.{5,50}[？?])\s*(.{5,200})", para)
        if m and not any(
            text_hash(clean_question(m.group(1).strip())) == text_hash(it["question"])
            for it in items[-5:]
        ):
            q = clean_question(m.group(1).strip())
            a = shorten_answer(m.group(2).strip())
            if is_valid_pair(q, a):
                items.append({
                    "question": q,
                    "answer": a,
                    "chapter": current_chapter[0],
                    "chapterId": current_chapter[1],
                    "source": name,
                })

        # ── 模式4: "X. 知识点内容是YYY" → 转为问答 ──
        m = re.match(
            r"^\d+\.\s*(.{5,50}(?:是|为|包括|包含|指).{5,100})$",
            para,
        )
        if m:
            content = m.group(1).strip()
            split_m = re.match(
                r"^(.{3,30})(?:是|为|包括|包含|指)(.{3,100})$", content
            )
            if split_m:
                topic = split_m.group(1).strip()
                definition = split_m.group(2).strip()
                q = f"{topic}是什么？"
                a = shorten_answer(definition)
                q = clean_question(q)
                if is_valid_pair(q, a):
                    items.append({
                        "question": q,
                        "answer": a,
                        "chapter": current_chapter[0],
                        "chapterId": current_chapter[1],
                        "source": name,
                    })

    return items


def extract_from_knowledge_points(paras, name, current_chapter):
    """从知识小点.docx 格式提取（编号知识点 + 简短答案）"""
    items = []
    for idx, para in enumerate(paras):
        if not para:
            continue

        ch = detect_chapter(para)
        if ch:
            current_chapter = ch
            continue

        # ── 模式: "X. YYY" 知识点，可以转为问答 ──
        m = re.match(r"^\d+\.\s*(.{8,120})$", para)
        if m:
            content = m.group(1).strip()

            # "X. YYY是ZZZ" 格式
            split_m = re.match(
                r"^(.{3,40})(?:是|为|包括|包含|指|就要|就是要|就是)(.{3,120})$",
                content,
            )
            if split_m:
                topic = split_m.group(1).strip()
                definition = split_m.group(2).strip()
                q = f"{topic}是什么？"
                a = shorten_answer(definition)
                q = clean_question(q)
                if is_valid_pair(q, a):
                    items.append({
                        "question": q,
                        "answer": a,
                        "chapter": current_chapter[0],
                        "chapterId": current_chapter[1],
                        "source": name,
                    })
                    continue

            # "X. XXXX，YYYY" 两段式知识点
            if "，" in content and len(content) > 15:
                parts = content.split("，", 1)
                if len(parts) == 2 and len(parts[0]) > 5 and len(parts[1]) > 5:
                    q = parts[0].strip() + "是什么？"
                    a = shorten_answer(parts[1].strip())
                    q = clean_question(q)
                    if is_valid_pair(q, a):
                        items.append({
                            "question": q,
                            "answer": a,
                            "chapter": current_chapter[0],
                            "chapterId": current_chapter[1],
                            "source": name,
                        })
                        continue

        # ── 模式: "X. YYY。" 简短知识点 ──
        m = re.match(r"^\d+\.\s*(.{8,100})[。]?\s*$", para)
        if m:
            content = m.group(1).strip()
            split_m = re.match(
                r"^(.{3,40})(?:是|为|包括|包含|指|就要|就是要|就是)(.{3,100})$",
                content,
            )
            if split_m:
                topic = split_m.group(1).strip()
                definition = split_m.group(2).strip()
                q = f"{topic}是什么？"
                a = shorten_answer(definition)
                q = clean_question(q)
                if is_valid_pair(q, a):
                    items.append({
                        "question": q,
                        "answer": a,
                        "chapter": current_chapter[0],
                        "chapterId": current_chapter[1],
                        "source": name,
                    })

    return items


def extract_from_single_choice(paras, name, current_chapter):
    """从单选知识点.docx 格式提取（无编号的陈述性知识点）"""
    items = []
    for idx, para in enumerate(paras):
        if not para:
            continue

        ch = detect_chapter(para)
        if ch:
            current_chapter = ch
            continue

        # 跳过太短的行
        if len(para) < 10:
            continue

        # ── 模式1: "XXX是/为/包括/指 YYY" 陈述句 ──
        m = re.match(
            r"^(.{5,40})(?:是|为|包括|包含|指)(.{5,120})$",
            para,
        )
        if m:
            topic = m.group(1).strip()
            definition = m.group(2).strip().rstrip("。")
            # 确保topic不是太长（不像一句话）
            if len(topic) <= 35:
                q = f"{topic}是什么？"
                a = shorten_answer(definition)
                q = clean_question(q)
                if is_valid_pair(q, a):
                    items.append({
                        "question": q,
                        "answer": a,
                        "chapter": current_chapter[0],
                        "chapterId": current_chapter[1],
                        "source": name,
                    })
                    continue

        # ── 模式2: "XXX提出/指出/强调 YYY" → 转为问答 ──
        m = re.match(
            r"^(.{5,30})(?:提出|指出|强调|明确|要求|决定|通过|将|把)(.{5,120})$",
            para,
        )
        if m:
            topic = m.group(1).strip()
            content = m.group(2).strip().rstrip("。")
            # 转为问题
            q = f"{topic}提出了什么？"
            a = shorten_answer(content)
            q = clean_question(q)
            if is_valid_pair(q, a):
                items.append({
                    "question": q,
                    "answer": a,
                    "chapter": current_chapter[0],
                    "chapterId": current_chapter[1],
                    "source": name,
                })
                continue

        # ── 模式3: 包含"是"的较长陈述句 ──
        if "是" in para and len(para) > 15 and len(para) < 150:
            # 尝试在"是"处拆分
            parts = para.split("是", 1)
            if len(parts) == 2 and len(parts[0]) > 5 and len(parts[1]) > 5:
                topic = parts[0].strip().rstrip("，,")
                definition = parts[1].strip().rstrip("。")
                if len(topic) <= 35 and len(definition) > 5:
                    q = f"{topic}是什么？"
                    a = shorten_answer(definition)
                    q = clean_question(q)
                    if is_valid_pair(q, a):
                        # 避免重复
                        already = any(
                            text_hash(q) == text_hash(it["question"])
                            for it in items[-10:]
                        )
                        if not already:
                            items.append({
                                "question": q,
                                "answer": a,
                                "chapter": current_chapter[0],
                                "chapterId": current_chapter[1],
                                "source": name,
                            })

    return items


def extract_from_docx(source):
    """从单个 docx 文件中提取问答对"""
    path = source["path"]
    name = source["name"]

    try:
        doc = Document(path)
    except Exception as e:
        print(f"  [WARN] Cannot open {name}: {e}", file=sys.stderr)
        return []

    paragraphs = [p.text.strip() for p in doc.paragraphs]

    # 根据文件类型使用不同的提取策略
    if "重点常考总结" in name:
        return extract_from_keystrokes(paragraphs, name, ("导论", "guide"))
    elif "课后题详解" in name:
        return extract_from_lesson_qa(paragraphs, name, ("导论", "guide"))
    elif "重点" in name and "总结" not in name:
        return extract_from_outline(paragraphs, name, ("导论", "guide"))
    elif "知识小点" in name:
        return extract_from_knowledge_points(paragraphs, name, ("导论", "guide"))
    elif "单选知识点" in name:
        return extract_from_single_choice(paragraphs, name, ("导论", "guide"))
    else:
        return extract_from_outline(paragraphs, name, ("导论", "guide"))


def deduplicate(items):
    """去重：相同或高度相似的问题只保留一条"""
    seen_hashes = set()
    unique = []

    for item in items:
        h = text_hash(item["question"])
        if h in seen_hashes:
            continue
        seen_hashes.add(h)
        unique.append(item)

    # 第二轮：检查问题相似度（前12字相同视为重复）
    final = []
    seen_prefixes = set()
    for item in unique:
        prefix = re.sub(r"\s+", "", item["question"])[:12]
        if prefix in seen_prefixes:
            continue
        seen_prefixes.add(prefix)
        final.append(item)

    return final


def main():
    all_items = []

    for source in SOURCES:
        print(f"Processing: {source['name']} ...", file=sys.stderr)
        items = extract_from_docx(source)
        print(f"  Extracted {len(items)} raw Q&A pairs", file=sys.stderr)
        all_items.extend(items)

    print(f"\nTotal raw: {len(all_items)}", file=sys.stderr)

    # 质量过滤
    valid_items = [
        item for item in all_items if is_valid_pair(item["question"], item["answer"])
    ]
    print(f"After quality filter: {len(valid_items)}", file=sys.stderr)

    # 去重
    deduped = deduplicate(valid_items)
    print(f"After dedup: {len(deduped)}", file=sys.stderr)

    # 最终处理
    final_items = []
    for idx, item in enumerate(deduped):
        q = shorten_question(item["question"])
        a = shorten_answer(item["answer"])
        if not is_valid_pair(q, a):
            continue
        final_items.append({
            "id": f"slide_{idx+1:03d}",
            "question": q,
            "answer": a,
            "chapter": item["chapter"],
            "chapterId": item["chapterId"],
            "type": classify_type(q),
            "source": item["source"],
        })

    print(f"Final output: {len(final_items)}", file=sys.stderr)

    # 构建输出
    output = {
        "meta": {
            "subject": "习近平新时代中国特色社会主义思想概论",
            "version": "1.0",
            "totalSlides": len(final_items),
        },
        "items": final_items,
    }

    # 写入文件
    Path(OUTPUT_PATH).parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\nWritten to: {OUTPUT_PATH}", file=sys.stderr)

    # 统计
    chapters = {}
    sources_stat = {}
    types = {}
    for item in final_items:
        chapters[item["chapter"]] = chapters.get(item["chapter"], 0) + 1
        sources_stat[item["source"]] = sources_stat.get(item["source"], 0) + 1
        types[item["type"]] = types.get(item["type"], 0) + 1

    print("\n=== By Chapter ===", file=sys.stderr)
    for ch, cnt in sorted(chapters.items(), key=lambda x: -x[1]):
        print(f"  {ch}: {cnt}", file=sys.stderr)

    print("\n=== By Source ===", file=sys.stderr)
    for src, cnt in sorted(sources_stat.items(), key=lambda x: -x[1]):
        print(f"  {src}: {cnt}", file=sys.stderr)

    print("\n=== By Type ===", file=sys.stderr)
    for tp, cnt in sorted(types.items(), key=lambda x: -x[1]):
        print(f"  {tp}: {cnt}", file=sys.stderr)


if __name__ == "__main__":
    main()
