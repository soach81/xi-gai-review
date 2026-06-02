#!/usr/bin/env python3
"""
extract_flashcards_v3.py - 重新提取和优化卡片数据，重点增加挖空知识点类型
"""

import json
import re
import sys
import os
from pathlib import Path
from typing import List, Dict, Tuple, Optional
import docx

# 章节映射
CHAPTER_MAP = {
    "guide": "导论",
    "ch01": "第一章",
    "ch02": "第二章",
    "ch03": "第三章",
    "ch04": "第四章",
    "ch05": "第五章",
    "ch06": "第六章",
    "ch07": "第七章",
    "ch08": "第八章",
    "ch09": "第九章",
    "ch10": "第十章",
    "ch11": "第十一章",
    "ch12": "第十二章",
    "ch13": "第十三章",
    "ch14": "第十四章",
    "ch15": "第十五章",
    "ch16": "第十六章",
    "ch17": "第十七章"
}


def read_docx(docx_path: str) -> List[str]:
    """读取docx文件，返回段落文本列表"""
    if not os.path.exists(docx_path):
        print(f"Warning: {docx_path} not found, skipping...")
        return []

    try:
        doc = docx.Document(docx_path)
        paragraphs = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                paragraphs.append(text)
        return paragraphs
    except Exception as e:
        print(f"Error reading {docx_path}: {e}")
        return []


def read_questions_json(json_path: str) -> List[Dict]:
    """读取questions.json，返回选择题列表"""
    if not os.path.exists(json_path):
        print(f"Warning: {json_path} not found, skipping...")
        return []

    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        questions = []
        for chapter in data.get('chapters', []):
            chapter_id = chapter.get('id', '')
            for q in chapter.get('questions', []):
                if q.get('type') == 'single':
                    q['chapter_id'] = chapter_id
                    questions.append(q)
        return questions
    except Exception as e:
        print(f"Error reading {json_path}: {e}")
        return []


def extract_blanks_from_question(question: Dict) -> Optional[Dict]:
    """从选择题提取挖空知识点"""
    q_text = question.get('question', '')
    options = question.get('options', [])
    answer_indices = question.get('answer', [])
    chapter_id = question.get('chapter_id', '')

    if not options or not answer_indices:
        return None

    # 获取正确答案文本
    correct_answers = []
    for idx in answer_indices:
        if 0 <= idx < len(options):
            answer_text = options[idx].strip()
            # 清理答案文本中的多余内容
            answer_text = re.sub(r'选择一项：', '', answer_text).strip()
            answer_text = re.sub(r'选择多项：', '', answer_text).strip()
            correct_answers.append(answer_text)

    if not correct_answers:
        return None

    # 将题干转换为挖空形式
    front = q_text

    # 清理题干中的多余标记
    front = re.sub(r'\( \)', '____', front)  # 将括号转换为挖空
    front = re.sub(r'（ ）', '____', front)  # 将中文括号转换为挖空
    front = re.sub(r'\(\)', '____', front)    # 将无空格括号转换为挖空

    # 如果题干已经包含挖空标记，直接使用
    if '____' in front:
        pass
    else:
        # 尝试识别题干中的核心概念并挖空
        patterns = [
            r'(.+?)是\(',
            r'(.+?)是（',
            r'(.+?)的（',
            r'(.+?)的\(',
            r'(.+?)是',
            r'(.+?)的',
        ]

        found = False
        for pattern in patterns:
            match = re.search(pattern, front)
            if match:
                front = match.group(1) + '是____'
                found = True
                break

        if not found:
            # 如果没有找到典型模式，将最后部分挖空
            front = front.rstrip('。，、；') + '是____'

    # 确保front不超过40字
    if len(front) > 40:
        front = front[:37] + '...'

    # 组合back答案
    back = '；'.join(correct_answers)
    if len(back) > 50:
        back = back[:47] + '...'

    # 确定章节
    chapter_name = CHAPTER_MAP.get(chapter_id, "导论")

    return {
        'front': front,
        'back': back,
        'chapter': chapter_name,
        'chapterId': chapter_id,
        'type': '挖空',
        'source': '选择题'
    }


def extract_from_danxuan(lines: List[str]) -> List[Dict]:
    """从单选知识点提取卡片"""
    cards = []

    for line in lines:
        # 跳过空行和元数据行
        if not line or line.startswith('[') and '习概：' in line:
            continue

        # 清理行号标记
        clean_line = re.sub(r'^\[\d+\]\s*\[.*?\]\s*', '', line).strip()

        if not clean_line or len(clean_line) < 10:
            continue

        # 尝试提取挖空知识点
        front = None
        back = None
        card_type = '挖空'
        chapter = '导论'
        chapter_id = 'guide'

        # 模式1：包含"是"的知识点
        if '是' in clean_line:
            parts = clean_line.split('是', 1)
            if len(parts) == 2 and len(parts[1]) > 5:
                front = parts[0].strip() + '是____'
                back = parts[1].strip().rstrip('。，、；')
                card_type = '挖空'

        # 模式2：包含"包括"的知识点
        elif '包括' in clean_line:
            parts = clean_line.split('包括', 1)
            if len(parts) == 2:
                front = parts[0].strip() + '包括哪些？'
                back = parts[1].strip()
                card_type = '列举'

        # 模式3：包含"目标"的知识点
        elif '目标' in clean_line:
            front = clean_line.split('目标')[0].strip() + '的目标是____'
            back = clean_line.split('目标')[1].strip().rstrip('。，、；')
            card_type = '挖空'

        # 模式4：包含"原则"的知识点
        elif '原则' in clean_line:
            front = clean_line.split('原则')[0].strip() + '的原则是____'
            back = clean_line.split('原则')[1].strip().rstrip('。，、；')
            card_type = '挖空'

        # 如果无法提取，跳过
        if not front or not back:
            continue

        # 确保front不超过40字
        if len(front) > 40:
            front = front[:37] + '...'

        # 确保back不超过50字
        if len(back) > 50:
            back = back[:47] + '...'

        cards.append({
            'front': front,
            'back': back,
            'chapter': chapter,
            'chapterId': chapter_id,
            'type': card_type,
            'source': '单选知识点'
        })

    return cards


def extract_from_zhongdian(lines: List[str]) -> List[Dict]:
    """从重点总结提取卡片"""
    cards = []
    current_chapter = '导论'
    current_chapter_id = 'guide'

    for line in lines:
        # 清理行号标记
        clean_line = re.sub(r'^\[\d+\]\s*\[.*?\]\s*', '', line).strip()

        if not clean_line:
            continue

        # 识别章节标题
        if re.match(r'^第[一二三四五六七八九十]+章', clean_line):
            chapter_match = re.match(r'^(第[一二三四五六七八九十]+章)', clean_line)
            if chapter_match:
                current_chapter = chapter_match.group(1)
                # 查找对应的chapter_id
                for cid, cname in CHAPTER_MAP.items():
                    if cname == current_chapter:
                        current_chapter_id = cid
                        break
            continue

        if clean_line == '导论':
            current_chapter = '导论'
            current_chapter_id = 'guide'
            continue

        # 提取知识点
        # 模式1：包含"是什么"的知识点
        if '是什么' in clean_line:
            front = clean_line.replace('是什么', '是____')
            back = front.split('是____')[0].strip() + '的概念'
            card_type = '挖空'

        # 模式2：包含"包括"的知识点
        elif '包括' in clean_line:
            front = clean_line.split('包括')[0].strip() + '包括哪些？'
            back = clean_line.split('包括')[1].strip()
            card_type = '列举'

        # 模式3：包含"内容"的知识点
        elif '内容' in clean_line:
            front = clean_line.split('内容')[0].strip() + '的内容是什么？'
            back = clean_line.split('内容')[1].strip()
            card_type = '列举'

        # 模式4：包含"特征"的知识点
        elif '特征' in clean_line:
            front = clean_line.split('特征')[0].strip() + '的特征是____'
            back = clean_line.split('特征')[1].strip()
            card_type = '挖空'

        # 模式5：包含"本质"的知识点
        elif '本质' in clean_line:
            front = clean_line.split('本质')[0].strip() + '的本质是____'
            back = clean_line.split('本质')[1].strip()
            card_type = '挖空'

        else:
            continue

        # 确保front不超过40字
        if len(front) > 40:
            front = front[:37] + '...'

        # 确保back不超过50字
        if len(back) > 50:
            back = back[:47] + '...'

        cards.append({
            'front': front,
            'back': back,
            'chapter': current_chapter,
            'chapterId': current_chapter_id,
            'type': card_type,
            'source': '重点总结'
        })

    return cards


def extract_from_zhishi(lines: List[str]) -> List[Dict]:
    """从知识小点提取卡片"""
    cards = []
    current_chapter = '导论'
    current_chapter_id = 'guide'

    for line in lines:
        # 清理行号标记
        clean_line = re.sub(r'^\[\d+\]\s*\[.*?\]\s*', '', line).strip()

        if not clean_line:
            continue

        # 识别章节标题
        if re.match(r'^第八章', clean_line):
            current_chapter = '第八章'
            current_chapter_id = 'ch08'
            continue
        elif re.match(r'^第九章', clean_line):
            current_chapter = '第九章'
            current_chapter_id = 'ch09'
            continue
        elif re.match(r'^第十章', clean_line):
            current_chapter = '第十章'
            current_chapter_id = 'ch10'
            continue

        # 提取知识点
        # 模式1：包含"是"的知识点
        if '是' in clean_line and len(clean_line) > 20:
            parts = clean_line.split('是', 1)
            if len(parts) == 2 and len(parts[1]) > 5:
                front = parts[0].strip() + '是____'
                back = parts[1].strip().rstrip('。，、；')
                card_type = '挖空'

                # 确保front不超过40字
                if len(front) > 40:
                    front = front[:37] + '...'

                # 确保back不超过50字
                if len(back) > 50:
                    back = back[:47] + '...'

                cards.append({
                    'front': front,
                    'back': back,
                    'chapter': current_chapter,
                    'chapterId': current_chapter_id,
                    'type': card_type,
                    'source': '知识小点'
                })

    return cards


def deduplicate_cards(cards: List[Dict]) -> List[Dict]:
    """去重：相同front只保留一条"""
    seen_fronts = set()
    unique_cards = []

    for card in cards:
        front = card['front']
        if front not in seen_fronts:
            seen_fronts.add(front)
            unique_cards.append(card)

    return unique_cards


def validate_card(card: Dict) -> bool:
    """验证卡片质量"""
    front = card.get('front', '')
    back = card.get('back', '')

    # 检查front长度
    if len(front) > 40:
        return False

    # 检查back长度
    if len(back) > 50:
        return False

    # 检查front和back内容是否高度不同
    if front == back:
        return False

    # 检查是否纯标题或页码
    if re.match(r'^第[一二三四五六七八九十]+章$', front):
        return False
    if re.match(r'^\d+$', front):
        return False

    # 检查back是否包含多余标记
    if '选择一项' in back or '选择多项' in back:
        return False

    # 检查front语法问题
    if '主要有是' in front or '是是' in front:
        return False

    # 检查front是否以"是____"结尾但内容太短
    if front.endswith('是____') and len(front) < 15:
        return False

    return True


def filter_cards_by_quality(cards: List[Dict]) -> List[Dict]:
    """按质量筛选卡片"""
    # 先验证所有卡片
    valid_cards = [card for card in cards if validate_card(card)]

    # 按类型统计
    type_counts = {}
    for card in valid_cards:
        card_type = card.get('type', '未知')
        type_counts[card_type] = type_counts.get(card_type, 0) + 1

    # 如果挖空卡片超过40%，保留所有；否则调整
    total = len(valid_cards)
    blank_count = type_counts.get('挖空', 0)

    if total > 0 and blank_count / total >= 0.4:
        return valid_cards

    # 否则，优先保留挖空卡片，然后补充其他类型
    blank_cards = [card for card in valid_cards if card.get('type') == '挖空']
    other_cards = [card for card in valid_cards if card.get('type') != '挖空']

    # 计算需要补充的数量
    needed = max(150, total)  # 至少150张
    other_needed = needed - len(blank_cards)

    # 按章节均匀补充其他类型
    if other_needed > 0 and other_cards:
        # 按章节分组
        chapter_cards = {}
        for card in other_cards:
            chapter = card.get('chapter', '未知')
            if chapter not in chapter_cards:
                chapter_cards[chapter] = []
            chapter_cards[chapter].append(card)

        # 从每个章节取一些
        selected_other = []
        for chapter, chapter_cards_list in chapter_cards.items():
            take = min(len(chapter_cards_list), other_needed // len(chapter_cards))
            selected_other.extend(chapter_cards_list[:take])

        return blank_cards + selected_other

    return valid_cards


def main():
    """主函数"""
    # 文件路径
    base_dir = Path(r"C:\Users\99114\Desktop\习概")
    output_dir = Path(r"C:\Users\99114\Desktop\习概复习\data")

    docx_files = {
        'danxuan': base_dir / "04_重点预测" / "单选知识点.docx",
        'zhongdian': base_dir / "04_重点预测" / "重点常考总结.docx",
        'zhishi': base_dir / "06_教材参考" / "习思想 知识小点 2021(OCR).docx"
    }

    json_files = {
        'questions': output_dir / "questions.json"
    }

    all_cards = []

    # 1. 从questions.json提取选择题知识点（挖空形式）
    print("从questions.json提取选择题知识点...")
    questions = read_questions_json(str(json_files['questions']))
    print(f"读取到 {len(questions)} 道选择题")

    for q in questions:
        card = extract_blanks_from_question(q)
        if card:
            all_cards.append(card)

    print(f"从选择题提取了 {len([c for c in all_cards if c['source'] == '选择题'])} 张卡片")

    # 2. 从temp文件提取知识点
    # 读取temp文件
    temp_dir = Path(r"C:\Users\99114\Desktop\习概复习")

    # 读取temp_danxuan.txt
    danxuan_path = temp_dir / "temp_danxuan.txt"
    if danxuan_path.exists():
        with open(danxuan_path, 'r', encoding='utf-8') as f:
            danxuan_lines = f.readlines()
        cards = extract_from_danxuan(danxuan_lines)
        all_cards.extend(cards)
        print(f"从单选知识点提取了 {len(cards)} 张卡片")

    # 读取temp_zhongdian.txt
    zhongdian_path = temp_dir / "temp_zhongdian.txt"
    if zhongdian_path.exists():
        with open(zhongdian_path, 'r', encoding='utf-8') as f:
            zhongdian_lines = f.readlines()
        cards = extract_from_zhongdian(zhongdian_lines)
        all_cards.extend(cards)
        print(f"从重点总结提取了 {len(cards)} 张卡片")

    # 读取temp_zhishi.txt
    zhishi_path = temp_dir / "temp_zhishi.txt"
    if zhishi_path.exists():
        with open(zhishi_path, 'r', encoding='utf-8') as f:
            zhishi_lines = f.readlines()
        cards = extract_from_zhishi(zhishi_lines)
        all_cards.extend(cards)
        print(f"从知识小点提取了 {len(cards)} 张卡片")

    # 3. 去重
    print(f"\n去重前卡片总数: {len(all_cards)}")
    all_cards = deduplicate_cards(all_cards)
    print(f"去重后卡片总数: {len(all_cards)}")

    # 4. 质量筛选
    print(f"\n质量筛选前卡片总数: {len(all_cards)}")
    all_cards = filter_cards_by_quality(all_cards)
    print(f"质量筛选后卡片总数: {len(all_cards)}")

    # 4. 按章节分组并排序
    chapter_order = ['guide'] + [f'ch{i:02d}' for i in range(1, 18)]

    def sort_key(card):
        chapter_id = card.get('chapterId', 'guide')
        try:
            idx = chapter_order.index(chapter_id)
        except ValueError:
            idx = len(chapter_order)
        return idx

    all_cards.sort(key=sort_key)

    # 5. 添加ID
    for i, card in enumerate(all_cards, 1):
        card['id'] = f'card_{i:03d}'

    # 6. 统计卡片类型
    type_counts = {}
    for card in all_cards:
        card_type = card.get('type', '未知')
        type_counts[card_type] = type_counts.get(card_type, 0) + 1

    print("\n卡片类型分布:")
    for card_type, count in type_counts.items():
        print(f"  {card_type}: {count} ({count*100/len(all_cards):.1f}%)")

    # 7. 统计章节分布
    chapter_counts = {}
    for card in all_cards:
        chapter = card.get('chapter', '未知')
        chapter_counts[chapter] = chapter_counts.get(chapter, 0) + 1

    print("\n章节分布:")
    for chapter, count in sorted(chapter_counts.items()):
        print(f"  {chapter}: {count} 张")

    # 8. 生成输出
    output = {
        'meta': {
            'subject': '习近平新时代中国特色社会主义思想概论',
            'version': '3.0',
            'totalCards': len(all_cards)
        },
        'cards': all_cards
    }

    # 9. 保存到文件
    output_path = output_dir / "flashcards.json"
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\n已保存到 {output_path}")
    print(f"总计 {len(all_cards)} 张卡片")

    return 0


if __name__ == '__main__':
    sys.exit(main())