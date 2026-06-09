# -*- coding: utf-8 -*-
"""Validate and count slides from slides.json"""
import json
import os
import sys

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')

def main():
    json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data', 'slides.json')

    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    items = data.get('items', [])
    print(f"Total slides: {len(items)}")
    print()

    # Count by chapter
    chapter_count = {}
    for item in items:
        ch = item.get('chapter', 'Unknown')
        chapter_count[ch] = chapter_count.get(ch, 0) + 1

    print("Chapter distribution:")
    print("-" * 50)
    for ch, count in chapter_count.items():
        print(f"  {ch}: {count} questions")
    print("-" * 50)
    print(f"  Total: {len(items)} questions")
    print()

    # Type distribution
    type_count = {}
    for item in items:
        typ = item.get('type', 'Unknown')
        type_count[typ] = type_count.get(typ, 0) + 1

    print("Type distribution:")
    print("-" * 50)
    for typ, count in type_count.items():
        print(f"  {typ}: {count}")
    print("-" * 50)

    # Check duplicates
    questions = [item['question'] for item in items]
    if len(questions) != len(set(questions)):
        print("\nWARNING: Duplicate questions found!")
        seen = set()
        for q in questions:
            if q in seen:
                print(f"  Duplicate: {q}")
            seen.add(q)
    else:
        print("\nNo duplicate questions. OK")

    # Validate lengths
    long_q = [i for i in items if len(i['question']) > 30]
    long_a = [i for i in items if len(i['answer']) > 80]
    print(f"\nQuestions > 30 chars: {len(long_q)}")
    print(f"Answers > 80 chars: {len(long_a)}")

if __name__ == "__main__":
    main()
