"""统计文本文件中的词频"""

import sys
from collections import Counter


def count_words(filepath):
    with open(filepath, "r", encoding="utf-8") as f:
        text = f.read()
    words = text.lower().split()
    return Counter(words).most_common(10)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python word_counter.py <文件路径>")
        sys.exit(1)

    top_words = count_words(sys.argv[1])
    print("Top 10 高频词:")
    for word, count in top_words:
        print(f"  {word}: {count}")
