#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
from pathlib import Path


PROTECTED_PATTERNS: tuple[re.Pattern[str], ...] = (
    re.compile(r"\A---\s*\n.*?\n---\s*(?:\n|$)", re.DOTALL),
    re.compile(r"(?ms)^(```+|~~~+)[^\n]*\n.*?^\1[ \t]*$"),
    re.compile(r"`+[^`\n]*`+"),
    re.compile(r"<!--.*?-->", re.DOTALL),
    re.compile(r"<table\b.*?</table>", re.DOTALL | re.IGNORECASE),
    re.compile(r"\$\$.*?\$\$", re.DOTALL),
    re.compile(r"\\\[.*?\\\]", re.DOTALL),
    re.compile(r"\\\(.*?\\\)", re.DOTALL),
    re.compile(r"(?<!\\)\$(?!\$)(?:\\.|[^$\\\n])+(?<!\\)\$"),
    re.compile(r"\\begin\{([^{}]+)\}.*?\\end\{\1\}", re.DOTALL),
    re.compile(r"</?[A-Za-z][^>]*?>", re.DOTALL),
)


def is_escaped(text: str, index: int) -> bool:
    backslash_count = 0
    j = index - 1
    while j >= 0 and text[j] == "\\":
        backslash_count += 1
        j -= 1
    return backslash_count % 2 == 1


def find_protected_ranges(text: str) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    for pattern in PROTECTED_PATTERNS:
        for match in pattern.finditer(text):
            ranges.append((match.start(), match.end()))

    if not ranges:
        return []

    ranges.sort()
    merged: list[list[int]] = [[ranges[0][0], ranges[0][1]]]
    for start, end in ranges[1:]:
        last = merged[-1]
        if start <= last[1]:
            if end > last[1]:
                last[1] = end
            continue
        merged.append([start, end])
    return [(start, end) for start, end in merged]


def convert_quotes_in_plain_text(text: str, is_open_quote: bool) -> tuple[str, bool]:
    result: list[str] = []
    for i, ch in enumerate(text):
        if ch != '"' or is_escaped(text, i):
            result.append(ch)
            continue

        result.append("“" if is_open_quote else "”")
        is_open_quote = not is_open_quote
    return "".join(result), is_open_quote


def convert_straight_double_quotes_to_curly(text: str) -> str:
    protected_ranges = find_protected_ranges(text)
    if not protected_ranges:
        converted, _ = convert_quotes_in_plain_text(text, is_open_quote=True)
        return converted

    result: list[str] = []
    cursor = 0
    is_open_quote = True

    for start, end in protected_ranges:
        plain_segment = text[cursor:start]
        converted_segment, is_open_quote = convert_quotes_in_plain_text(plain_segment, is_open_quote)
        result.append(converted_segment)
        result.append(text[start:end])
        cursor = end

    tail_segment, _ = convert_quotes_in_plain_text(text[cursor:], is_open_quote)
    result.append(tail_segment)
    return "".join(result)


def process_file(input_path: Path, encoding: str = "utf-8") -> None:
    source_text = input_path.read_text(encoding=encoding)
    converted_text = convert_straight_double_quotes_to_curly(source_text)
    input_path.write_text(converted_text, encoding=encoding)


def backup_file(input_path: Path, backup_suffix: str, encoding: str = "utf-8") -> Path:
    backup_path = input_path.with_name(f"{input_path.name}{backup_suffix}")
    backup_path.write_text(input_path.read_text(encoding=encoding), encoding=encoding)
    return backup_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description='将中文文档中的直双引号（"）转换为弯双引号（“”），先备份输入文件，再原地覆盖。'
    )
    parser.add_argument("input", nargs="?", type=Path, default=Path("thesis.qmd"), help="输入文件路径（默认 thesis.qmd）")
    parser.add_argument("--backup-suffix", default=".bak", help="备份文件后缀，默认 .bak")
    parser.add_argument("--encoding", default="utf-8", help="文件编码，默认 utf-8")
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if not args.input.exists():
        raise FileNotFoundError(f"输入文件不存在: {args.input}")

    backup_path = backup_file(args.input, args.backup_suffix, encoding=args.encoding)
    process_file(args.input, encoding=args.encoding)
    print(f"已完成转换: {args.input}（备份: {backup_path}）")


if __name__ == "__main__":
    main()
