#!/usr/bin/env python3
"""Lightweight VBA reconstruction lint checks."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path


COMMON_OFFICE_PATTERNS = [
    "Shapes.AddShape",
    "Shapes.AddTextbox",
    "Shapes.AddLine",
    "Fill.ForeColor.RGB",
    "Line.ForeColor.RGB",
    "BuildFreeform",
]

CLEANUP_PATTERNS = [
    r"\.Delete\b",
    r"Shapes\.Count",
    r"For\s+Each\s+.*\s+In\s+.*Shapes",
    r"Do\s+While\s+.*Shapes\.Count",
]


def strip_comments(line: str) -> str:
    in_string = False
    result = []
    i = 0
    while i < len(line):
        char = line[i]
        if char == '"':
            if in_string and i + 1 < len(line) and line[i + 1] == '"':
                result.append(char)
                result.append(line[i + 1])
                i += 2
                continue
            in_string = not in_string
        if char == "'" and not in_string:
            break
        result.append(char)
        i += 1
    return "".join(result)


def has_unclosed_string(text: str) -> bool:
    for line in text.splitlines():
        code = strip_comments(line)
        in_string = False
        i = 0
        while i < len(code):
            if code[i] == '"':
                if in_string and i + 1 < len(code) and code[i + 1] == '"':
                    i += 2
                    continue
                in_string = not in_string
            i += 1
        if in_string:
            return True
    return False


def parentheses_balance(text: str) -> int:
    balance = 0
    in_string = False
    for line in text.splitlines():
        code = strip_comments(line)
        i = 0
        while i < len(code):
            char = code[i]
            if char == '"':
                if in_string and i + 1 < len(code) and code[i + 1] == '"':
                    i += 2
                    continue
                in_string = not in_string
            elif not in_string:
                if char == "(":
                    balance += 1
                elif char == ")":
                    balance -= 1
            i += 1
    return balance


def lint_vba(text: str) -> dict:
    errors: list[str] = []
    warnings: list[str] = []

    if not re.search(r"^\s*Sub\s+\w+", text, re.IGNORECASE | re.MULTILINE):
        errors.append("Missing Sub declaration.")
    if not re.search(r"^\s*End\s+Sub\b", text, re.IGNORECASE | re.MULTILINE):
        errors.append("Missing End Sub.")

    balance = parentheses_balance(text)
    if balance != 0:
        errors.append(f"Parentheses appear unbalanced: net balance {balance}.")

    if has_unclosed_string(text):
        errors.append("At least one line appears to contain an unclosed string literal.")

    found_patterns = [pattern for pattern in COMMON_OFFICE_PATTERNS if pattern in text]
    if not found_patterns:
        warnings.append("No common Office shape creation or formatting API patterns found.")
    elif not any(pattern in text for pattern in ("Shapes.AddShape", "Shapes.AddTextbox", "Shapes.AddLine", "BuildFreeform")):
        warnings.append("Formatting APIs found, but no obvious shape creation calls found.")

    if "Fill.ForeColor.RGB" not in text and "RGB(" not in text:
        warnings.append("No obvious RGB fill/color usage found.")

    if not any(re.search(pattern, text, re.IGNORECASE) for pattern in CLEANUP_PATTERNS):
        warnings.append("No obvious canvas cleanup logic found.")

    if ".Name" not in text:
        warnings.append("No shape naming assignments found.")

    return {
        "passed": not errors,
        "warnings": warnings,
        "errors": errors,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Run lightweight checks against generated Office VBA code.",
    )
    parser.add_argument("vba_file", help="Path to a .bas/.vba/.txt file containing VBA code.")
    parser.add_argument(
        "--pretty",
        action="store_true",
        help="Pretty-print JSON output.",
    )
    args = parser.parse_args()

    path = Path(args.vba_file)
    if not path.exists():
        print(json.dumps({"passed": False, "warnings": [], "errors": [f"File not found: {path}"]}))
        return 2
    if not path.is_file():
        print(json.dumps({"passed": False, "warnings": [], "errors": [f"Not a file: {path}"]}))
        return 2

    try:
        text = path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        try:
            text = path.read_text(encoding="utf-8-sig")
        except UnicodeDecodeError as exc:
            print(json.dumps({"passed": False, "warnings": [], "errors": [str(exc)]}))
            return 2
    except OSError as exc:
        print(json.dumps({"passed": False, "warnings": [], "errors": [str(exc)]}))
        return 2

    result = lint_vba(text)
    print(json.dumps(result, indent=2 if args.pretty else None))
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    sys.exit(main())
