#!/usr/bin/env python3
"""Attempt to run generated PowerPoint VBA through macOS PowerPoint automation."""

from __future__ import annotations

import argparse
import json
import platform
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


POWERPOINT_APP = Path("/Applications/Microsoft PowerPoint.app")


def apple_string(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def extract_sub_body(source: str, macro_name: str) -> str | None:
    pattern = re.compile(
        rf"^\s*(?:Public\s+|Private\s+)?Sub\s+{re.escape(macro_name)}\b[^\n]*\n"
        rf"(?P<body>.*?)^\s*End\s+Sub\b",
        re.IGNORECASE | re.MULTILINE | re.DOTALL,
    )
    match = pattern.search(source)
    if not match:
        return None
    body = match.group("body").strip()
    return body or None


def run_vba_file(vba_file: Path) -> subprocess.CompletedProcess[str]:
    script = [
        f'set vbaFile to POSIX file "{apple_string(str(vba_file))}"',
        "set vbaSource to read vbaFile",
        'tell application "Microsoft PowerPoint"',
        "activate",
        "do Visual Basic vbaSource",
        "end tell",
    ]
    command: list[str] = ["osascript"]
    for line in script:
        command.extend(["-e", line])
    return subprocess.run(command, text=True, capture_output=True, check=False)


def result(status: str, **extra) -> dict:
    payload = {"status": status}
    payload.update(extra)
    return payload


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Run generated PowerPoint VBA through AppleScript. This depends on "
            "Microsoft PowerPoint for Mac and local automation/macro permissions."
        )
    )
    parser.add_argument("vba_file", help="Path to a generated .bas/.vba/.txt file.")
    parser.add_argument(
        "--macro",
        default="ReconstructFromImage",
        help="Macro name used if the script needs to retry with only the Sub body.",
    )
    parser.add_argument(
        "--no-body-fallback",
        action="store_true",
        help="Do not retry by extracting the requested Sub body.",
    )
    parser.add_argument("--pretty", action="store_true", help="Pretty-print JSON output.")
    args = parser.parse_args()

    if platform.system() != "Darwin":
        print(json.dumps(result("unsupported_platform", platform=platform.system())))
        return 2
    if not POWERPOINT_APP.exists():
        print(json.dumps(result("powerpoint_not_found", expected_path=str(POWERPOINT_APP))))
        return 2
    if shutil.which("osascript") is None:
        print(json.dumps(result("osascript_not_found")))
        return 2

    vba_path = Path(args.vba_file).expanduser().resolve()
    if not vba_path.is_file():
        print(json.dumps(result("vba_file_not_found", path=str(vba_path))))
        return 2

    first = run_vba_file(vba_path)
    attempts = [
        {
            "mode": "full_source",
            "returncode": first.returncode,
            "stdout": first.stdout.strip(),
            "stderr": first.stderr.strip(),
        }
    ]
    if first.returncode == 0:
        print(json.dumps(result("ran", vba_file=str(vba_path), attempts=attempts), indent=2 if args.pretty else None))
        return 0

    if not args.no_body_fallback:
        try:
            source = vba_path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            source = vba_path.read_text(encoding="utf-8-sig")
        body = extract_sub_body(source, args.macro)
        if body:
            with tempfile.NamedTemporaryFile("w", suffix=".vba", encoding="utf-8", delete=False) as handle:
                handle.write(body)
                body_path = Path(handle.name)
            second = run_vba_file(body_path)
            attempts.append(
                {
                    "mode": "sub_body",
                    "temp_file": str(body_path),
                    "returncode": second.returncode,
                    "stdout": second.stdout.strip(),
                    "stderr": second.stderr.strip(),
                }
            )
            if second.returncode == 0:
                print(
                    json.dumps(
                        result(
                            "ran_with_sub_body",
                            vba_file=str(vba_path),
                            macro=args.macro,
                            attempts=attempts,
                        ),
                        indent=2 if args.pretty else None,
                    )
                )
                return 0

    print(
        json.dumps(
            result(
                "automation_failed",
                vba_file=str(vba_path),
                macro=args.macro,
                attempts=attempts,
                fallback=(
                    "Open PowerPoint, open the VBA editor, import or paste the VBA file, "
                    "then run the macro manually. Check macOS automation permissions and "
                    "Office macro security settings."
                ),
            ),
            indent=2 if args.pretty else None,
        )
    )
    return 1


if __name__ == "__main__":
    sys.exit(main())
