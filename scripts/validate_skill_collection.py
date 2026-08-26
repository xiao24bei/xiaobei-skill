#!/usr/bin/env python3
"""Validate the machine-readable catalog and independent Skill packages."""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CATALOG_PATH = ROOT / "skill-catalog.json"
NAME_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
FRONTMATTER_NAME_RE = re.compile(r"^name:\s*[\"']?([^\"'\s]+)[\"']?\s*$", re.MULTILINE)
DISPLAY_NAME_RE = re.compile(r'^\s*display_name:\s*["\'](.+)["\']\s*$', re.MULTILINE)
DEFAULT_PROMPT_RE = re.compile(r'^\s*default_prompt:\s*["\'](.+)["\']\s*$', re.MULTILINE)
ICON_RE = re.compile(r'^\s*icon_(?:small|large):\s*["\'](.+)["\']\s*$', re.MULTILINE)
BRAND_PREFIX = "小北在读研 · "


def fail(message: str) -> None:
    print(f"ERROR: {message}", file=sys.stderr)
    raise SystemExit(1)


def main() -> None:
    if (ROOT / "SKILL.md").exists():
        fail("repository root must not contain SKILL.md")

    try:
        catalog = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        fail(f"cannot read catalog: {exc}")

    if catalog.get("root_installable") is not False:
        fail("catalog.root_installable must be false")
    if catalog.get("unspecified_install_policy") != "ask_user_to_choose":
        fail("unspecified installs must ask the user to choose")

    entries = catalog.get("skills")
    if not isinstance(entries, list) or not entries:
        fail("catalog.skills must be a non-empty list")

    seen_names: set[str] = set()
    seen_paths: set[str] = set()
    for entry in entries:
        name = entry.get("name")
        rel_path = entry.get("path")
        display_name = entry.get("display_name")
        description = entry.get("description")
        if not isinstance(name, str) or not NAME_RE.fullmatch(name):
            fail(f"invalid skill name: {name!r}")
        if name in seen_names:
            fail(f"duplicate skill name: {name}")
        if not isinstance(rel_path, str) or rel_path in seen_paths:
            fail(f"invalid or duplicate skill path: {rel_path!r}")
        if Path(rel_path).name != name:
            fail(f"skill directory basename must match name: {rel_path}")
        if not isinstance(display_name, str) or not display_name.startswith(BRAND_PREFIX):
            fail(f"display_name must start with {BRAND_PREFIX!r}: {display_name!r}")
        if not isinstance(description, str) or not description.strip():
            fail(f"catalog description must be non-empty: {name}")

        skill_dir = ROOT / rel_path
        required = [
            skill_dir / "SKILL.md",
            skill_dir / "agents" / "openai.yaml",
            skill_dir / "requirements.txt",
        ]
        for path in required:
            if not path.is_file():
                fail(f"missing required file: {path.relative_to(ROOT)}")

        skill_md = (skill_dir / "SKILL.md").read_text(encoding="utf-8")
        match = FRONTMATTER_NAME_RE.search(skill_md)
        if not match or match.group(1) != name:
            fail(f"SKILL.md name does not match catalog: {name}")

        openai_yaml = (skill_dir / "agents" / "openai.yaml").read_text(encoding="utf-8")
        match = DISPLAY_NAME_RE.search(openai_yaml)
        if not match or match.group(1) != display_name:
            fail(f"openai.yaml display_name does not match catalog: {name}")
        prompt_match = DEFAULT_PROMPT_RE.search(openai_yaml)
        if not prompt_match or f"${name}" not in prompt_match.group(1):
            fail(f"openai.yaml default_prompt must mention ${name}")
        for icon_match in ICON_RE.finditer(openai_yaml):
            icon_value = icon_match.group(1)
            icon_path = (skill_dir / icon_value).resolve()
            try:
                icon_path.relative_to(skill_dir.resolve())
            except ValueError:
                fail(f"icon path escapes skill directory: {icon_value}")
            if not icon_path.is_file():
                fail(f"missing icon asset: {icon_value}")

        seen_names.add(name)
        seen_paths.add(rel_path)

    actual_skill_dirs = {
        str(path.relative_to(ROOT))
        for path in (ROOT / "skills").iterdir()
        if path.is_dir() and (path / "SKILL.md").is_file()
    }
    if actual_skill_dirs != seen_paths:
        fail(
            "catalog paths differ from installable skill directories: "
            f"catalog={sorted(seen_paths)}, actual={sorted(actual_skill_dirs)}"
        )

    print(f"Validated {len(entries)} independently installable XiaoBei skills.")


if __name__ == "__main__":
    main()
