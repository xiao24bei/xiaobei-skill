# XiaoBei Skill Collection Routing

This repository is a collection of independently installable Codex skills. The repository root is intentionally not an installable skill and must not contain `SKILL.md`.

When a user asks to install `xiaobei-skill`, this repository, or its root URL without naming a child skill:

1. Read `skill-catalog.json`.
2. Show every listed skill's `display_name`, description, and path.
3. Ask the user which one they want.
4. After the user chooses, install only that exact child path with `$skill-installer`.
5. Do not infer a default and do not install all skills unless the user explicitly asks for all of them.

When a user explicitly names one child skill, install that child directly. When the user explicitly asks for all skills, install every path listed in the catalog.

Do not treat repository documentation, examples, gallery assets, or the repository root as skills.
