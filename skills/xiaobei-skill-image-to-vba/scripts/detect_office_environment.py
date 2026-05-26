#!/usr/bin/env python3
"""Detect likely local Office/WPS presentation runtimes for materialization."""

from __future__ import annotations

import json
import os
import platform
import shutil
from pathlib import Path


def existing(paths: list[str]) -> list[str]:
    found = []
    for value in paths:
        path = Path(os.path.expandvars(value)).expanduser()
        if path.exists():
            found.append(str(path))
    return found


def windows_app_path(executable: str) -> list[str]:
    if platform.system() != "Windows":
        return []
    try:
        import winreg
    except ImportError:
        return []

    found: list[str] = []
    subkey = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{executable}"
    for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        try:
            with winreg.OpenKey(root, subkey) as key:
                value, _ = winreg.QueryValueEx(key, "")
                if value and Path(value).exists():
                    found.append(value)
        except OSError:
            continue
    return found


def command_path(command: str) -> list[str]:
    path = shutil.which(command)
    return [path] if path else []


def detect() -> dict:
    system = platform.system()
    result = {
        "platform": system,
        "presentation_apps": [],
        "automation_tools": [],
        "recommended_path": "manual_vba",
        "notes": [],
    }

    if system == "Darwin":
        powerpoint = existing(["/Applications/Microsoft PowerPoint.app"])
        wps = existing(["/Applications/WPS Office.app", "/Applications/WPS.app"])
        if powerpoint:
            result["presentation_apps"].append({"name": "Microsoft PowerPoint", "paths": powerpoint})
        if wps:
            result["presentation_apps"].append({"name": "WPS Office", "paths": wps})
        if shutil.which("osascript"):
            result["automation_tools"].append({"name": "osascript", "path": shutil.which("osascript")})
        if powerpoint and shutil.which("osascript"):
            result["recommended_path"] = "macos_powerpoint_applescript"
        elif wps:
            result["recommended_path"] = "wps_manual_or_open_only"
            result["notes"].append("WPS macro automation is not assumed; provide .bas and manual WPS steps unless verified locally.")
        else:
            result["recommended_path"] = "no_local_presentation_app"

    elif system == "Windows":
        powerpoint = windows_app_path("POWERPNT.EXE") + existing(
            [
                r"%ProgramFiles%\Microsoft Office\root\Office16\POWERPNT.EXE",
                r"%ProgramFiles(x86)%\Microsoft Office\root\Office16\POWERPNT.EXE",
                r"%ProgramFiles%\Microsoft Office\Office16\POWERPNT.EXE",
                r"%ProgramFiles(x86)%\Microsoft Office\Office16\POWERPNT.EXE",
            ]
        )
        wps = windows_app_path("wpp.exe") + existing(
            [
                r"%ProgramFiles%\Kingsoft\WPS Office\office6\wpp.exe",
                r"%ProgramFiles(x86)%\Kingsoft\WPS Office\office6\wpp.exe",
                r"%LOCALAPPDATA%\Kingsoft\WPS Office\office6\wpp.exe",
                r"%ProgramFiles%\WPS Office\office6\wpp.exe",
                r"%ProgramFiles(x86)%\WPS Office\office6\wpp.exe",
            ]
        )
        if powerpoint:
            result["presentation_apps"].append({"name": "Microsoft PowerPoint", "paths": sorted(set(powerpoint))})
        if wps:
            result["presentation_apps"].append({"name": "WPS Presentation", "paths": sorted(set(wps))})
        powershell = command_path("powershell.exe") + command_path("pwsh.exe")
        if powershell:
            result["automation_tools"].append({"name": "PowerShell", "paths": powershell})
        if powerpoint and powershell:
            result["recommended_path"] = "windows_powerpoint_com"
            result["notes"].append("VBA import may require PowerPoint Trust access to the VBA project object model.")
        elif wps:
            result["recommended_path"] = "wps_manual_or_open_only"
            result["notes"].append("Do not assume WPS exposes PowerPoint-compatible COM/VBA automation.")
        else:
            result["recommended_path"] = "no_local_presentation_app"

    else:
        wps_commands = command_path("wps") + command_path("wpp")
        if wps_commands:
            result["presentation_apps"].append({"name": "WPS Office", "paths": wps_commands})
            result["recommended_path"] = "wps_manual_or_open_only"
            result["notes"].append("Linux WPS can open presentation files, but VBA execution/automation is not assumed.")
        else:
            result["recommended_path"] = "no_local_presentation_app"

    if result["recommended_path"] == "no_local_presentation_app":
        result["notes"].append("Generate the VBA source and, when possible, an editable PPTX fallback; do not return only a PNG.")
    return result


def main() -> int:
    print(json.dumps(detect(), indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
