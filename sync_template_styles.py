#!/usr/bin/env python3

from __future__ import annotations

from pathlib import Path

from template_styles import sync_all_templates


def main() -> None:
    root = Path(__file__).resolve().parent
    templates_dir = root / "templates"
    template_paths = sorted(templates_dir.glob("*_template.docx"))
    sync_all_templates(template_paths)
    print(f"Synchronized styles for {len(template_paths)} templates.")


if __name__ == "__main__":
    main()
