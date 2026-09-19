#!/usr/bin/env python3
"""Create a project-specific JSON configuration from the bundled default."""

from __future__ import annotations

import argparse
import copy
import json
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path, help="Path of the project configuration to create")
    parser.add_argument("--name", help="Project name")
    parser.add_argument("--document-type", help="Document type")
    parser.add_argument("--domain", action="append", help="Engineering domain; may be repeated")
    parser.add_argument("--mode", choices=("single", "team"), help="Collaboration mode")
    parser.add_argument("--target-pages", type=int, help="Target page count")
    parser.add_argument("--page-min", type=int, help="Minimum acceptable page count")
    parser.add_argument("--page-max", type=int, help="Maximum acceptable page count")
    parser.add_argument("--template", help="User-specified template path")
    parser.add_argument("--force", action="store_true", help="Overwrite an existing output file")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.output.exists() and not args.force:
        raise SystemExit(f"Refusing to overwrite existing file: {args.output}")

    default_path = Path(__file__).resolve().parents[1] / "config" / "default_config.json"
    data = copy.deepcopy(json.loads(default_path.read_text(encoding="utf-8")))

    if args.name:
        data["project"]["project_name"] = args.name
    if args.document_type:
        data["project"]["document_type"] = args.document_type
    if args.domain:
        data["project"]["engineering_domain"] = args.domain
    if args.mode:
        data["collaboration"]["mode"] = args.mode
    if args.target_pages is not None:
        data["scope"]["target_pages"] = args.target_pages
    low, high = data["scope"]["acceptable_page_range"]
    if args.page_min is not None:
        low = args.page_min
    if args.page_max is not None:
        high = args.page_max
    data["scope"]["acceptable_page_range"] = [low, high]
    if args.template:
        data["document"]["template_path"] = args.template

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(args.output.resolve())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
