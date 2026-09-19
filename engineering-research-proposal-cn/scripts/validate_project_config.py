#!/usr/bin/env python3
"""Validate a project configuration and surface decisions still requiring confirmation."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("config", type=Path)
    parser.add_argument("--strict", action="store_true", help="Treat unresolved decisions as errors")
    parser.add_argument("--json", action="store_true", help="Print machine-readable result")
    return parser.parse_args()


def nested(data: dict[str, Any], *keys: str) -> Any:
    current: Any = data
    for key in keys:
        if not isinstance(current, dict) or key not in current:
            return None
        current = current[key]
    return current


def validate(data: dict[str, Any]) -> dict[str, list[str]]:
    errors: list[str] = []
    warnings: list[str] = []
    questions: list[str] = []

    for key in ("project", "collaboration", "scope", "structure", "evidence", "technical_indicators", "document", "formulas", "diagrams", "quality"):
        if not isinstance(data.get(key), dict):
            errors.append(f"缺少对象：{key}")

    if errors:
        return {"errors": errors, "warnings": warnings, "questions": questions}

    if nested(data, "project", "project_name") in (None, "", "待填写"):
        questions.append("请填写项目名称。")
    domains = nested(data, "project", "engineering_domain")
    if not domains or domains == ["待填写"]:
        questions.append("请填写主要工程领域；交叉领域可填写多项。")

    mode = nested(data, "collaboration", "mode")
    if mode not in ("single", "team"):
        errors.append("collaboration.mode 必须为 single 或 team。")
    if mode == "team":
        checks = (
            ("responsible_heading_level", "请确认负责的标题等级。"),
            ("responsible_heading_count", "请确认该等级下负责的标题数量。"),
            ("parent_heading", "请确认上级标题。"),
        )
        for field, message in checks:
            if nested(data, "collaboration", field) in (None, "", []):
                questions.append(message)
        if not nested(data, "collaboration", "responsible_heading_ids"):
            questions.append("请确认既定标题编号或明确编号由合稿负责人统一分配。")

    target = nested(data, "scope", "target_pages")
    page_range = nested(data, "scope", "acceptable_page_range")
    if not isinstance(target, int) or target <= 0:
        errors.append("scope.target_pages 必须为正整数。")
    if not isinstance(page_range, list) or len(page_range) != 2 or not all(isinstance(x, int) for x in page_range):
        errors.append("scope.acceptable_page_range 必须包含两个整数。")
    elif page_range[0] <= 0 or page_range[0] > page_range[1]:
        errors.append("页数范围无效。")
    elif isinstance(target, int) and not page_range[0] <= target <= page_range[1]:
        warnings.append("目标页数不在允许页数范围内。")

    subheading_range = nested(data, "structure", "preferred_parallel_subheading_range")
    if subheading_range != [3, 5]:
        warnings.append("并列小标题建议范围不是 3—5 个；请确认是否由固定模板规定。")
    if nested(data, "structure", "regroup_when_exceeded") is not True:
        errors.append("并列小标题超过建议数量时必须启用重新分组检查。")

    policy = nested(data, "technical_indicators", "policy")
    if policy != "ask_before_fixing_values":
        errors.append("technical_indicators.policy 必须保持 ask_before_fixing_values。")
    if not nested(data, "technical_indicators", "confirmed"):
        questions.append("请确认是否需要写确定技术指标，以及是否仅采用用户材料中的数值。")
    elif not nested(data, "technical_indicators", "register"):
        warnings.append("指标已标记为确认，但指标登记表为空。")

    if nested(data, "document", "template_path") is None:
        warnings.append("未指定用户模板，将采用 Skill 默认模板。")
    if nested(data, "document", "western_body_font") != "Times New Roman":
        warnings.append("西文正文字体不是 Times New Roman，请确认是否由用户模板规定。")
    if nested(data, "formulas", "disclose_actual_object_type") is not True:
        errors.append("必须如实披露公式对象类型。")
    if str(nested(data, "diagrams", "source_format") or "").upper() != "VSDX":
        errors.append("工程图源文件格式必须为 VSDX。")
    if nested(data, "diagrams", "require_glued_endpoints") is not True:
        errors.append("必须启用连接器端点粘附检查。")

    return {"errors": errors, "warnings": warnings, "questions": questions}


def main() -> int:
    args = parse_args()
    try:
        data = json.loads(args.config.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(json.dumps({"errors": [str(exc)], "warnings": [], "questions": []}, ensure_ascii=False, indent=2))
        return 2

    result = validate(data)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        for label, key in (("错误", "errors"), ("提醒", "warnings"), ("待确认", "questions")):
            print(f"{label}（{len(result[key])}）")
            for item in result[key]:
                print(f"- {item}")
    if result["errors"] or (args.strict and result["questions"]):
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
