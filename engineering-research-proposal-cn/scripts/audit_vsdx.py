#!/usr/bin/env python3
"""Audit VSDX package structure and connector endpoint glue relationships."""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("vsdx", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--strict", action="store_true", help="Fail when any connector endpoint is unglued")
    return parser.parse_args()


def cell_value(shape: ET.Element, name: str) -> str | None:
    for child in shape.iter():
        if local_name(child.tag) == "Cell" and child.get("N") == name:
            return child.get("V")
    return None


def shape_text(shape: ET.Element) -> str:
    parts: list[str] = []
    for child in shape.iter():
        if local_name(child.tag) == "Text":
            parts.extend(child.itertext())
    return re.sub(r"\s+", " ", "".join(parts)).strip()


def audit(path: Path) -> dict[str, object]:
    result: dict[str, object] = {"file": str(path.resolve()), "errors": [], "warnings": [], "pages": []}
    if path.suffix.lower() != ".vsdx":
        result["errors"].append("工程图源文件扩展名必须为 .vsdx。")
    if not path.exists():
        result["errors"].append("文件不存在。")
        return result
    try:
        with zipfile.ZipFile(path) as zf:
            bad = zf.testzip()
            if bad:
                result["errors"].append(f"VSDX 包内文件损坏：{bad}")
                return result
            page_names = sorted(
                n for n in zf.namelist()
                if re.fullmatch(r"visio/pages/page\d+\.xml", n, flags=re.I)
            )
            if not page_names:
                result["errors"].append("未找到 Visio 页面 XML。")
                return result

            total_connectors = 0
            total_fully_glued = 0
            for page_name in page_names:
                root = ET.fromstring(zf.read(page_name))
                shapes = [node for node in root.iter() if local_name(node.tag) == "Shape"]
                ids = {shape.get("ID") for shape in shapes if shape.get("ID")}
                one_d: dict[str, ET.Element] = {}
                for shape in shapes:
                    sid = shape.get("ID")
                    explicit_one_d = cell_value(shape, "OneD") in ("1", "TRUE", "true")
                    endpoint_cells = cell_value(shape, "BeginX") is not None and cell_value(shape, "EndX") is not None
                    if sid and (explicit_one_d or endpoint_cells):
                        one_d[sid] = shape

                endpoint_connections: dict[str, set[str]] = {sid: set() for sid in one_d}
                invalid_targets: list[dict[str, str]] = []
                for node in root.iter():
                    if local_name(node.tag) != "Connect":
                        continue
                    source = node.get("FromSheet")
                    from_cell = node.get("FromCell") or ""
                    target = node.get("ToSheet") or ""
                    if source in endpoint_connections and from_cell in ("BeginX", "EndX"):
                        endpoint_connections[source].add(from_cell)
                        if target not in ids:
                            invalid_targets.append({"connector": source, "endpoint": from_cell, "target": target})

                connector_rows = []
                for sid, shape in one_d.items():
                    endpoints = endpoint_connections.get(sid, set())
                    fully = endpoints == {"BeginX", "EndX"}
                    connector_rows.append({
                        "shape_id": sid,
                        "text": shape_text(shape)[:100],
                        "begin_glued": "BeginX" in endpoints,
                        "end_glued": "EndX" in endpoints,
                    })
                    total_connectors += 1
                    total_fully_glued += int(fully)

                result["pages"].append({
                    "part": page_name,
                    "shape_count": len(shapes),
                    "connector_count": len(one_d),
                    "fully_glued_connector_count": sum(1 for row in connector_rows if row["begin_glued"] and row["end_glued"]),
                    "connectors": connector_rows,
                    "invalid_connection_targets": invalid_targets,
                })

            result["page_count"] = len(page_names)
            result["connector_count"] = total_connectors
            result["fully_glued_connector_count"] = total_fully_glued
            result["unglued_connector_count"] = total_connectors - total_fully_glued
            if total_connectors == 0:
                result["warnings"].append("未检测到一维连接器；若图中有连线，请检查是否误用了普通线段。")
            if total_connectors != total_fully_glued:
                result["warnings"].append("存在至少一个未同时粘附起点和终点的连接器。")
    except (OSError, zipfile.BadZipFile, ET.ParseError) as exc:
        result["errors"].append(str(exc))
    return result


def main() -> int:
    args = parse_args()
    result = audit(args.vsdx)
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(text + "\n", encoding="utf-8")
    print(text)
    if result["errors"]:
        return 2
    if args.strict and result.get("unglued_connector_count", 0):
        return 3
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
