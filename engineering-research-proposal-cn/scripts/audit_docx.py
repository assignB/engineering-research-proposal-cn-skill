#!/usr/bin/env python3
"""Inspect a DOCX package for structure, editable objects, fonts, and common writing risks."""

from __future__ import annotations

import argparse
import collections
import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "o": "urn:schemas-microsoft-com:office:office",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
}
W = "{" + NS["w"] + "}"

GENERIC_PHRASES = (
    "综上所述", "不难发现", "众所周知", "意义重大", "赋能", "抓手", "形成闭环", "全方位",
    "国际领先", "全面突破", "显著提升", "具有重要的理论意义和现实意义",
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--output", type=Path, help="Write the JSON report to this path")
    parser.add_argument("--template", action="store_true", help="Treat placeholders as expected template content")
    return parser.parse_args()


def xml_root(zf: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(name))
    except KeyError:
        return None


def attr(element: ET.Element | None, name: str) -> str | None:
    return None if element is None else element.get(W + name)


def word_bool(element: ET.Element | None) -> bool:
    if element is None:
        return False
    value = attr(element, "val")
    return value is None or value.lower() not in ("0", "false", "off", "no")


def paragraph_text(p: ET.Element) -> str:
    return "".join((node.text or "") for node in p.findall(".//w:t", NS)).strip()


def heading_level(style_name: str) -> int | None:
    match = re.match(r"^(?:Heading|标题)\s*([1-9])", style_name, re.I)
    if match:
        return int(match.group(1))
    match = re.match(r"^([一二三四五六七八九])级标题$", style_name)
    if match:
        return "一二三四五六七八九".index(match.group(1)) + 1
    return None


def style_map(styles: ET.Element | None) -> tuple[dict[str, str], dict[str, dict[str, object]]]:
    names: dict[str, str] = {}
    details: dict[str, dict[str, object]] = {}
    if styles is None:
        return names, details
    for style in styles.findall("w:style", NS):
        style_id = attr(style, "styleId") or ""
        name = attr(style.find("w:name", NS), "val") or style_id
        names[style_id] = name
        rpr = style.find("w:rPr", NS)
        rfonts = rpr.find("w:rFonts", NS) if rpr is not None else None
        size = rpr.find("w:sz", NS) if rpr is not None else None
        details[style_id] = {
            "name": name,
            "east_asia_font": attr(rfonts, "eastAsia"),
            "ascii_font": attr(rfonts, "ascii"),
            "hansi_font": attr(rfonts, "hAnsi"),
            "size_pt": (int(attr(size, "val")) / 2) if attr(size, "val") and attr(size, "val").isdigit() else None,
            "bold": word_bool(rpr.find("w:b", NS)) if rpr is not None else False,
        }
    return names, details


def audit(path: Path, template_mode: bool) -> dict[str, object]:
    report: dict[str, object] = {"file": str(path.resolve()), "errors": [], "warnings": []}
    if path.suffix.lower() != ".docx":
        report["errors"].append("文件扩展名不是 .docx。")
    if not path.exists():
        report["errors"].append("文件不存在。")
        return report

    try:
        with zipfile.ZipFile(path) as zf:
            bad = zf.testzip()
            if bad:
                report["errors"].append(f"ZIP 包内文件损坏：{bad}")
                return report
            names = set(zf.namelist())
            document = xml_root(zf, "word/document.xml")
            styles = xml_root(zf, "word/styles.xml")
            app = xml_root(zf, "docProps/app.xml")
            if document is None:
                report["errors"].append("缺少 word/document.xml。")
                return report

            style_names, style_details = style_map(styles)
            paragraphs = document.findall(".//w:body/w:p", NS)
            all_texts = [paragraph_text(p) for p in paragraphs]
            nonempty = [text for text in all_texts if text]
            heading_counts: collections.Counter[str] = collections.Counter()
            caption_count = 0
            style_use: collections.Counter[str] = collections.Counter()
            english_runs = 0
            explicit_non_tnr_runs = 0
            heading_entries: list[dict[str, object]] = []

            for p in paragraphs:
                text = paragraph_text(p)
                pstyle = p.find("w:pPr/w:pStyle", NS)
                style_id = attr(pstyle, "val") or ""
                style_name = style_names.get(style_id, style_id or "(none)")
                style_use[style_name] += 1
                level = heading_level(style_name)
                if level is not None:
                    heading_counts[style_name] += 1
                    heading_entries.append({"level": level, "text": text})
                if re.match(r"^(图|表)\s*\d", text):
                    caption_count += 1
                for run in p.findall("w:r", NS):
                    run_text = "".join((t.text or "") for t in run.findall("w:t", NS))
                    if re.search(r"[A-Za-z]", run_text):
                        english_runs += 1
                        rf = run.find("w:rPr/w:rFonts", NS)
                        explicit = attr(rf, "ascii") or attr(rf, "hAnsi")
                        if explicit and explicit.lower() != "times new roman":
                            explicit_non_tnr_runs += 1

            duplicates = [
                {"text": text[:120], "count": count}
                for text, count in collections.Counter(t for t in nonempty if len(t) >= 20).items()
                if count > 1
            ]
            joined = "\n".join(nonempty)
            placeholders = re.findall(r"待填写|待补充|TODO|TBD|XXX|\[待[^\]]*\]", joined, flags=re.I)
            generic = {phrase: joined.count(phrase) for phrase in GENERIC_PHRASES if phrase in joined}

            parent_stack: dict[int, dict[str, object]] = {}
            child_counts: dict[int, int] = collections.Counter()
            parent_rows: dict[int, dict[str, object]] = {}
            for index, entry in enumerate(heading_entries):
                level = int(entry["level"])
                parent_stack[level] = {"index": index, **entry}
                for deeper in [key for key in parent_stack if key > level]:
                    del parent_stack[deeper]
                if level > 1 and level - 1 in parent_stack:
                    parent = parent_stack[level - 1]
                    parent_index = int(parent["index"])
                    child_counts[parent_index] += 1
                    parent_rows[parent_index] = parent
            parallel_groups = [
                {"parent_level": row["level"], "parent_text": row["text"], "direct_subheading_count": child_counts[index]}
                for index, row in parent_rows.items()
            ]
            too_many_groups = [row for row in parallel_groups if row["direct_subheading_count"] > 5]

            xml_names = [n for n in names if n.startswith("word/") and n.endswith(".xml")]
            xml_bytes = b"\n".join(zf.read(n) for n in xml_names)
            xml_text = xml_bytes.decode("utf-8", errors="ignore")
            progids = re.findall(r"ProgID=\"([^\"]+)\"", xml_text, flags=re.I)
            math_progids = [p for p in progids if "equation" in p.lower() or "mathtype" in p.lower()]
            visio_progids = [p for p in progids if "visio" in p.lower()]
            embedded = [n for n in names if n.startswith("word/embeddings/")]
            media = [n for n in names if n.startswith("word/media/")]
            media_types = collections.Counter(Path(n).suffix.lower() for n in media)
            fields = re.findall(r"<w:instrText[^>]*>(.*?)</w:instrText>", xml_text, flags=re.I | re.S)
            field_kinds = collections.Counter((re.sub(r"\s+", " ", f).strip().split(" ")[0].upper() if f.strip() else "") for f in fields)
            pages = None
            if app is not None:
                node = app.find("ep:Pages", NS)
                if node is not None and node.text and node.text.isdigit():
                    pages = int(node.text)

            report.update({
                "pages_from_properties": pages,
                "paragraph_count": len(paragraphs),
                "nonempty_paragraph_count": len(nonempty),
                "table_count": len(document.findall(".//w:tbl", NS)),
                "heading_counts": dict(heading_counts),
                "parallel_subheading_groups": parallel_groups,
                "too_many_parallel_subheading_groups": too_many_groups,
                "caption_count": caption_count,
                "style_use_top": dict(style_use.most_common(20)),
                "key_style_definitions": {
                    style_id: detail for style_id, detail in style_details.items()
                    if str(detail["name"]).lower() in ("normal", "正文", "heading 1", "heading 2", "heading 3", "标题 1", "标题 2", "标题 3", "caption", "题注")
                    or re.match(r"^[一二三四五六七八九]级标题$", str(detail["name"]))
                },
                "english_runs": english_runs,
                "english_runs_with_explicit_non_times_new_roman": explicit_non_tnr_runs,
                "ole_progids": dict(collections.Counter(progids)),
                "mathtype_or_equation_ole_count": len(math_progids),
                "visio_ole_count": len(visio_progids),
                "omml_formula_count": len(re.findall(r"<m:oMath(?:Para)?\b", xml_text)),
                "embedded_object_count": len(embedded),
                "media_count": len(media),
                "media_types": dict(media_types),
                "field_kinds": dict(field_kinds),
                "duplicate_paragraphs": duplicates,
                "placeholder_count": len(placeholders),
                "generic_phrase_counts": generic,
            })

            if duplicates and not template_mode:
                report["warnings"].append("存在重复的长段落，请人工判断是否为必要重复。")
            if placeholders and not template_mode:
                report["warnings"].append("正文仍含待填写或占位标记。")
            if explicit_non_tnr_runs:
                report["warnings"].append("部分英文文本显式使用了非 Times New Roman 字体。")
            if generic:
                report["warnings"].append("检测到可能空泛或证据不足的常见表达，请结合语境复核。")
            if too_many_groups:
                report["warnings"].append("存在同一上级标题下超过 5 个并列小标题的结构，请优先重新分组。")
    except (OSError, zipfile.BadZipFile, ET.ParseError) as exc:
        report["errors"].append(str(exc))
    return report


def main() -> int:
    args = parse_args()
    result = audit(args.docx, args.template)
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 2 if result["errors"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
