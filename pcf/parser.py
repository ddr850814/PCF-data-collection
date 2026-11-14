"""Lossless PCF parser.

Structure rules (PCF Reference Guide V13):
- A line starting at column 1 is a record: header keyword, PIPELINE-REFERENCE,
  a component type, a connection record, or MATERIALS.
- An indented line (column 5; tabs normalised to 4 spaces) is an attribute of
  the current record.
- Everything after MATERIALS is the materials section, kept verbatim.

The parser never discards or rewrites a line it does not understand: unknown
records become components with zero typed meaning but full raw fidelity, so
``write_document(parse_document(text)) == text``.
"""

from __future__ import annotations

from pathlib import Path

from .model import AttrLine, Component, PcfDocument, Pipeline

_PIPELINE_KEY = "PIPELINE-REFERENCE"


def _is_indented(line: str) -> bool:
    return line.startswith((" ", "\t"))


def parse_document(text: str, source_path: str | None = None) -> PcfDocument:
    newline = "\r\n" if "\r\n" in text else "\n"
    doc = PcfDocument(newline=newline, source_path=source_path)

    lines = text.split("\n")
    # A trailing newline produces one empty final element; drop it but remember.
    trailing_newline = bool(lines) and lines[-1] == ""
    if trailing_newline:
        lines = lines[:-1]
    lines = [line.rstrip("\r") for line in lines]

    doc._had_trailing_newline = trailing_newline  # type: ignore[attr-defined]

    current_pipeline: Pipeline | None = None
    current_component: Component | None = None
    in_materials = False

    for line in lines:
        stripped = line.strip()
        upper = stripped.upper()

        if in_materials:
            doc.materials_lines.append(line)
            continue

        if stripped == "":
            # Preserve blank lines wherever they sit.
            if current_component is not None:
                current_component.attributes.append(AttrLine(line))
            elif current_pipeline is not None:
                current_pipeline.header_lines.append(AttrLine(line))
            else:
                doc.header_lines.append(line)
            continue

        if not _is_indented(line):
            if upper == "MATERIALS":
                in_materials = True
                doc.materials_lines.append(line)
                current_component = None
                continue
            if upper.startswith(_PIPELINE_KEY):
                reference = stripped[len(_PIPELINE_KEY):].strip()
                current_pipeline = Pipeline(reference=reference)
                current_pipeline.raw_line = line
                doc.pipelines.append(current_pipeline)
                current_component = None
                continue
            if current_pipeline is None:
                doc.header_lines.append(line)
                continue
            current_component = Component(type=stripped)
            current_pipeline.components.append(current_component)
            continue

        # Indented attribute line.
        if current_component is not None:
            current_component.attributes.append(AttrLine(line))
        elif current_pipeline is not None:
            current_pipeline.header_lines.append(AttrLine(line))
        else:
            doc.header_lines.append(line)

    return doc


def write_document(doc: PcfDocument) -> str:
    out: list[str] = []
    out.extend(doc.header_lines)
    for pipeline in doc.pipelines:
        raw = pipeline.raw_line
        if raw is not None and raw.strip()[len(_PIPELINE_KEY):].strip() == pipeline.reference:
            out.append(raw)
        else:
            out.append(f"{_PIPELINE_KEY} {pipeline.reference}".rstrip())
        out.extend(attr.raw for attr in pipeline.header_lines)
        for comp in pipeline.components:
            out.append(comp.type)
            out.extend(attr.raw for attr in comp.attributes)
    out.extend(doc.materials_lines)

    text = doc.newline.join(out)
    if getattr(doc, "_had_trailing_newline", True) and out:
        text += doc.newline
    return text


def _decode_pcf(raw: bytes, encoding: str) -> tuple[str, str]:
    """PCF 字节 → (文本, 实际使用的编码)。auto 的判定规则与收集器
    （pcf_parser._open_file）同源：BOM → UTF-8-SIG；否则 GBK → UTF-8
    依次严格试解；latin-1 兜底（中文项目多为 GBK，GBK 优先可避免中文
    描述被替换符污染）。"""
    if encoding != "auto":
        text = raw.decode(encoding, errors="replace")
        if text.startswith("\ufeff"):
            # 显式编码也要剥 BOM：否则首行变 \ufeffISOGEN，UNITS 头
            # 匹配失败会静默回退英寸单位
            text = text[1:]
        return text, encoding
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw[3:].decode("utf-8", errors="replace"), "utf-8-sig"
    for enc in ("gbk", "utf-8"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1"), "latin-1"


def decode_pcf_bytes(raw: bytes, encoding: str = "auto") -> str:
    """PCF 字节 → 文本，编码规则见 _decode_pcf。"""
    return _decode_pcf(raw, encoding)[0]


def load(path: str | Path, encoding: str = "auto") -> PcfDocument:
    path = Path(path)
    text, used = _decode_pcf(path.read_bytes(), encoding)
    doc = parse_document(text, source_path=str(path))
    # 记住读入编码，save 按原编码写回——GBK 原文件往返不再漂移成 UTF-8
    doc.source_encoding = used
    return doc


def save(doc: PcfDocument, path: str | Path) -> Path:
    path = Path(path)
    enc = getattr(doc, "source_encoding", None) or "utf-8"
    text = write_document(doc)
    try:
        path.write_text(text, encoding=enc, newline="")
    except UnicodeEncodeError:
        # 原编码装不下新内容（如 GBK 原文件被编辑加入 GBK 外字符）：
        # 退 UTF-8 带 BOM，读端 auto/utf-8-sig 都能正确识别
        path.write_text(text, encoding="utf-8-sig", newline="")
    return path
