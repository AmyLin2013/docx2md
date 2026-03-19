# Copyright (c) 2025 AmyLin <zhi_lin@qq.com>
# Licensed under the MIT License. See LICENSE file for details.

"""Extract heading structure (outline levels + auto-numbering) from Word (.docx).

Word stores:
- Heading levels via `outlineLvl` in paragraph/style properties (the authoritative
  source for heading hierarchy, not style names)
- Auto-numbering (e.g. "第一章", "1.1") in numbering.xml, not as inline text

This module parses the .docx XML to extract both, producing a list of
paragraph entries with their heading level and numbering prefix.
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": WORD_NS}


@dataclass
class ParagraphInfo:
    """Information extracted from a single Word paragraph."""
    text: str                     # Plain text content
    heading_level: Optional[int]  # 1-6 if heading, None if body text
    numbering_prefix: Optional[str]  # Auto-numbering text, e.g. "第一章"


# ── Number formatting ──

_CHINESE_NUMS = [
    "零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
    "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九",
    "二十", "二十一", "二十二", "二十三", "二十四", "二十五", "二十六",
    "二十七", "二十八", "二十九", "三十",
]

_ROMAN_VALS = [
    (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
    (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
    (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
]


def _to_roman(num: int) -> str:
    result = ""
    for val, sym in _ROMAN_VALS:
        while num >= val:
            result += sym
            num -= val
    return result


def _format_number(value: int, num_fmt: str) -> str:
    """Format a counter value according to a Word numbering format."""
    if num_fmt == "decimal":
        return str(value)
    elif num_fmt in (
        "chineseCounting", "chineseCountingThousand",
        "ideographTraditional", "japaneseCounting",
        "koreanDigital2",
    ):
        if 0 <= value < len(_CHINESE_NUMS):
            return _CHINESE_NUMS[value]
        return str(value)
    elif num_fmt == "upperLetter":
        return chr(ord("A") + value - 1) if 1 <= value <= 26 else str(value)
    elif num_fmt == "lowerLetter":
        return chr(ord("a") + value - 1) if 1 <= value <= 26 else str(value)
    elif num_fmt == "upperRoman":
        return _to_roman(value)
    elif num_fmt == "lowerRoman":
        return _to_roman(value).lower()
    elif num_fmt == "none":
        return ""
    else:
        return str(value)


def _format_level_text(
    lvl_text: str,
    levels: dict[int, dict],
    counters: dict[int, int],
    current_ilvl: int,
) -> str:
    """Expand a level-text pattern like '第%1章' or '%1.%2'.

    When the level has isLgl (isLegalNumberingStyle), placeholders that
    reference OTHER levels are formatted as decimal, but the current level
    keeps its own numFmt. This matches Word's behavior where "第%1章"
    at ilvl=0 still uses Chinese numerals (koreanDigital2), while "%1.%2"
    at ilvl=1 converts %1 (ilvl=0's koreanDigital2) to decimal "1".
    """
    is_lgl = levels.get(current_ilvl, {}).get("isLgl", False)
    result = lvl_text
    for ilvl in range(current_ilvl + 1):
        placeholder = f"%{ilvl + 1}"
        if placeholder in result:
            value = counters.get(ilvl, levels.get(ilvl, {}).get("start", 1))
            if is_lgl and ilvl != current_ilvl:
                # isLgl: convert other levels' numbers to decimal
                fmt = "decimal"
            else:
                fmt = levels.get(ilvl, {}).get("numFmt", "decimal")
            result = result.replace(placeholder, _format_number(value, fmt))
    return result


# ── XML parsing helpers ──

def _attr(el, name: str) -> Optional[str]:
    """Get a w: namespaced attribute value."""
    return el.get(f"{{{WORD_NS}}}{name}")


def _parse_abstract_nums(numbering_root) -> dict[str, dict[int, dict]]:
    """Parse abstractNum definitions → {abstractNumId: {ilvl: {numFmt, lvlText, start, pStyle, isLgl}}}."""
    result = {}
    for abstract_num in numbering_root.findall(".//w:abstractNum", NS):
        abstract_id = _attr(abstract_num, "abstractNumId")
        levels = {}
        for lvl in abstract_num.findall("w:lvl", NS):
            ilvl = int(_attr(lvl, "ilvl") or "0")
            num_fmt_el = lvl.find("w:numFmt", NS)
            lvl_text_el = lvl.find("w:lvlText", NS)
            start_el = lvl.find("w:start", NS)
            pstyle_el = lvl.find("w:pStyle", NS)
            is_lgl_el = lvl.find("w:isLgl", NS)

            level_info = {
                "numFmt": _attr(num_fmt_el, "val") if num_fmt_el is not None else "decimal",
                "lvlText": _attr(lvl_text_el, "val") if lvl_text_el is not None else "",
                "start": int(_attr(start_el, "val") or "1") if start_el is not None else 1,
            }
            if pstyle_el is not None:
                level_info["pStyle"] = _attr(pstyle_el, "val")
            if is_lgl_el is not None:
                level_info["isLgl"] = True
            levels[ilvl] = level_info
        result[abstract_id] = levels
    return result


def _parse_num_mappings(numbering_root) -> dict[str, str]:
    """Parse num elements → {numId: abstractNumId}."""
    result = {}
    for num in numbering_root.findall(".//w:num", NS):
        num_id = _attr(num, "numId")
        abstract_ref = num.find("w:abstractNumId", NS)
        if abstract_ref is not None:
            result[num_id] = _attr(abstract_ref, "val")
    return result


def _parse_num_overrides(numbering_root) -> dict[str, dict[int, dict]]:
    """Parse lvlOverride in num elements.

    Handles both <w:startOverride> (simple start value override) and full
    <w:lvl> redefinitions inside <w:lvlOverride> (which can change numFmt,
    lvlText, start, pStyle, isLgl, etc.).
    """
    overrides = {}
    for num in numbering_root.findall(".//w:num", NS):
        num_id = _attr(num, "numId")
        lvl_overrides = {}
        for override in num.findall("w:lvlOverride", NS):
            ilvl = int(_attr(override, "ilvl") or "0")
            lvl_info: dict = {}

            # Simple start override
            start_override_el = override.find("w:startOverride", NS)
            if start_override_el is not None:
                lvl_info["start"] = int(_attr(start_override_el, "val") or "1")

            # Full level redefinition
            lvl_el = override.find("w:lvl", NS)
            if lvl_el is not None:
                num_fmt_el = lvl_el.find("w:numFmt", NS)
                lvl_text_el = lvl_el.find("w:lvlText", NS)
                start_el = lvl_el.find("w:start", NS)
                pstyle_el = lvl_el.find("w:pStyle", NS)
                is_lgl_el = lvl_el.find("w:isLgl", NS)

                if num_fmt_el is not None:
                    lvl_info["numFmt"] = _attr(num_fmt_el, "val")
                if lvl_text_el is not None:
                    lvl_info["lvlText"] = _attr(lvl_text_el, "val")
                if start_el is not None and "start" not in lvl_info:
                    lvl_info["start"] = int(_attr(start_el, "val") or "1")
                if pstyle_el is not None:
                    lvl_info["pStyle"] = _attr(pstyle_el, "val")
                if is_lgl_el is not None:
                    lvl_info["isLgl"] = True

            if lvl_info:
                lvl_overrides[ilvl] = lvl_info
        if lvl_overrides:
            overrides[num_id] = lvl_overrides
    return overrides


def _build_pstyle_map(
    abstract_nums: dict[str, dict[int, dict]],
    num_to_abstract: dict[str, str],
    num_overrides: dict[str, dict[int, dict]],
) -> dict[str, tuple[str, int]]:
    """Build {pStyle: (numId, ilvl)} mapping from numbering definitions.

    When a <w:lvl> within a numbering definition contains <w:pStyle>,
    it establishes a direct link between that paragraph style and the
    numbering level. This allows heading styles to be correctly mapped
    to their intended numId + ilvl, even when the style's own numPr
    references a different numId.

    Override-level pStyle definitions take priority over abstract-level ones.
    """
    pstyle_map: dict[str, tuple[str, int]] = {}

    # First pass: abstract num pStyle entries
    for abstract_id, levels in abstract_nums.items():
        # Find numIds that reference this abstractNum
        for num_id, ref_abstract in num_to_abstract.items():
            if ref_abstract != abstract_id:
                continue
            for ilvl, level_def in levels.items():
                if "pStyle" in level_def:
                    pstyle_map[level_def["pStyle"]] = (num_id, ilvl)

    # Second pass: override pStyle entries (take priority)
    for num_id, lvl_overrides in num_overrides.items():
        for ilvl, override_def in lvl_overrides.items():
            if "pStyle" in override_def:
                pstyle_map[override_def["pStyle"]] = (num_id, ilvl)

    return pstyle_map


def _get_effective_levels(
    num_id: str,
    abstract_nums: dict[str, dict[int, dict]],
    num_to_abstract: dict[str, str],
    num_overrides: dict[str, dict[int, dict]],
) -> dict[int, dict]:
    """Get effective level definitions for a numId, merging abstract + overrides.

    Override properties take priority; unoverridden properties come from the
    abstract numbering definition.
    """
    abstract_id = num_to_abstract.get(num_id, "")
    base_levels = abstract_nums.get(abstract_id, {})
    result = {}
    for ilvl, base_def in base_levels.items():
        merged = dict(base_def)
        if num_id in num_overrides and ilvl in num_overrides[num_id]:
            merged.update(num_overrides[num_id][ilvl])
        result[ilvl] = merged
    return result


def _parse_styles(styles_root) -> dict[str, dict]:
    """Parse styles.xml → {styleId: {outlineLvl, numId, ilvl, basedOn}}.

    Extracts outline level, numbering reference, and style inheritance.
    """
    result = {}
    for style in styles_root.findall(".//w:style", NS):
        style_id = _attr(style, "styleId")
        info: dict = {}

        # Outline level (in pPr)
        ppr = style.find("w:pPr", NS)
        if ppr is not None:
            outline_el = ppr.find("w:outlineLvl", NS)
            if outline_el is not None:
                val = _attr(outline_el, "val")
                if val is not None:
                    lvl = int(val)
                    if 0 <= lvl <= 8:  # 0-8, where 9 = body text
                        info["outlineLvl"] = lvl

            numpr = ppr.find("w:numPr", NS)
            if numpr is not None:
                num_id_el = numpr.find("w:numId", NS)
                ilvl_el = numpr.find("w:ilvl", NS)
                if num_id_el is not None:
                    num_id = _attr(num_id_el, "val")
                    if num_id:
                        # Store numId even if "0" — it means "disable numbering, don't inherit"
                        info["numId"] = num_id
                # Store ilvl even without numId — the numId may be inherited
                # from a parent style, and this style's ilvl should override.
                if ilvl_el is not None:
                    info["ilvl"] = int(_attr(ilvl_el, "val") or "0")

        # Style inheritance
        based_on = style.find("w:basedOn", NS)
        if based_on is not None:
            info["basedOn"] = _attr(based_on, "val")

        if info:
            result[style_id] = info

    return result


def _resolve_outline_level(
    style_id: Optional[str],
    styles: dict[str, dict],
    _visited: Optional[set] = None,
) -> Optional[int]:
    """Resolve outlineLvl for a style, following basedOn inheritance chain."""
    if not style_id or style_id not in styles:
        return None

    if _visited is None:
        _visited = set()
    if style_id in _visited:
        return None  # prevent loops
    _visited.add(style_id)

    info = styles[style_id]
    if "outlineLvl" in info:
        return info["outlineLvl"]

    # Follow inheritance
    return _resolve_outline_level(info.get("basedOn"), styles, _visited)


def _resolve_numpr(
    style_id: Optional[str],
    styles: dict[str, dict],
    _visited: Optional[set] = None,
) -> Optional[tuple[str, int]]:
    """Resolve numId+ilvl for a style, following basedOn inheritance.

    When a child style has an explicit ilvl in its numPr but no numId,
    and the numId is inherited from a parent style, the child's ilvl
    takes precedence over the parent's ilvl.
    """
    if not style_id or style_id not in styles:
        return None

    if _visited is None:
        _visited = set()
    if style_id in _visited:
        return None
    _visited.add(style_id)

    info = styles[style_id]
    if "numId" in info:
        return (info["numId"], info.get("ilvl", 0))

    # Inherit numId from parent, but prefer this style's ilvl if set
    parent_result = _resolve_numpr(info.get("basedOn"), styles, _visited)
    if parent_result:
        num_id, parent_ilvl = parent_result
        # Use this style's ilvl if explicitly set, otherwise parent's
        ilvl = info.get("ilvl", parent_ilvl)
        return (num_id, ilvl)
    return None


def _get_paragraph_text(para) -> str:
    """Extract plain text from a w:p element."""
    texts = []
    for r in para.findall(".//w:r", NS):
        for t in r.findall("w:t", NS):
            if t.text:
                texts.append(t.text)
    return "".join(texts).strip()


# ── Main API ──

def extract_paragraph_info(docx_path: str | Path) -> list[ParagraphInfo]:
    """Extract heading levels, numbering prefixes, and text for all paragraphs.

    Heading level is determined by `outlineLvl` in paragraph properties or
    inherited from the style definition — this is the authoritative source
    that Word uses, independent of style naming.

    Args:
        docx_path: Path to the .docx file.

    Returns:
        List of ParagraphInfo for all non-empty paragraphs, in document order.
    """
    docx_path = Path(docx_path)

    with zipfile.ZipFile(docx_path, "r") as z:
        namelist = z.namelist()

        document_xml = ET.fromstring(z.read("word/document.xml"))

        # Parse numbering.xml
        has_numbering = "word/numbering.xml" in namelist
        if has_numbering:
            numbering_xml = ET.fromstring(z.read("word/numbering.xml"))
            abstract_nums = _parse_abstract_nums(numbering_xml)
            num_to_abstract = _parse_num_mappings(numbering_xml)
            num_overrides = _parse_num_overrides(numbering_xml)
        else:
            abstract_nums = {}
            num_to_abstract = {}
            num_overrides = {}

        # Parse styles.xml (with outlineLvl + numPr + basedOn)
        has_styles = "word/styles.xml" in namelist
        if has_styles:
            styles_xml = ET.fromstring(z.read("word/styles.xml"))
            styles = _parse_styles(styles_xml)
        else:
            styles = {}

    # Process document paragraphs
    counters: dict[str, dict[int, int]] = {}
    results: list[ParagraphInfo] = []

    body = document_xml.find("w:body", NS)
    if body is None:
        return results

    for para in body.findall(".//w:p", NS):
        para_text = _get_paragraph_text(para)
        if not para_text:
            continue

        ppr = para.find("w:pPr", NS)
        style_id = None
        num_id = None
        ilvl = None
        outline_lvl = None

        if ppr is not None:
            # Style reference
            style_el = ppr.find("w:pStyle", NS)
            if style_el is not None:
                style_id = _attr(style_el, "val")

            # Direct outlineLvl on paragraph
            outline_el = ppr.find("w:outlineLvl", NS)
            if outline_el is not None:
                val = _attr(outline_el, "val")
                if val is not None:
                    lvl = int(val)
                    if 0 <= lvl <= 8:
                        outline_lvl = lvl

            # Direct numPr on paragraph
            numpr = ppr.find("w:numPr", NS)
            if numpr is not None:
                num_id_el = numpr.find("w:numId", NS)
                ilvl_el = numpr.find("w:ilvl", NS)
                if num_id_el is not None:
                    num_id = _attr(num_id_el, "val")
                if ilvl_el is not None:
                    ilvl = int(_attr(ilvl_el, "val") or "0")

        # Resolve outline level from style if not directly on paragraph
        if outline_lvl is None and style_id:
            outline_lvl = _resolve_outline_level(style_id, styles)

        # Resolve numPr from style if not directly on paragraph
        if (num_id is None or num_id == "0") and style_id:
            style_numpr = _resolve_numpr(style_id, styles)
            if style_numpr:
                num_id, style_ilvl = style_numpr
                if ilvl is None:
                    ilvl = style_ilvl

        if ilvl is None:
            ilvl = 0

        # Convert outline level 0-8 → heading level 1-6
        heading_level = None
        if outline_lvl is not None and outline_lvl <= 5:
            heading_level = outline_lvl + 1  # outlineLvl 0 = H1, 1 = H2, etc.

        # Compute numbering prefix
        numbering_prefix = None
        if num_id and num_id != "0":
            abstract_id = num_to_abstract.get(num_id)
            if abstract_id and abstract_id in abstract_nums:
                levels = abstract_nums[abstract_id]
                if ilvl in levels:
                    level_def = levels[ilvl]

                    effective_start = level_def["start"]
                    if num_id in num_overrides and ilvl in num_overrides[num_id]:
                        effective_start = num_overrides[num_id][ilvl]["start"]

                    if num_id not in counters:
                        counters[num_id] = {}

                    if ilvl not in counters[num_id]:
                        counters[num_id][ilvl] = effective_start
                    else:
                        counters[num_id][ilvl] += 1

                    # Reset lower-level counters
                    for other_ilvl in levels:
                        if other_ilvl > ilvl and other_ilvl in counters.get(num_id, {}):
                            other_start = levels[other_ilvl]["start"]
                            if num_id in num_overrides and other_ilvl in num_overrides[num_id]:
                                other_start = num_overrides[num_id][other_ilvl]["start"]
                            counters[num_id][other_ilvl] = other_start - 1

                    numbering_prefix = _format_level_text(
                        level_def["lvlText"], levels, counters[num_id], ilvl,
                    )

        results.append(ParagraphInfo(
            text=para_text,
            heading_level=heading_level,
            numbering_prefix=numbering_prefix,
        ))

    return results


def build_numbering_map(docx_path: str | Path) -> list[tuple[str, Optional[str]]]:
    """Build ordered list of (text, numbering_prefix) for paragraphs with numbering."""
    entries = extract_paragraph_info(docx_path)
    return [(e.text, e.numbering_prefix) for e in entries if e.numbering_prefix]


def build_heading_map(docx_path: str | Path) -> list[ParagraphInfo]:
    """Build ordered list of ParagraphInfo for all non-empty paragraphs.

    This is the primary API used by word2md.py to fix both heading levels
    and numbering in the mammoth-generated HTML.
    """
    return extract_paragraph_info(docx_path)
