"""Word (.docx) to Markdown converter.

Directly parses .docx XML to build Markdown. Heading levels come from the
outlineLvl property, auto-numbering from numPr — both read per-paragraph
from the XML, with no text matching or style-name guessing.
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Optional

from converter.numbering import (
    _attr, _parse_abstract_nums, _parse_num_mappings, _parse_num_overrides,
    _parse_styles, _resolve_outline_level, _resolve_numpr,
    _format_level_text, _build_pstyle_map, _get_effective_levels,
    NS, WORD_NS,
)

# Additional XML namespaces for images and hyperlinks
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
VML_NS = "urn:schemas-microsoft-com:vml"
O_NS = "urn:schemas-microsoft-com:office:office"


def _local_tag(element) -> str:
    """Get the local tag name without namespace."""
    tag = element.tag
    return tag.split("}")[-1] if "}" in tag else tag


class _DocxConverter:
    """Walks through a .docx file's XML and builds Markdown directly.

    For each paragraph, heading level and numbering are read from the XML
    properties of THAT paragraph (outlineLvl, numPr), not from text matching.
    """

    # toc_mode values:
    #   "none"                   – keep everything
    #   "toc_only"               – remove only the TOC paragraphs
    #   "before_toc"             – remove TOC and everything before it
    #   "before_toc_keep_abstract"– remove TOC + before, but keep abstract sections

    def __init__(self, docx_path: Path, output_dir: Path, extract_images: bool,
                 skip_cover: bool = False, toc_mode: str = "none",
                 table_cell_break: str = "space"):
        self.output_dir = output_dir
        self.extract_images = extract_images
        self.skip_cover = skip_cover
        self.toc_mode = toc_mode
        self.table_cell_break = table_cell_break
        self.image_index = 1
        self.counters: dict[str, dict[int, int]] = {}  # numbering state

        with zipfile.ZipFile(docx_path, "r") as z:
            namelist = z.namelist()

            self.doc_xml = ET.fromstring(z.read("word/document.xml"))

            # Styles (outlineLvl + numPr + basedOn inheritance)
            self.styles: dict = {}
            if "word/styles.xml" in namelist:
                self.styles = _parse_styles(ET.fromstring(z.read("word/styles.xml")))

            # Numbering definitions
            self.abstract_nums: dict = {}
            self.num_to_abstract: dict = {}
            self.num_overrides: dict = {}
            self.pstyle_map: dict = {}
            if "word/numbering.xml" in namelist:
                num_xml = ET.fromstring(z.read("word/numbering.xml"))
                self.abstract_nums = _parse_abstract_nums(num_xml)
                self.num_to_abstract = _parse_num_mappings(num_xml)
                self.num_overrides = _parse_num_overrides(num_xml)
                self.pstyle_map = _build_pstyle_map(
                    self.abstract_nums, self.num_to_abstract, self.num_overrides,
                )

            # Relationships (images, hyperlinks)
            self.rels: dict[str, dict] = {}
            rels_path = "word/_rels/document.xml.rels"
            if rels_path in namelist:
                for rel in ET.fromstring(z.read(rels_path)):
                    rid = rel.get("Id")
                    if rid:
                        self.rels[rid] = {
                            "target": rel.get("Target", ""),
                            "type": rel.get("Type", ""),
                        }

            # Pre-read all media files
            self.media: dict[str, bytes] = {}
            for name in namelist:
                if name.startswith("word/media/"):
                    self.media[name] = z.read(name)

    # ── Public API ──

    def convert(self) -> str:
        body = self.doc_xml.find("w:body", NS)
        if body is None:
            return ""

        children = list(body)
        children = self._apply_filters(children)

        parts: list[str] = []
        for child in children:
            tag = _local_tag(child)
            if tag == "p":
                md = self._convert_paragraph(child)
                if md is not None:
                    parts.append(md)
            elif tag == "tbl":
                md = self._convert_table(child)
                if md:
                    parts.append(md)
            elif tag == "sdt":
                # Structured Document Tag — may contain paragraphs/tables
                if self._is_toc_sdt(child):
                    continue  # TOC SDTs already handled by filter
                for inner in child.iter(f"{{{WORD_NS}}}p"):
                    md = self._convert_paragraph(inner)
                    if md is not None:
                        parts.append(md)
            # sectPr, bookmarkStart, etc. → skip

        return "\n\n".join(parts)

    # ── Filtering logic ──

    def _apply_filters(self, children: list) -> list:
        """Apply skip_cover and toc_mode filters to the element list."""
        result = children

        # ❶ Skip cover page (content before first page break)
        if self.skip_cover:
            cover_end = self._find_first_page_break(result)
            if cover_end >= 0:
                result = result[cover_end + 1:]

        # ❷ TOC filtering
        if self.toc_mode == "toc_only":
            # Remove only TOC elements, keep everything else
            result = [c for c in result if not self._is_toc_element(c)]

        elif self.toc_mode == "before_toc":
            # Remove TOC and everything before it
            toc_end = self._find_toc_end(result)
            if toc_end >= 0:
                result = result[toc_end + 1:]

        elif self.toc_mode == "before_toc_keep_abstract":
            # Remove TOC and everything before it, but keep abstract sections
            toc_end = self._find_toc_end(result)
            if toc_end >= 0:
                abstracts = self._extract_abstract_elements(result[:toc_end + 1])
                result = abstracts + result[toc_end + 1:]

        return result

    # ── Cover page detection ──

    @staticmethod
    def _find_first_page_break(children) -> int:
        """Find the index of the element containing the first page break.

        Detects:
        - <w:br w:type="page"/> (explicit page break in a run)
        - <w:sectPr> with <w:type w:val="nextPage"/> inside <w:pPr> (section break)

        Returns -1 if no page break is found.
        """
        for i, child in enumerate(children):
            tag = _local_tag(child)
            if tag == "p":
                # Explicit page break: <w:br w:type="page"/>
                for br in child.iter(f"{{{WORD_NS}}}br"):
                    br_type = _attr(br, "type")
                    if br_type == "page":
                        return i
                # Section break in paragraph properties
                pPr = child.find("w:pPr", NS)
                if pPr is not None:
                    sectPr = pPr.find("w:sectPr", NS)
                    if sectPr is not None:
                        sect_type = sectPr.find("w:type", NS)
                        if sect_type is not None:
                            val = _attr(sect_type, "val")
                            if val in ("nextPage", "oddPage", "evenPage"):
                                return i
                        else:
                            # No type element = default = nextPage
                            return i
        return -1

    # ── TOC detection ──

    @staticmethod
    def _is_toc_sdt(sdt) -> bool:
        """Check if a <w:sdt> element is a Table of Contents block."""
        for gallery in sdt.iter(f"{{{WORD_NS}}}docPartGallery"):
            val = _attr(gallery, "val")
            if val and "Contents" in val:
                return True
        return False

    @staticmethod
    def _is_toc_paragraph(para) -> bool:
        """Check if a paragraph belongs to the TOC (by style ID)."""
        pPr = para.find("w:pPr", NS)
        if pPr is None:
            return False
        style_el = pPr.find("w:pStyle", NS)
        if style_el is None:
            return False
        style_id = _attr(style_el, "val") or ""
        if re.match(r"^TOC\d{0,2}$", style_id, re.IGNORECASE):
            return True
        if style_id.lower() in ("tocheading", "toc heading", "tableofcontents"):
            return True
        return False

    @staticmethod
    def _has_toc_field(para) -> bool:
        """Check if a paragraph contains a TOC field instruction."""
        for instr in para.iter(f"{{{WORD_NS}}}instrText"):
            if instr.text and "TOC" in instr.text.upper():
                return True
        return False

    @classmethod
    def _is_toc_element(cls, child) -> bool:
        """Check if an element is any kind of TOC content."""
        tag = _local_tag(child)
        if tag == "sdt" and cls._is_toc_sdt(child):
            return True
        if tag == "p":
            if cls._is_toc_paragraph(child):
                return True
            if cls._has_toc_field(child):
                return True
        return False

    @classmethod
    def _find_toc_end(cls, children) -> int:
        """Find the index of the last TOC-related element.

        Returns -1 if no TOC is found.
        """
        last_toc_idx = -1
        for i, child in enumerate(children):
            if cls._is_toc_element(child):
                last_toc_idx = i
        return last_toc_idx

    # ── Abstract detection ──

    @staticmethod
    def _get_paragraph_text(para) -> str:
        """Get plain text from a paragraph element."""
        parts = []
        for t in para.iter(f"{{{WORD_NS}}}t"):
            if t.text:
                parts.append(t.text)
        return "".join(parts).strip()

    @classmethod
    def _is_abstract_heading(cls, para) -> bool:
        """Check if a paragraph is an abstract heading.

        Matches headings containing '摘要', 'Abstract', or common variations.
        """
        # Must be a heading (has outlineLvl or heading style)
        pPr = para.find("w:pPr", NS)
        style_id = None
        is_heading = False

        if pPr is not None:
            style_el = pPr.find("w:pStyle", NS)
            if style_el is not None:
                style_id = (_attr(style_el, "val") or "").lower()
            if pPr.find("w:outlineLvl", NS) is not None:
                is_heading = True

        # Heading styles
        if style_id and re.match(r"^(heading|\u6807\u9898)\s*\d*$", style_id, re.IGNORECASE):
            is_heading = True

        text = cls._get_paragraph_text(para).strip()
        if not text:
            return False

        text_lower = text.lower()
        # Match abstract keywords
        abstract_keywords = ["摘\u8981", "abstract", "摘 要", "内容摘要"]
        has_keyword = any(kw in text_lower for kw in abstract_keywords)

        if not has_keyword:
            return False

        # If it has heading formatting or is short enough to be a title line
        if is_heading:
            return True
        # Even without outlineLvl, a short line that is just the keyword is likely a heading
        if len(text) <= 30:
            return True

        return False

    @classmethod
    def _extract_abstract_elements(cls, children: list) -> list:
        """Extract abstract sections (heading + body) from a list of elements.

        Finds each abstract heading and collects all elements until the next
        heading-level element, page break, or TOC element.
        """
        result = []
        i = 0
        while i < len(children):
            child = children[i]
            tag = _local_tag(child)

            if tag == "p" and cls._is_abstract_heading(child):
                # Found an abstract heading — collect it and following body paragraphs
                result.append(child)
                i += 1
                while i < len(children):
                    next_child = children[i]
                    next_tag = _local_tag(next_child)

                    # Stop at TOC elements
                    if cls._is_toc_element(next_child):
                        break

                    # Stop at next heading
                    if next_tag == "p":
                        pPr = next_child.find("w:pPr", NS)
                        if pPr is not None and pPr.find("w:outlineLvl", NS) is not None:
                            # Check if this is another abstract heading
                            if cls._is_abstract_heading(next_child):
                                break  # will be picked up by outer loop
                            else:
                                break  # non-abstract heading = section end
                        # Check style-based headings
                        if pPr is not None:
                            style_el = pPr.find("w:pStyle", NS)
                            if style_el is not None:
                                sid = (_attr(style_el, "val") or "").lower()
                                if re.match(r"^(heading|\u6807\u9898)\s*\d+$", sid):
                                    if cls._is_abstract_heading(next_child):
                                        break
                                    else:
                                        break

                    # Stop at page breaks
                    if next_tag == "p":
                        has_page_break = False
                        for br in next_child.iter(f"{{{WORD_NS}}}br"):
                            if _attr(br, "type") == "page":
                                has_page_break = True
                                break
                        if has_page_break:
                            # Include this paragraph (it may have text before the break)
                            result.append(next_child)
                            i += 1
                            break

                    result.append(next_child)
                    i += 1
            else:
                i += 1

        return result

    # ── Paragraph conversion ──

    def _convert_paragraph(self, para) -> Optional[str]:
        pPr = para.find("w:pPr", NS)
        style_id = None

        if pPr is not None:
            el = pPr.find("w:pStyle", NS)
            if el is not None:
                style_id = _attr(el, "val")

        # ➊ Heading level — from outlineLvl property (paragraph or style)
        heading_level = self._get_heading_level(pPr, style_id)

        # ➋ Numbering — from numPr property (paragraph or style)
        num_id, ilvl = self._get_numpr(pPr, style_id)
        numbering_prefix = self._compute_numbering(num_id, ilvl, is_heading=bool(heading_level))

        # ➌ Formatted text from runs
        text = self._runs_to_md(para)
        if not text.strip():
            return None

        text = text.strip()

        # ➍ Assemble Markdown
        if heading_level:
            hashes = "#" * heading_level
            prefix = f"{numbering_prefix} " if numbering_prefix else ""
            # Extract images from heading text — images break Markdown headings
            images = re.findall(r"!\[[^\]]*\]\([^)]*\)", text)
            if images:
                clean = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", text).strip()
                return f"{hashes} {prefix}{clean}\n\n" + "\n\n".join(images)
            return f"{hashes} {prefix}{text}"

        if numbering_prefix and num_id:
            # Non-heading paragraph with numbering → list item
            num_fmt = self._get_num_fmt(num_id, ilvl)
            if num_fmt == "bullet":
                return f"- {text}"
            else:
                return f"{numbering_prefix} {text}"

        return text

    def _get_heading_level(self, pPr, style_id: Optional[str]) -> Optional[int]:
        """Read outlineLvl from the paragraph or its style chain. Returns 1-6."""
        outline_lvl = None

        # Direct outlineLvl on paragraph
        if pPr is not None:
            el = pPr.find("w:outlineLvl", NS)
            if el is not None:
                val = _attr(el, "val")
                if val is not None:
                    lvl = int(val)
                    if 0 <= lvl <= 8:
                        outline_lvl = lvl

        # Fall back to style (follows basedOn inheritance)
        if outline_lvl is None and style_id:
            outline_lvl = _resolve_outline_level(style_id, self.styles)

        if outline_lvl is not None and outline_lvl <= 5:
            return outline_lvl + 1   # outlineLvl 0 → H1
        return None

    def _get_numpr(self, pPr, style_id: Optional[str]) -> tuple[Optional[str], int]:
        """Read numId + ilvl from the paragraph or its style chain.

        Resolution order:
        1. Direct numPr on the paragraph
           - If numId=0 or missing: numId is unset
           - If numId>0 with ilvl: use both
           - If numId>0 without ilvl: use default ilvl=0
        2. If paragraph has NO numPr element at all: try style inheritance
        3. pStyle mapping (numbering levels linked to styles via <w:pStyle>)
        4. Style numPr inheritance (following basedOn chain)

        IMPORTANT: If a paragraph has an explicit numPr with numId="0", this
        means "do not number this paragraph" and we should NOT fall back to
        the style's numId. Only inherit if there is NO numPr element at all.
        """
        num_id = None
        ilvl = 0
        para_has_numpr = False  # New flag: paragraph has numPr element
        para_has_ilvl = False

        if pPr is not None:
            numpr = pPr.find("w:numPr", NS)
            if numpr is not None:
                para_has_numpr = True  # Mark that paragraph explicitly set numPr
                el = numpr.find("w:numId", NS)
                if el is not None:
                    num_id = _attr(el, "val")
                el = numpr.find("w:ilvl", NS)
                if el is not None:
                    ilvl = int(_attr(el, "val") or "0")
                    para_has_ilvl = True

        # Only inherit from style if paragraph has NO numPr element at all
        # (not if it has numPr with numId=0, which explicitly disables numbering)
        if not para_has_numpr and style_id:
            if style_id in self.pstyle_map:
                ps_num_id, ps_ilvl = self.pstyle_map[style_id]
                num_id = ps_num_id
                ilvl = ps_ilvl
            else:
                # Fall back to style inheritance
                result = _resolve_numpr(style_id, self.styles)
                if result:
                    style_num_id, style_ilvl = result
                    num_id = style_num_id
                    if not para_has_ilvl:
                        ilvl = style_ilvl

        return num_id, ilvl

    def _compute_numbering(self, num_id: Optional[str], ilvl: int,
                           is_heading: bool = False) -> str:
        """Compute rendered numbering prefix from the numbering definitions.

        Uses effective levels (abstract + overrides merged) so that level
        redefinitions in <w:lvlOverride> (including isLgl, format changes)
        are properly applied.

        For heading paragraphs, counters are shared across all numIds that
        reference the same abstractNumId. This matches Word's behavior where
        heading styles may use different numId values but share a single
        multi-level numbering chain.
        """
        if not num_id or num_id == "0":
            return ""

        abstract_id = self.num_to_abstract.get(num_id)
        if not abstract_id or abstract_id not in self.abstract_nums:
            return ""

        # Get effective levels: abstract definition merged with overrides
        levels = _get_effective_levels(
            num_id, self.abstract_nums, self.num_to_abstract, self.num_overrides,
        )
        if ilvl not in levels:
            return ""

        level_def = levels[ilvl]

        # Start value from effective level definition
        start = level_def.get("start", 1)

        # For headings, share counters across all numIds with the same
        # abstractNumId so that chapter/section numbering is coherent.
        counter_key = ("heading", abstract_id) if is_heading else ("list", num_id)

        if counter_key not in self.counters:
            self.counters[counter_key] = {}

        if ilvl not in self.counters[counter_key]:
            self.counters[counter_key][ilvl] = start
        else:
            self.counters[counter_key][ilvl] += 1

        # Reset sub-level counters
        for other_ilvl in levels:
            if other_ilvl > ilvl and other_ilvl in self.counters.get(counter_key, {}):
                other_start = levels[other_ilvl].get("start", 1)
                self.counters[counter_key][other_ilvl] = other_start - 1

        return _format_level_text(level_def["lvlText"], levels, self.counters[counter_key], ilvl)

    def _get_num_fmt(self, num_id: str, ilvl: int) -> str:
        abstract_id = self.num_to_abstract.get(num_id, "")
        levels = self.abstract_nums.get(abstract_id, {})
        return levels.get(ilvl, {}).get("numFmt", "decimal")

    # ── Run / inline content conversion ──

    def _runs_to_md(self, para) -> str:
        """Convert all inline children of a <w:p> to Markdown text."""
        parts: list[str] = []
        anchored_images: list[str] = []

        for child in para:
            tag = _local_tag(child)
            if tag == "r":
                parts.append(self._convert_run(child, anchored_images))
            elif tag == "hyperlink":
                parts.append(self._convert_hyperlink(child))
            elif tag in ("fldSimple", "smartTag", "sdt"):
                # Wrapped structures — extract runs from inside
                for r in child.iter(f"{{{WORD_NS}}}r"):
                    parts.append(self._convert_run(r, anchored_images))
            # pPr, bookmarkStart, proofErr, etc. → skip

        text = "".join(parts)

        # Floating/anchored images should be rendered after paragraph text,
        # matching typical Word reading flow.
        if anchored_images:
            text = text.rstrip()
            if text:
                text += "\n\n" + "\n\n".join(anchored_images)
            else:
                text = "\n\n".join(anchored_images)

        # Merge adjacent bold+italic spans:
        # **text1****text2** → **text1text2**
        # ***text1******text2*** → ***text1text2***
        text = text.replace("******", "")   # bold+italic merge
        text = text.replace("****", "")     # bold merge

        return text

    def _convert_run(self, run, anchored_images: Optional[list[str]] = None) -> str:
        """Convert a <w:r> element to Markdown text with formatting."""
        rPr = run.find("w:rPr", NS)
        is_bold = False
        is_italic = False
        is_strike = False
        is_code = False

        if rPr is not None:
            is_bold = self._check_toggle(rPr, "b")
            is_italic = self._check_toggle(rPr, "i")
            is_strike = self._check_toggle(rPr, "strike")

            # Monospace font → inline code
            fonts = rPr.find("w:rFonts", NS)
            if fonts is not None:
                ascii_font = (_attr(fonts, "ascii") or "").lower()
                if any(f in ascii_font for f in ("consolas", "courier", "mono", "menlo")):
                    is_code = True

        # Process children in document order
        text_parts: list[str] = []
        image_parts: list[str] = []

        for child in run:
            child_tag = _local_tag(child)
            if child_tag == "t":
                text_parts.append(child.text or "")
            elif child_tag == "br":
                br_type = _attr(child, "type")
                text_parts.append("\n\n---\n\n" if br_type == "page" else "\n")
            elif child_tag == "cr":
                text_parts.append("\n")
            elif child_tag == "tab":
                text_parts.append("  ")
            elif child_tag == "noBreakHyphen":
                text_parts.append("-")
            elif child_tag == "drawing":
                img = self._convert_drawing(child)
                if img:
                    is_anchored = child.find(f".//{{{WP_NS}}}anchor") is not None
                    if is_anchored and anchored_images is not None:
                        anchored_images.append(img)
                    else:
                        image_parts.append(img)
            elif child_tag == "pict":
                img = self._convert_pict(child)
                if img:
                    if anchored_images is not None:
                        anchored_images.append(img)
                    else:
                        image_parts.append(img)
            elif child_tag == "AlternateContent":
                # Some Word docs wrap drawing/pict inside mc:AlternateContent.
                for nested in child.iter():
                    nested_tag = _local_tag(nested)
                    if nested_tag == "drawing":
                        img = self._convert_drawing(nested)
                        if img:
                            is_anchored = nested.find(f".//{{{WP_NS}}}anchor") is not None
                            if is_anchored and anchored_images is not None:
                                anchored_images.append(img)
                            else:
                                image_parts.append(img)
                            break
                    elif nested_tag == "pict":
                        img = self._convert_pict(nested)
                        if img:
                            if anchored_images is not None:
                                anchored_images.append(img)
                            else:
                                image_parts.append(img)
                            break

        # Images returned directly (no bold/italic wrapper)
        if image_parts:
            prefix = "".join(text_parts)
            return (prefix + "".join(image_parts)) if prefix.strip() else "".join(image_parts)

        text = "".join(text_parts)
        if not text:
            return ""

        # Apply character formatting
        if text.strip():
            if is_code:
                text = f"`{text.strip()}`"
            else:
                stripped = text.strip()
                leading = text[: len(text) - len(text.lstrip())]
                trailing = text[len(text.rstrip()) :]

                if is_bold and is_italic:
                    stripped = f"***{stripped}***"
                elif is_bold:
                    stripped = f"**{stripped}**"
                elif is_italic:
                    stripped = f"*{stripped}*"
                if is_strike:
                    stripped = f"~~{stripped}~~"

                text = leading + stripped + trailing

        return text

    @staticmethod
    def _check_toggle(rPr, prop_name: str) -> bool:
        """Check a toggle property like <w:b/> or <w:b w:val='0'/>."""
        el = rPr.find(f"w:{prop_name}", NS)
        if el is None:
            return False
        val = _attr(el, "val")
        return val not in ("0", "false")

    def _convert_hyperlink(self, hyperlink) -> str:
        rid = hyperlink.get(f"{{{R_NS}}}id")
        url = ""
        if rid and rid in self.rels:
            url = self.rels[rid]["target"]

        anchor = _attr(hyperlink, "anchor")
        if anchor and not url:
            url = f"#{anchor}"

        parts = []
        for run in hyperlink.findall("w:r", NS):
            parts.append(self._convert_run(run))
        text = "".join(parts)

        return f"[{text}]({url})" if url else text

    def _convert_drawing(self, drawing) -> Optional[str]:
        """Extract an image from a <w:drawing> element."""
        blip = drawing.find(f".//{{{DRAWING_NS}}}blip")
        if blip is None:
            return None

        embed = blip.get(f"{{{R_NS}}}embed")
        if not embed:
            return None

        alt = ""
        docPr = drawing.find(f".//{{{WP_NS}}}docPr")
        if docPr is not None:
            alt = docPr.get("descr", "") or docPr.get("name", "")

        return self._convert_image_by_rel_id(embed, alt)

    def _convert_pict(self, pict) -> Optional[str]:
        """Extract an image from a legacy <w:pict>/<v:imagedata> element."""
        imagedata = pict.find(f".//{{{VML_NS}}}imagedata")
        if imagedata is None:
            return None

        rid = imagedata.get(f"{{{R_NS}}}id")
        if not rid:
            return None

        alt = (
            imagedata.get("title", "")
            or imagedata.get(f"{{{O_NS}}}title", "")
            or imagedata.get("alt", "")
        )
        return self._convert_image_by_rel_id(rid, alt)

    def _convert_image_by_rel_id(self, rel_id: str, alt: str = "") -> Optional[str]:
        """Resolve relationship id and emit image markdown with extracted asset."""
        if rel_id not in self.rels:
            return None

        image_target = self.rels[rel_id]["target"]
        if not image_target.startswith("word/"):
            image_target = f"word/{image_target}"

        if not self.extract_images or image_target not in self.media:
            return "![image]()"

        images_dir = self.output_dir / "images"
        images_dir.mkdir(parents=True, exist_ok=True)

        ext = Path(image_target).suffix or ".png"
        filename = f"image_{self.image_index:03d}{ext}"
        self.image_index += 1
        (images_dir / filename).write_bytes(self.media[image_target])

        # Clean up Office AI-generated alt text (e.g. "文本, 信件\n\nAI 生成的内容可能不正确。")
        if alt:
            # Remove lines containing AI-generated disclaimer
            lines = [ln.strip() for ln in alt.splitlines()
                     if ln.strip() and "AI 生成的内容可能不正确" not in ln]
            alt = " ".join(lines)
            # Fall back to generic label if nothing meaningful remains
            if not alt or all(c in ", 、，" for c in alt.replace(" ", "")):
                alt = f"图片 {self.image_index - 1}"

        return f"![{alt}](images/{filename})"

    # ── Table conversion ──

    def _convert_table(self, tbl) -> str:
        rows: list[list[str]] = []
        for tr in tbl.findall("w:tr", NS):
            cells = []
            for tc in tr.findall("w:tc", NS):
                cell_parts = []
                for p in tc.findall("w:p", NS):
                    t = self._runs_to_md(p).strip()
                    if t:
                        cell_parts.append(t)
                cell_text = " ".join(cell_parts)
                # Markdown tables require each row on a single line.
                # Strip page-break markers first, then handle soft breaks.
                cell_text = cell_text.replace("\n\n---\n\n", " ")
                if self.table_cell_break == "br":
                    cell_text = cell_text.replace("\n", "<br>")
                else:
                    cell_text = cell_text.replace("\n", " ")
                cell_text = cell_text.replace("|", "\\|")
                cells.append(cell_text)
            if cells:
                rows.append(cells)

        if not rows:
            return ""

        max_cols = max(len(r) for r in rows)
        for r in rows:
            while len(r) < max_cols:
                r.append("")

        lines = [
            "| " + " | ".join(rows[0]) + " |",
            "| " + " | ".join(["---"] * max_cols) + " |",
        ]
        for row in rows[1:]:
            lines.append("| " + " | ".join(row) + " |")

        return "\n".join(lines)


# ── Public API ──

def convert_word_to_markdown(
    input_path: str | Path,
    output_path: Optional[str | Path] = None,
    extract_images: bool = True,
    skip_cover: bool = False,
    toc_mode: str = "none",
    table_cell_break: str = "space",
) -> str:
    """Convert a .docx file to Markdown.

    Heading levels and auto-numbering are read directly from each paragraph's
    XML properties (outlineLvl, numPr). No text matching or style-name
    guessing is involved.

    Args:
        input_path: Path to the .docx file.
        output_path: Optional .md output path.
        extract_images: Extract images to images/ directory.
        skip_cover: Remove the first page (cover page).
        toc_mode: How to handle the Table of Contents:
            "none"  – keep everything (default)
            "toc_only" – remove only the TOC
            "before_toc" – remove TOC and everything before it
            "before_toc_keep_abstract" – remove TOC + before, keep abstracts
        table_cell_break: How to handle line breaks inside table cells:
            "space" – replace with spaces (default, compatible with all renderers)
            "br"    – replace with <br> tags (preserves line breaks, but some
                       renderers may display the raw tag as text)

    Returns:
        The Markdown string.
    """
    input_path = Path(input_path)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    if input_path.suffix.lower() != ".docx":
        raise ValueError(f"Unsupported format: {input_path.suffix}")

    output_dir = Path(output_path).parent if output_path else input_path.parent

    converter = _DocxConverter(
        input_path, output_dir, extract_images,
        skip_cover=skip_cover, toc_mode=toc_mode,
        table_cell_break=table_cell_break,
    )
    markdown = converter.convert()

    # Clean up
    markdown = re.sub(r"\n{3,}", "\n\n", markdown)
    markdown = markdown.strip() + "\n"

    if output_path:
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(markdown, encoding="utf-8")

    return markdown
