import io
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, BinaryIO, Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

from markitdown import DocumentConverter, DocumentConverterResult
from markitdown._stream_info import StreamInfo

try:
    from markitdown.converter_utils.docx.pre_process import pre_process_docx
except Exception:  # pragma: no cover - only used when MarkItDown internals change
    pre_process_docx = None


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W_NS}

ACCEPTED_MIME_TYPE_PREFIXES = [
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
]

ACCEPTED_FILE_EXTENSIONS = [".docx"]


def w_tag(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def w_attr(element: ET.Element, name: str, default: Optional[str] = None) -> Optional[str]:
    return element.attrib.get(w_tag(name), default)


def r_attr(element: ET.Element, name: str, default: Optional[str] = None) -> Optional[str]:
    return element.attrib.get(f"{{{R_NS}}}{name}", default)


def clean_text(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = re.sub(r"^[]\s*", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def escape_markdown_cell(text: str) -> str:
    text = clean_text(text)
    text = text.replace("\\", "\\\\").replace("|", "\\|")
    return re.sub(r"\s*\n\s*", "<br>", text)


def normalize_inline_markdown(text: str) -> str:
    text = re.sub(r"\*\*(!\[[^\]]*\]\([^)]+\))\*\*", r"\1", text)
    text = re.sub(r"\[DELETED: ([^\]]*)\]\[DELETED: ", r"[DELETED: \1", text)
    previous = None
    while previous != text:
        previous = text
        text = re.sub(r"\*\*([^*\n]+)\*\*\*\*([^*\n]+)\*\*", r"**\1\2**", text)
    return text


def is_code_like_text(text: str) -> bool:
    if len(text) < 80:
        return False

    signals = 0
    if text.count(";") >= 2:
        signals += 1
    if any(token in text for token in ["//", "/*", "*/", "(", ")", "="]):
        signals += 1
    if any(
        keyword in text
        for keyword in ["new ", "Map<", "Request", "config.", "System.out", "public ", "class "]
    ):
        signals += 1

    return signals >= 2


@dataclass
class Comment:
    marker: str
    author: str
    date: str
    text: str
    reply_to: Optional[str] = None

    def render(self, target: Optional[str] = None) -> str:
        pieces: List[str] = [self.marker]
        if self.reply_to:
            pieces.append(f"回复{self.reply_to}")
        if target:
            pieces.append(f"针对{target}")
        prefix = " ".join(pieces)
        byline = self.author
        if self.date:
            byline = f"{byline} {self.date}" if byline else self.date
        body = f"{byline}: {self.text}" if byline else self.text
        return f"[{prefix}: {body}]"


@dataclass
class NumberingLevel:
    num_format: str
    level_text: str
    start: int


@dataclass
class RunToken:
    text: str
    bold: bool = False
    strike: bool = False


class ReadableDocxConverter(DocumentConverter):
    """
    Converts DOCX files to Markdown while preserving readable Word semantics.
    """

    def accepts(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> bool:
        extension = (stream_info.extension or "").lower()
        mimetype = (stream_info.mimetype or "").lower()

        if extension in ACCEPTED_FILE_EXTENSIONS:
            return True

        for prefix in ACCEPTED_MIME_TYPE_PREFIXES:
            if mimetype.startswith(prefix):
                return True

        return False

    def convert(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> DocumentConverterResult:
        data = file_stream.read()
        if not data:
            raise ValueError("Empty DOCX stream")

        if pre_process_docx is not None:
            try:
                processed = pre_process_docx(io.BytesIO(data))
                data = processed.read()
            except Exception:
                pass

        markdown = ReadableDocxMarkdownBuilder(
            data,
            image_output_dir=kwargs.get("image_output_dir"),
            image_output_prefix=kwargs.get("image_output_prefix", "./images"),
        ).convert()
        return DocumentConverterResult(markdown=markdown)


class ReadableDocxMarkdownBuilder:
    def __init__(
        self,
        data: bytes,
        *,
        image_output_dir: Optional[str] = None,
        image_output_prefix: str = "./images",
    ):
        self.data = data
        self.comments: Dict[str, Comment] = {}
        self.numbering: Dict[str, Dict[str, NumberingLevel]] = {}
        self.numbering_counters: Dict[str, int] = {}
        self.image_output_dir = Path(image_output_dir) if image_output_dir else None
        self.image_output_prefix = image_output_prefix.rstrip("/")
        self.image_rels: Dict[str, str] = {}
        self.image_blobs: Dict[str, bytes] = {}
        self.image_extensions: Dict[str, str] = {}
        self.image_paths: Dict[str, str] = {}
        self.image_index = 1
        self.hyperlinks: Dict[str, str] = {}
        self.style_heading_levels: Dict[str, int] = {}

    def convert(self) -> str:
        with zipfile.ZipFile(io.BytesIO(self.data), mode="r") as docx:
            self.comments = self._read_comments(docx)
            self.numbering = self._read_numbering(docx)
            self.style_heading_levels = self._read_style_heading_levels(docx)
            self.image_rels, self.hyperlinks = self._read_document_relationships(docx)
            self._read_images(docx)
            document_xml = docx.read("word/document.xml")

        document = ET.fromstring(document_xml)
        body = document.find("w:body", NS)
        if body is None:
            return ""

        blocks: List[str] = []
        table_index = self._append_blocks(body, blocks, 1)

        return "\n\n".join(blocks).strip()

    def _append_blocks(self, parent: ET.Element, blocks: List[str], table_index: int) -> int:
        for child in parent:
            if child.tag == w_tag("p"):
                paragraph = self._paragraph_to_markdown(child)
                if paragraph:
                    blocks.append(paragraph)
            elif child.tag == w_tag("tbl"):
                table = self._table_to_markdown(child, table_index)
                if table:
                    blocks.append(table)
                    table_index += 1
            elif child.tag == w_tag("sdt"):
                content = child.find("w:sdtContent", NS)
                if content is not None:
                    table_index = self._append_blocks(content, blocks, table_index)
        return table_index

    def _read_document_relationships(self, docx: zipfile.ZipFile) -> Tuple[Dict[str, str], Dict[str, str]]:
        rels_path = "word/_rels/document.xml.rels"
        if rels_path not in docx.namelist():
            return {}, {}

        root = ET.fromstring(docx.read(rels_path))
        image_rels: Dict[str, str] = {}
        hyperlinks: Dict[str, str] = {}
        for rel in root.findall(f"{{{REL_NS}}}Relationship"):
            rel_id = rel.attrib.get("Id")
            rel_type = rel.attrib.get("Type", "")
            target = rel.attrib.get("Target")
            if not rel_id or not target:
                continue
            if target.startswith("media/"):
                image_rels[rel_id] = f"word/{target}"
            elif rel_type.endswith("/hyperlink"):
                hyperlinks[rel_id] = target
        return image_rels, hyperlinks

    def _read_images(self, docx: zipfile.ZipFile) -> None:
        for rel_id, zip_path in self.image_rels.items():
            if zip_path not in docx.namelist():
                continue
            self.image_blobs[rel_id] = docx.read(zip_path)
            self.image_extensions[rel_id] = Path(zip_path).suffix or ".jpg"

    def _read_numbering(self, docx: zipfile.ZipFile) -> Dict[str, Dict[str, NumberingLevel]]:
        if "word/numbering.xml" not in docx.namelist():
            return {}

        root = ET.fromstring(docx.read("word/numbering.xml"))
        abstract_levels: Dict[str, Dict[str, NumberingLevel]] = {}
        for abstract_num in root.findall("w:abstractNum", NS):
            abstract_num_id = w_attr(abstract_num, "abstractNumId")
            if abstract_num_id is None:
                continue
            levels: Dict[str, NumberingLevel] = {}
            for level in abstract_num.findall("w:lvl", NS):
                ilvl = w_attr(level, "ilvl", "0") or "0"
                num_format = level.find("w:numFmt", NS)
                level_text = level.find("w:lvlText", NS)
                start = level.find("w:start", NS)
                start_value = w_attr(start, "val", "1") if start is not None else "1"
                levels[ilvl] = NumberingLevel(
                    num_format=w_attr(num_format, "val", "") if num_format is not None else "",
                    level_text=w_attr(level_text, "val", "") if level_text is not None else "",
                    start=int(start_value) if start_value and start_value.isdigit() else 1,
                )
            abstract_levels[abstract_num_id] = levels

        numbering: Dict[str, Dict[str, NumberingLevel]] = {}
        for num in root.findall("w:num", NS):
            num_id = w_attr(num, "numId")
            abstract_ref = num.find("w:abstractNumId", NS)
            abstract_num_id = w_attr(abstract_ref, "val") if abstract_ref is not None else None
            if num_id is not None and abstract_num_id in abstract_levels:
                numbering[num_id] = abstract_levels[abstract_num_id]
        return numbering

    def _read_style_heading_levels(self, docx: zipfile.ZipFile) -> Dict[str, int]:
        if "word/styles.xml" not in docx.namelist():
            return {}

        root = ET.fromstring(docx.read("word/styles.xml"))
        levels: Dict[str, int] = {}
        for style in root.findall("w:style", NS):
            if w_attr(style, "type") != "paragraph":
                continue

            style_id = w_attr(style, "styleId")
            if not style_id:
                continue

            outline_level = self._style_outline_level(style)
            if outline_level is not None:
                levels[style_id] = outline_level
                continue

            name = style.find("w:name", NS)
            style_name = w_attr(name, "val", "") if name is not None else ""
            heading_level = self._heading_level(style_name or style_id)
            if heading_level:
                levels[style_id] = heading_level

        return levels

    def _style_outline_level(self, style: ET.Element) -> Optional[int]:
        outline = style.find("w:pPr/w:outlineLvl", NS)
        value = w_attr(outline, "val") if outline is not None else None
        if value and value.isdigit():
            return min(int(value) + 1, 6)
        return None

    def _read_comments(self, docx: zipfile.ZipFile) -> Dict[str, Comment]:
        if "word/comments.xml" not in docx.namelist():
            return {}

        root = ET.fromstring(docx.read("word/comments.xml"))
        comments: Dict[str, Comment] = {}
        comment_para_ids: Dict[str, str] = {}
        for comment in root.findall("w:comment", NS):
            comment_id = w_attr(comment, "id")
            if comment_id is None:
                continue
            paragraph = comment.find("w:p", NS)
            para_id = paragraph.attrib.get(f"{{{W14_NS}}}paraId") if paragraph is not None else None
            if para_id:
                comment_para_ids[comment_id] = para_id
            text = self._collect_text(comment)
            comments[comment_id] = Comment(
                marker=f"COMMENT-{int(comment_id) + 1 if comment_id.isdigit() else comment_id}",
                author=w_attr(comment, "author", "") or "",
                date=(w_attr(comment, "date", "") or "")[:10],
                text=text,
            )
        self._apply_comment_reply_relationships(docx, comments, comment_para_ids)
        return comments

    def _apply_comment_reply_relationships(
        self,
        docx: zipfile.ZipFile,
        comments: Dict[str, Comment],
        comment_para_ids: Dict[str, str],
    ) -> None:
        if "word/commentsExtended.xml" not in docx.namelist():
            return

        comment_id_by_para_id = {
            para_id: comment_id
            for comment_id, para_id in comment_para_ids.items()
        }
        if not comment_id_by_para_id:
            return

        root = ET.fromstring(docx.read("word/commentsExtended.xml"))
        parent_para_ids: Dict[str, str] = {}
        for comment_ex in root.findall(f"{{{W15_NS}}}commentEx"):
            para_id = comment_ex.attrib.get(f"{{{W15_NS}}}paraId")
            parent_para_id = comment_ex.attrib.get(f"{{{W15_NS}}}paraIdParent")
            if para_id and parent_para_id:
                parent_para_ids[para_id] = parent_para_id

        for comment_id, para_id in comment_para_ids.items():
            parent_para_id = parent_para_ids.get(para_id)
            parent_comment_id = comment_id_by_para_id.get(parent_para_id or "")
            if parent_comment_id and comment_id in comments and parent_comment_id in comments:
                comments[comment_id].reply_to = comments[parent_comment_id].marker

    def _paragraph_to_markdown(self, paragraph: ET.Element, preserve_hyperlinks: bool = True) -> str:
        text = self._inline_content(paragraph, preserve_hyperlinks=preserve_hyperlinks)
        numbering_prefix = "" if self._is_bold_label_bullet(paragraph) else self._numbering_prefix(paragraph)
        if numbering_prefix:
            text = f"{numbering_prefix} {text}" if text else numbering_prefix

        if not text:
            return ""

        heading_level = self._paragraph_heading_level(paragraph)
        if heading_level:
            text = f"{'#' * heading_level} {text}"

        return text

    def _is_bold_label_bullet(self, paragraph: ET.Element) -> bool:
        if not self._is_bullet_paragraph(paragraph):
            return False

        text_runs = [
            run
            for run in paragraph.findall("w:r", NS)
            if "".join(text.text or "" for text in run.findall("w:t", NS)).strip()
        ]
        if not text_runs:
            return False

        plain_text = self._collect_text(paragraph)
        if len(plain_text) > 80:
            return False

        return all(self._is_bold_run(run) for run in text_runs)

    def _is_bullet_paragraph(self, paragraph: ET.Element) -> bool:
        num_pr = paragraph.find("w:pPr/w:numPr", NS)
        if num_pr is None:
            return False

        num_id_element = num_pr.find("w:numId", NS)
        ilvl_element = num_pr.find("w:ilvl", NS)
        num_id = w_attr(num_id_element, "val") if num_id_element is not None else None
        ilvl = w_attr(ilvl_element, "val", "0") if ilvl_element is not None else "0"
        if num_id is None or ilvl is None:
            return False

        level = self.numbering.get(num_id, {}).get(ilvl)
        return level is not None and level.num_format == "bullet"

    def _paragraph_heading_level(self, paragraph: ET.Element) -> Optional[int]:
        outline = paragraph.find("w:pPr/w:outlineLvl", NS)
        outline_value = w_attr(outline, "val") if outline is not None else None
        if outline_value and outline_value.isdigit():
            return min(int(outline_value) + 1, 6)

        style = paragraph.find("w:pPr/w:pStyle", NS)
        style_value = w_attr(style, "val", "") if style is not None else ""
        if style_value:
            if style_value in self.style_heading_levels:
                return self.style_heading_levels[style_value]
            return self._heading_level(style_value)

        return None

    def _numbering_prefix(self, paragraph: ET.Element) -> str:
        num_pr = paragraph.find("w:pPr/w:numPr", NS)
        if num_pr is None:
            return ""

        num_id_element = num_pr.find("w:numId", NS)
        ilvl_element = num_pr.find("w:ilvl", NS)
        num_id = w_attr(num_id_element, "val") if num_id_element is not None else None
        ilvl = w_attr(ilvl_element, "val", "0") if ilvl_element is not None else "0"
        if num_id is None or ilvl is None:
            return ""

        level = self.numbering.get(num_id, {}).get(ilvl)
        if level is None:
            return ""

        key = f"{num_id}:{ilvl}"
        current = self.numbering_counters.get(key, level.start - 1) + 1
        self.numbering_counters[key] = current

        level_index = int(ilvl) if ilvl.isdigit() else 0
        if level.num_format == "bullet":
            return f"{'  ' * self._bullet_indent_level(level, level_index)}-"
        if level.num_format in {"decimal", ""}:
            value = str(current)
        elif level.num_format == "lowerLetter":
            value = self._alpha_number(current).lower()
        elif level.num_format == "upperLetter":
            value = self._alpha_number(current).upper()
        elif level.num_format == "lowerRoman":
            value = self._roman_number(current).lower()
        elif level.num_format == "upperRoman":
            value = self._roman_number(current).upper()
        else:
            value = str(current)

        level_text = level.level_text or f"%{level_index + 1}."
        return re.sub(r"%\d+", value, level_text)

    def _bullet_indent_level(self, level: NumberingLevel, level_index: int) -> int:
        if level.level_text in {"", "\uf0d8", "\uf075", ""}:
            return level_index + 1
        return level_index

    def _alpha_number(self, value: int) -> str:
        result = ""
        while value > 0:
            value -= 1
            result = chr(ord("A") + value % 26) + result
            value //= 26
        return result or "A"

    def _roman_number(self, value: int) -> str:
        numerals = [
            (1000, "M"),
            (900, "CM"),
            (500, "D"),
            (400, "CD"),
            (100, "C"),
            (90, "XC"),
            (50, "L"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I"),
        ]
        result = ""
        for number, numeral in numerals:
            while value >= number:
                result += numeral
                value -= number
        return result or "I"

    def _heading_level(self, style_value: str) -> Optional[int]:
        lowered = style_value.lower()
        match = re.search(r"heading(\d+)", lowered)
        if match:
            return min(int(match.group(1)), 6)
        match = re.search(r"(\d+)$", style_value)
        if lowered.startswith(("title", "heading")) and match:
            return min(int(match.group(1)), 6)
        return None

    def _table_to_markdown(self, table: ET.Element, table_index: int) -> str:
        codeblock = self._code_like_single_cell_table(table)
        if codeblock:
            return codeblock

        matrix = self._table_matrix(table)
        if not matrix:
            return ""

        matrix = self._trim_empty_edges(matrix)
        if not matrix:
            return ""

        header_rows = self._detect_header_rows(matrix)
        if header_rows > 1:
            headers = self._flatten_headers(matrix[:header_rows])
            body_rows = matrix[header_rows:]
        else:
            headers = matrix[0]
            body_rows = matrix[1:]

        if not body_rows:
            body_rows = [["" for _ in headers]]

        width = max(len(headers), *(len(row) for row in body_rows))
        headers = self._pad_row(headers, width)
        body_rows = [self._pad_row(row, width) for row in body_rows]

        lines = [f"### 表格 {table_index}", ""]
        if header_rows > 1:
            lines.extend(
                [
                    "说明：表头中的合并单元格已展开，多级表头使用 `.` 连接。",
                    "",
                ]
            )

        lines.append("| " + " | ".join(escape_markdown_cell(cell) for cell in headers) + " |")
        lines.append("| " + " | ".join("---" for _ in headers) + " |")
        for row in body_rows:
            lines.append("| " + " | ".join(escape_markdown_cell(cell) for cell in row) + " |")
        return "\n".join(lines)

    def _code_like_single_cell_table(self, table: ET.Element) -> str:
        rows = table.findall("w:tr", NS)
        if len(rows) != 1:
            return ""

        cells = rows[0].findall("w:tc", NS)
        if len(cells) != 1:
            return ""

        lines: List[str] = []
        for paragraph in cells[0].findall("w:p", NS):
            text = self._paragraph_to_markdown(paragraph, preserve_hyperlinks=False)
            if text:
                lines.append(text)

        text = "\n".join(lines)
        if not is_code_like_text(text):
            return ""

        return "```\n" + text + "\n```"

    def _table_matrix(self, table: ET.Element) -> List[List[str]]:
        rows: List[List[str]] = []
        active_vmerges: Dict[int, str] = {}
        expected_width = len(table.findall("w:tblGrid/w:gridCol", NS))
        if expected_width == 0:
            expected_width = max((len(tr.findall("w:tc", NS)) for tr in table.findall("w:tr", NS)), default=0)

        for tr in table.findall("w:tr", NS):
            row: List[str] = []
            col = 0
            cells = tr.findall("w:tc", NS)
            for cell_index, tc in enumerate(cells):
                remaining_span = sum(self._grid_span(cell) for cell in cells[cell_index:])
                while (
                    col in active_vmerges
                    and expected_width
                    and len(row) + remaining_span < expected_width
                ):
                    row.append(active_vmerges[col])
                    col += 1

                text = self._cell_text(tc)
                grid_span = self._grid_span(tc)
                vmerge = self._vmerge_value(tc)

                if vmerge == "continue":
                    text = active_vmerges.get(col, text)
                elif vmerge == "restart":
                    for offset in range(grid_span):
                        active_vmerges[col + offset] = text
                else:
                    for offset in range(grid_span):
                        active_vmerges.pop(col + offset, None)

                for _ in range(grid_span):
                    row.append(text)
                    col += 1

            while col < expected_width:
                if col in active_vmerges:
                    row.append(active_vmerges[col])
                else:
                    row.append("")
                col += 1

            expected_width = max(expected_width, len(row))
            rows.append(row)

        width = max((len(row) for row in rows), default=0)
        return [self._pad_row(row, width) for row in rows]

    def _cell_text(self, cell: ET.Element) -> str:
        parts: List[str] = []
        for child in cell:
            if child.tag == w_tag("p"):
                text = self._paragraph_to_markdown(child)
                if text:
                    parts.append(text)
            elif child.tag == w_tag("tbl"):
                nested = self._table_to_markdown(child, 0)
                if nested:
                    parts.append(nested)
            elif child.tag == w_tag("sdt"):
                content = child.find("w:sdtContent", NS)
                if content is not None:
                    nested_parts: List[str] = []
                    self._append_blocks(content, nested_parts, 0)
                    parts.extend(nested_parts)
        return clean_text("\n".join(parts))

    def _grid_span(self, cell: ET.Element) -> int:
        grid_span = cell.find("w:tcPr/w:gridSpan", NS)
        value = w_attr(grid_span, "val") if grid_span is not None else None
        if value and value.isdigit():
            return max(int(value), 1)
        return 1

    def _vmerge_value(self, cell: ET.Element) -> Optional[str]:
        vmerge = cell.find("w:tcPr/w:vMerge", NS)
        if vmerge is None:
            return None
        return w_attr(vmerge, "val", "continue") or "continue"

    def _detect_header_rows(self, matrix: List[List[str]]) -> int:
        if len(matrix) < 2:
            return 1

        # Tables with duplicated top-level header labels usually came from merged
        # Word headers and are better flattened for downstream consumption.
        first = matrix[0]
        non_empty = [cell for cell in first if cell]
        if len(non_empty) != len(set(non_empty)):
            return 2

        return 1

    def _flatten_headers(self, header_rows: List[List[str]]) -> List[str]:
        width = max((len(row) for row in header_rows), default=0)
        flattened: List[str] = []
        for col in range(width):
            parts: List[str] = []
            for row in header_rows:
                cell = row[col] if col < len(row) else ""
                if cell and (not parts or parts[-1] != cell):
                    parts.append(cell)
            flattened.append(".".join(parts) if parts else f"列{col + 1}")
        return flattened

    def _trim_empty_edges(self, matrix: List[List[str]]) -> List[List[str]]:
        while matrix and not any(cell for cell in matrix[0]):
            matrix.pop(0)
        while matrix and not any(cell for cell in matrix[-1]):
            matrix.pop()
        if not matrix:
            return []

        width = max(len(row) for row in matrix)
        matrix = [self._pad_row(row, width) for row in matrix]
        keep_cols = [
            col
            for col in range(width)
            if any(row[col] for row in matrix)
        ]
        return [[row[col] for col in keep_cols] for row in matrix]

    def _pad_row(self, row: List[str], width: int) -> List[str]:
        return row + [""] * (width - len(row))

    def _inline_content(self, element: ET.Element, preserve_hyperlinks: bool = True) -> str:
        tokens: List[RunToken] = []
        for child in element:
            tokens.extend(self._inline_tokens(child, preserve_hyperlinks=preserve_hyperlinks))
        return normalize_inline_markdown(clean_text(self._render_inline_tokens(tokens)))

    def _render_inline_tokens(self, tokens: List[RunToken]) -> str:
        parts: List[str] = []
        current_parts: List[str] = []
        current_bold = False
        current_strike = False

        def flush_current(use_html_bold: bool = False) -> None:
            nonlocal current_bold, current_strike
            if not current_parts:
                return
            text = "".join(current_parts)
            current_parts.clear()
            match = re.match(r"^(\s*)(.*?)(\s*)$", text, flags=re.DOTALL)
            if match is None:
                parts.append(text)
                current_bold = False
                current_strike = False
                return
            leading, body, trailing = match.groups()
            if not body:
                parts.append(text)
            else:
                rendered = body
                if current_strike:
                    rendered = f"[DELETED: {body}]"
                elif current_bold:
                    rendered = f"<strong>{rendered}</strong>" if use_html_bold else f"**{rendered}**"
                parts.append(f"{leading}{rendered}{trailing}")
            current_bold = False
            current_strike = False

        for index, token in enumerate(tokens):
            if not token.text:
                continue
            if current_parts and (token.bold != current_bold or token.strike != current_strike):
                flush_current(use_html_bold=bool(token.text.strip()))
            current_bold = token.bold
            current_strike = token.strike
            current_parts.append(token.text)
        flush_current()
        return "".join(parts)

    def _inline_tokens(self, element: ET.Element, preserve_hyperlinks: bool = True) -> List[RunToken]:
        if element.tag == w_tag("r"):
            return self._run_tokens(element, preserve_formatting=preserve_hyperlinks)
        text = self._inline_child(element, preserve_hyperlinks=preserve_hyperlinks)
        return [RunToken(text)] if text else []

    def _inline_child(self, element: ET.Element, preserve_hyperlinks: bool = True) -> str:
        if element.tag == w_tag("hyperlink"):
            if preserve_hyperlinks:
                return self._hyperlink_to_markdown(element)
            return self._inline_content(element, preserve_hyperlinks=False)
        if element.tag == w_tag("del"):
            text = self._collect_text(element)
            return f"[DELETED: {text}]" if text else ""
        if element.tag == w_tag("ins"):
            text = self._inline_content(element, preserve_hyperlinks=preserve_hyperlinks)
            return f"[INSERTED: {text}]" if text else ""
        if element.tag == w_tag("commentReference"):
            comment_id = w_attr(element, "id")
            comment = self.comments.get(comment_id or "")
            return comment.render("此处") if comment else ""
        if element.tag == w_tag("commentRangeStart") or element.tag == w_tag("commentRangeEnd"):
            return ""
        if element.tag == w_tag("smartTag"):
            return self._inline_content(element, preserve_hyperlinks=preserve_hyperlinks)
        if element.tag == w_tag("sdt"):
            content = element.find("w:sdtContent", NS)
            if content is None:
                return ""
            return self._inline_content(content, preserve_hyperlinks=preserve_hyperlinks)
        return ""

    def _hyperlink_to_markdown(self, hyperlink: ET.Element) -> str:
        text = self._inline_content(hyperlink)
        if not text:
            return ""

        rel_id = r_attr(hyperlink, "id")
        anchor = w_attr(hyperlink, "anchor")
        target = self.hyperlinks.get(rel_id or "")
        if target:
            return f"[{text}]({target})"
        if anchor:
            return f"[{text}](#{anchor})"
        return text

    def _run_tokens(self, run: ET.Element, preserve_formatting: bool = True) -> List[RunToken]:
        parts: List[str] = []
        for child in run:
            if child.tag in {w_tag("t"), w_tag("delText")}:
                parts.append(child.text or "")
            elif child.tag == w_tag("tab"):
                parts.append("\t")
            elif child.tag == w_tag("br"):
                parts.append("\n")
            elif child.tag == w_tag("commentReference"):
                comment_id = w_attr(child, "id")
                comment = self.comments.get(comment_id or "")
                if comment:
                    parts.append(comment.render("此处"))
            elif child.tag == w_tag("drawing"):
                image_markdown = self._drawing_to_markdown(child)
                if image_markdown:
                    parts.append(image_markdown)
        text = "".join(parts)
        if not text:
            return []
        preserve_text_format = preserve_formatting and not text.startswith("![")
        bold = preserve_text_format and self._is_bold_run(run)
        strike = preserve_text_format and self._is_strike_run(run)
        return [RunToken(text, bold, strike)]

    def _is_bold_run(self, run: ET.Element) -> bool:
        bold = run.find("w:rPr/w:b", NS)
        if bold is None:
            return False
        return w_attr(bold, "val", "true") not in {"false", "0", "off"}

    def _is_strike_run(self, run: ET.Element) -> bool:
        strike = run.find("w:rPr/w:strike", NS)
        dstrike = run.find("w:rPr/w:dstrike", NS)
        return self._is_on(strike) or self._is_on(dstrike)

    def _is_on(self, element: Optional[ET.Element]) -> bool:
        if element is None:
            return False
        return w_attr(element, "val", "true") not in {"false", "0", "off"}

    def _drawing_to_markdown(self, drawing: ET.Element) -> str:
        blip = drawing.find(f".//{{{A_NS}}}blip")
        if blip is None:
            return ""

        rel_id = r_attr(blip, "embed") or r_attr(blip, "link")
        if rel_id is None:
            return ""

        image_path = self.image_paths.get(rel_id)
        if image_path:
            return f"![图片]({image_path})"
        image_path = self._save_image(rel_id)
        if image_path:
            return f"![图片]({image_path})"
        return ""

    def _save_image(self, rel_id: str) -> str:
        if self.image_output_dir is None or rel_id not in self.image_blobs:
            return ""

        self.image_output_dir.mkdir(parents=True, exist_ok=True)
        ext = self.image_extensions.get(rel_id, ".jpg")
        image_filename = f"image_{self.image_index:03d}{ext}"
        self.image_index += 1

        image_path = self.image_output_dir / image_filename
        image_path.write_bytes(self.image_blobs[rel_id])
        markdown_path = f"{self.image_output_prefix}/{image_filename}"
        self.image_paths[rel_id] = markdown_path
        return markdown_path

    def _collect_text(self, element: ET.Element) -> str:
        parts: List[str] = []
        for text in element.findall(".//w:t", NS):
            parts.append(text.text or "")
        for text in element.findall(".//w:delText", NS):
            parts.append(text.text or "")
        return clean_text("".join(parts))
