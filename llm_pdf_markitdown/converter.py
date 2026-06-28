import io
import sys

from pathlib import Path
from typing import Any, BinaryIO, Dict, List

from markitdown import DocumentConverter, DocumentConverterResult
from markitdown._exceptions import MissingDependencyException, MISSING_DEPENDENCY_MESSAGE
from markitdown._stream_info import StreamInfo

_dependency_exc_info = None
try:
    import fitz
    import pdfminer
    import pdfminer.high_level
except ImportError:
    _dependency_exc_info = sys.exc_info()


ACCEPTED_MIME_TYPE_PREFIXES = [
    "application/pdf",
    "application/x-pdf",
]

ACCEPTED_FILE_EXTENSIONS = [".pdf"]


class LlmPdfConverter(DocumentConverter):
    """
    Converts PDFs to Markdown text and appends page-level image references.
    """

    def accepts(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> bool:
        mimetype = (stream_info.mimetype or "").lower()
        extension = (stream_info.extension or "").lower()

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
        if _dependency_exc_info is not None:
            raise MissingDependencyException(
                MISSING_DEPENDENCY_MESSAGE.format(
                    converter=type(self).__name__,
                    extension=".pdf",
                    feature="pdf",
                )
            ) from _dependency_exc_info[
                1
            ].with_traceback(  # type: ignore[union-attr]
                _dependency_exc_info[2]
            )

        data = file_stream.read()
        if not data:
            raise ValueError("Empty PDF stream")

        markdown = pdfminer.high_level.extract_text(io.BytesIO(data))
        image_refs = self._extract_images(
            data,
            image_output_dir=kwargs.get("image_output_dir"),
            image_output_prefix=kwargs.get("image_output_prefix", "./images"),
        )
        if image_refs:
            markdown = markdown.rstrip() + "\n\n" + self._images_to_markdown(image_refs)

        return DocumentConverterResult(markdown=markdown)

    def _extract_images(
        self,
        data: bytes,
        *,
        image_output_dir: str,
        image_output_prefix: str,
    ) -> Dict[int, List[str]]:
        if not image_output_dir:
            return {}

        output_dir = Path(image_output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        image_refs: Dict[int, List[str]] = {}
        image_counter = 1
        doc = fitz.open(stream=data, filetype="pdf")
        try:
            for page_index in range(len(doc)):
                page = doc[page_index]
                seen_xrefs = set()
                for image in page.get_images(full=True):
                    xref = image[0]
                    if xref in seen_xrefs:
                        continue
                    seen_xrefs.add(xref)

                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    image_ext = base_image["ext"]
                    image_filename = f"image_{image_counter:03d}.{image_ext}"
                    image_path = output_dir / image_filename
                    image_path.write_bytes(image_bytes)

                    page_number = page_index + 1
                    image_refs.setdefault(page_number, []).append(
                        f"{image_output_prefix.rstrip('/')}/{image_filename}"
                    )
                    image_counter += 1
        finally:
            doc.close()

        return image_refs

    def _images_to_markdown(self, image_refs: Dict[int, List[str]]) -> str:
        lines = ["## PDF 图片索引"]
        for page_number in sorted(image_refs):
            lines.append("")
            lines.append(f"### 第 {page_number} 页图片")
            for image_index, image_path in enumerate(image_refs[page_number], 1):
                lines.append(f"![第 {page_number} 页图片 {image_index}]({image_path})")
        return "\n".join(lines)
