# markitdown-readable

Readable Markdown converters for [MarkItDown](https://github.com/microsoft/markitdown).

This project extends MarkItDown conversion for DOCX, PDF, and PPTX files so the output keeps more document structure that is useful for downstream reading and processing, including tables, comments, images, slide content, and page-level image indexes.

## Usage

```bash
uv run python markit.py ./files/example.docx
```

Converted Markdown is written to:

```text
results/<filename>_<ext>/content.md
```

Extracted images are written to:

```text
results/<filename>_<ext>/images/
```

## Converters

- `markitdown_readable.docx`: DOCX to Markdown with Word structure handling.
- `markitdown_readable.pdf`: PDF text extraction plus image index output.
- `markitdown_readable.pptx`: PPTX slide, table, chart, text, and image extraction.
