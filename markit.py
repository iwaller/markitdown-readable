import sys
import shutil
from pathlib import Path
from markitdown import MarkItDown
from markitdown_readable import (
    ReadableDocxConverter,
    ReadablePdfConverter,
    ReadablePptxConverter,
)


def process_document(file_path):
    """Convert a document to Markdown and write extracted images."""
    file_path = Path(file_path)

    if not file_path.exists():
        print(f"Error: File not found: {file_path}")
        sys.exit(1)

    file_stem = file_path.stem
    file_ext = file_path.suffix.lower().lstrip('.')
    output_dir_name = f"{file_stem}_{file_ext}"

    result_dir = Path("./results") / output_dir_name
    images_dir = result_dir / "images"
    if images_dir.exists():
        shutil.rmtree(images_dir)
    images_dir.mkdir(parents=True, exist_ok=True)

    md = MarkItDown(enable_plugins=False)
    md.register_converter(ReadableDocxConverter(), priority=-1)
    md.register_converter(ReadablePdfConverter(), priority=-1)
    md.register_converter(ReadablePptxConverter(), priority=-1)

    convert_kwargs = {
        "image_output_dir": str(images_dir),
        "image_output_prefix": "./images",
    }
    result = md.convert(str(file_path), **convert_kwargs)
    markdown = result.markdown
    images = sorted(path.name for path in images_dir.iterdir() if path.is_file())

    content_file = result_dir / "content.md"
    content_file.write_text(markdown, encoding='utf-8')

    print(f"✓ Converted: {file_path}")
    print(f"✓ Markdown saved to: {content_file}")
    print(f"✓ Extracted {len(images)} images to: {images_dir}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python markit.py <file_path>")
        print("Example: python markit.py ./files/ll.pptx")
        sys.exit(1)

    file_path = sys.argv[1]
    process_document(file_path)
