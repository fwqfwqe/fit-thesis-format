#!/usr/bin/env python3
"""Optional PDF-to-PNG rendering. Requires pypdfium2 and Pillow, not an AI host."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path


def digest(path: Path) -> str:
    with path.open("rb") as handle:
        return hashlib.file_digest(handle, "sha256").hexdigest() if hasattr(hashlib, "file_digest") else _digest_stream(handle)


def _digest_stream(handle) -> str:
    sha = hashlib.sha256()
    for block in iter(lambda: handle.read(1024 * 1024), b""):
        sha.update(block)
    return sha.hexdigest()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_pdf", type=Path)
    parser.add_argument("output_directory", type=Path, help="Must not already exist; renders every page.")
    parser.add_argument("--dpi", type=int, default=144, help="72–300, default 144.")
    args = parser.parse_args()
    source = args.input_pdf.expanduser().resolve()
    destination = args.output_directory.expanduser().resolve()
    if not source.is_file() or source.suffix.lower() != ".pdf":
        parser.error("input_pdf must be an existing .pdf file")
    if destination.exists():
        parser.error("refusing to reuse or overwrite an existing output directory")
    if not 72 <= args.dpi <= 300:
        parser.error("--dpi must be between 72 and 300")
    try:
        import pypdfium2 as pdfium
        import PIL.Image  # noqa: F401: bitmap.to_pil needs this optional dependency.
    except ImportError as exc:
        parser.error(f"Optional PDF rendering dependencies are unavailable: {exc}. Use an existing PDF viewer or approved conversion tool instead.")
    before = digest(source)
    pdf = pdfium.PdfDocument(str(source))
    pages = []
    try:
        if len(pdf) == 0:
            raise ValueError("The PDF contains no pages.")
        destination.mkdir(parents=True, exist_ok=False)
        for index in range(len(pdf)):
            page = pdf[index]
            bitmap = None
            picture = None
            try:
                bitmap = page.render(scale=args.dpi / 72)
                picture = bitmap.to_pil()
                filename = f"page-{index + 1:03d}.png"
                picture.save(destination / filename)
                pages.append({"page": index + 1, "file": filename, "width": picture.width, "height": picture.height})
            finally:
                if picture is not None:
                    picture.close()
                if bitmap is not None:
                    bitmap.close()
                page.close()
    finally:
        pdf.close()
    after = digest(source)
    if before != after:
        raise RuntimeError("Source PDF changed during rendering; do not deliver this run.")
    result = {"input": str(source), "input_sha256": before, "source_unchanged": True, "dpi": args.dpi, "page_count": len(pages), "pages": pages}
    (destination / "manifest.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"output_directory": str(destination), "page_count": len(pages), "source_unchanged": True}, ensure_ascii=False))


if __name__ == "__main__":
    main()
