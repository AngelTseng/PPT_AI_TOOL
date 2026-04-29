from __future__ import annotations

from pathlib import Path

from pypdf import PdfReader


def extract_pdf_to_payload(pdf_path: str) -> dict:
    path = Path(pdf_path)
    if not path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    reader = PdfReader(str(path))
    pages: list[dict] = []
    text_chunks: list[str] = []

    for idx, page in enumerate(reader.pages, start=1):
        page_text = (page.extract_text() or "").strip()
        pages.append({
            "page_number": idx,
            "text": page_text,
        })
        if page_text:
            text_chunks.append(f"[Page {idx}]\n{page_text}")

    return {
        "title": path.stem,
        "num_pages": len(reader.pages),
        "pages": pages,
        "raw_text": "\n\n".join(text_chunks),
    }
