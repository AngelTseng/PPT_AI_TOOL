from docx import Document

def extract_word_to_payload(path: str) -> dict:
    doc = Document(path)

    sections = []
    current = {"heading": "Untitled", "paragraphs": []}

    for p in doc.paragraphs:
        text = (p.text or "").strip()
        if not text:
            continue

        style_name = ""
        try:
            style_name = p.style.name.lower()
        except Exception:
            pass

        if "heading" in style_name:
            if current["paragraphs"]:
                sections.append(current)
            current = {"heading": text, "paragraphs": []}
        else:
            current["paragraphs"].append(text)

    if current["paragraphs"]:
        sections.append(current)

    raw_text = "\n".join(
        [f"{s['heading']}\n" + "\n".join(s["paragraphs"]) for s in sections]
    )

    title = sections[0]["heading"] if sections else "Word Report"

    return {
        "title": title,
        "sections": sections,
        "raw_text": raw_text,
    }