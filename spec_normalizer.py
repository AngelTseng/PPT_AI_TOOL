def _clean_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _expand_short_text(text: str, min_chars: int = 28) -> str:
    """Pad overly short content with a concise explanatory suffix for readability."""
    cleaned = _clean_text(text)
    if not cleaned:
        return ""
    if len(cleaned) >= min_chars:
        return cleaned
    return f"{cleaned}：補充關鍵背景、做法與預期效益。"


def _normalize_cards(cards, max_cards: int):
    normalized = []

    for idx, card in enumerate(cards or [], start=1):
        if not isinstance(card, dict):
            continue

        item = _clean_text(card.get("item")) or f"Point {idx}"
        content = _expand_short_text(card.get("content", "")) or "補充關鍵背景、做法與預期效益。"

        normalized.append({"item": item, "content": content})
        if len(normalized) >= max_cards:
            break

    # Keep card-count aligned with target layout capacity.
    while len(normalized) < max_cards:
        idx = len(normalized) + 1
        normalized.append({
            "item": f"Point {idx}",
            "content": "補充關鍵背景、做法與預期效益。"
        })

    return normalized




def _ensure_cover_and_end(slides: list[dict]) -> list[dict]:
    """Guarantee deck starts with cover and ends with end slide."""
    if not slides:
        return [
            {"type": "cover", "topic": "Presentation", "speaker": ""},
            {"type": "end"},
        ]

    cover_slide = None
    end_slide = None
    body_slides = []

    for slide in slides:
        t = slide.get("type")
        if t == "cover" and cover_slide is None:
            cover_slide = slide
            continue
        if t == "end":
            end_slide = slide
            continue
        body_slides.append(slide)

    if cover_slide is None:
        cover_slide = {"type": "cover", "topic": "Presentation", "speaker": ""}

    if end_slide is None:
        end_slide = {"type": "end"}

    return [cover_slide] + body_slides + [end_slide]

def normalize_beautified_spec(spec: dict) -> dict:
    slides = spec.get("slides", [])
    normalized = []

    content_2_types = {"content_2", "content_2_a", "content_2_b", "content_2_c"}
    content_4_types = {"content_4", "content_4_a", "content_4_b"}

    for slide in slides:
        t = slide.get("type")

        if t == "content_2":
            t = "content_2_a"
        elif t == "content_4":
            t = "content_4_a"

        if t == "cover":
            normalized.append({
                "type": "cover",
                "topic": _clean_text(slide.get("topic") or slide.get("title", "")),
                "speaker": _clean_text(slide.get("speaker", ""))
            })

        elif t == "section":
            section_name = _clean_text(slide.get("name") or slide.get("title", ""))
            if section_name:
                normalized.append({
                    "type": "section",
                    "name": section_name
                })

        elif t in content_2_types:
            cards = slide.get("cards")
            if cards is None:
                items = slide.get("items", [])
                cards = [{"item": f"Point {i}", "content": txt} for i, txt in enumerate(items[:2], start=1)]

            normalized.append({
                "type": t,
                "title": _clean_text(slide.get("title", "")) or "重點整理",
                "cards": _normalize_cards(cards, 2)
            })

        elif t == "content_3extra":
            cards = slide.get("cards")
            if cards is None:
                items = slide.get("items", [])
                cards = [{"item": f"Point {i}", "content": txt} for i, txt in enumerate(items[:3], start=1)]

            normalized.append({
                "type": "content_3extra",
                "title": _clean_text(slide.get("title", "")) or "重點整理",
                "cards": _normalize_cards(cards, 3)
            })

        elif t in content_4_types:
            cards = slide.get("cards")
            if cards is None:
                items = slide.get("items", [])
                cards = [{"item": f"Point {i}", "content": txt} for i, txt in enumerate(items[:4], start=1)]

            normalized.append({
                "type": t,
                "title": _clean_text(slide.get("title", "")) or "四大重點",
                "cards": _normalize_cards(cards, 4)
            })

        elif t == "content_image":
            normalized.append({
                "type": "content_image",
                "title": _clean_text(slide.get("title", "")) or "重點說明",
                "content": _expand_short_text(slide.get("content", "")),
                "image": _clean_text(slide.get("image", ""))
            })

        elif t == "agenda":
            normalized.append({
                "type": "agenda",
                "title": _clean_text(slide.get("title", "Agenda")) or "Agenda",
                "items": [_clean_text(x) for x in slide.get("items", []) if _clean_text(x)]
            })

        elif t == "table":
            normalized.append({
                "type": "table",
                "title": _clean_text(slide.get("title", "")),
                "columns": slide.get("columns", []),
                "rows": slide.get("rows", [])
            })

        elif t == "flow":
            normalized.append({
                "type": "flow",
                "title": _clean_text(slide.get("title", "")),
                "steps": [_expand_short_text(x, min_chars=12) for x in slide.get("steps", [])]
            })

        elif t == "end":
            normalized.append({"type": "end"})

        else:
            normalized.append(slide)

    normalized = _ensure_cover_and_end(normalized)

    return {"slides": normalized}
