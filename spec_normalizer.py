def _clean_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _expand_short_text(text: str, min_chars: int = 28) -> str:
    """Only clean text; do not auto-append suggestion sentences into PPT content."""
    return _clean_text(text)


def _normalize_content_image_content(slide: dict) -> str:
    """Ensure content_image keeps non-empty content without adding coaching text."""
    content = _expand_short_text(slide.get("content", ""))
    if content:
        return content

    title_hint = _clean_text(slide.get("title", ""))
    if title_hint:
        return title_hint

    return "N/A"


def _normalize_cards(cards, max_cards: int):
    normalized = []

    for idx, card in enumerate(cards or [], start=1):
        if not isinstance(card, dict):
            continue

        item = _clean_text(card.get("item")) or f"Point {idx}"
        content = _expand_short_text(card.get("content", ""))

        normalized.append({"item": item, "content": content})
        if len(normalized) >= max_cards:
            break

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



def _pick_content_variant(base_type: str, counter: int) -> str:
    if base_type == "content_2":
        variants = ["content_2_a", "content_2_b", "content_2_c"]
        return variants[counter % len(variants)]
    if base_type == "content_3extra":
        variants = ["content_3extra", "content_3extra_image"]
        return variants[counter % len(variants)]
    if base_type == "content_4":
        variants = ["content_4_a", "content_4_b"]
        return variants[counter % len(variants)]
    return base_type

def normalize_beautified_spec(spec: dict) -> dict:
    slides = spec.get("slides", [])
    normalized = []

    content_2_types = {"content_2", "content_2_a", "content_2_b", "content_2_c"}
    content_4_types = {"content_4", "content_4_a", "content_4_b"}
    content_3_types = {"content_3extra", "content_3extra_image"}
    content_2_counter = 0
    content_3_counter = 0
    content_4_counter = 0

    for slide in slides:
        t = slide.get("type")

        if t == "content_2":
            t = _pick_content_variant("content_2", content_2_counter)
            content_2_counter += 1
        elif t == "content_4":
            t = _pick_content_variant("content_4", content_4_counter)
            content_4_counter += 1

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

        elif t in content_3_types:
            cards = slide.get("cards")
            if cards is None:
                items = slide.get("items", [])
                cards = [{"item": f"Point {i}", "content": txt} for i, txt in enumerate(items[:3], start=1)]

            picked_3 = _pick_content_variant("content_3extra", content_3_counter)
            content_3_counter += 1
            normalized.append({
                "type": picked_3,
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

        elif t in {"content_image", "content_text"}:
            normalized.append({
                "type": t,
                "title": _clean_text(slide.get("title", "")) or "重點說明",
                "content": _normalize_content_image_content(slide),
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
