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


def _normalize_items_as_cards(items, max_cards: int):
    cards = []
    for i, txt in enumerate(items or [], start=1):
        cleaned = _clean_text(txt)
        if not cleaned:
            continue
        cards.append({"item": f"Point {i}", "content": cleaned})
        if len(cards) >= max_cards:
            break
    return cards


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
        t = _normalize_type(slide.get("type"))
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




def _normalize_type(value) -> str:
    t = _clean_text(value).lower()
    alias_map = {
        "content_1": "content_1",
        "content": "content",
        "content-1": "content_1",
        "content3": "content_3",
        "content_3": "content_3",
        "content-3": "content_3",
        "content_3extra": "content_3extra",
        "content_3_extra": "content_3extra",
        "content_3extra_image": "content_3extra_image",
        "content_3_extra_image": "content_3extra_image",
        "content_3image": "content_3extra_image",
    }
    return alias_map.get(t, t)


def _pick_content_variant(base_type: str, counter: int) -> str:
    if base_type == "content_2":
        variants = ["content_2_a", "content_2_b", "content_2_c"]
        return variants[counter % len(variants)]
    if base_type == "content_4":
        variants = ["content_4_a", "content_4_b"]
        return variants[counter % len(variants)]
    return base_type


def _pick_content_3_variant(counter: int) -> str:
    variants = ["content_3extra", "content_3extra_image"]
    return variants[counter % len(variants)]

def _has_real_image_hint(slide: dict) -> bool:
    image = _clean_text(slide.get("image", ""))
    image_url = _clean_text(slide.get("image_url", ""))
    image_path = _clean_text(slide.get("image_path", ""))
    return bool(image or image_url or image_path)


def _rebalance_single_content_variants(slides: list[dict]) -> list[dict]:
    """
    For one-content family:
    - pure text single page -> content_text (slide 10)
    - multiple pages -> alternate content_text / content_image
    """
    idxs = []
    for i, s in enumerate(slides):
        if s.get("type") in {"content", "content_1", "content_text", "content_image"}:
            idxs.append(i)

    if not idxs:
        return slides

    if len(idxs) == 1:
        i = idxs[0]
        slide = slides[i]
        if _has_real_image_hint(slide):
            slides[i]["type"] = "content_image"
        else:
            slides[i]["type"] = "content_text"
        return slides

    # 多頁時交替使用 slide 10 / slide 9
    for order, i in enumerate(idxs):
        slide = slides[i]
        if order % 2 == 0:
            slide["type"] = "content_text"   # slide 10 first
        else:
            slide["type"] = "content_image"  # slide 9 second

    return slides

def _rebalance_content_3_variants(slides: list[dict]) -> list[dict]:
    """
    For 3-card family:
    - if there are multiple slides, alternate content_3extra_image / content_3extra
    so slide 4 and slide 5 are both used.
    """
    idxs = []
    for i, s in enumerate(slides):
        if s.get("type") in {"content_3", "content_3extra", "content_3extra_image"}:
            idxs.append(i)

    if not idxs:
        return slides

    if len(idxs) == 1:
        i = idxs[0]
        if slides[i].get("type") == "content_3":
            slides[i]["type"] = "content_3extra"
        return slides

    for order, i in enumerate(idxs):
        if order % 2 == 0:
            slides[i]["type"] = "content_3extra_image"  # slide 4
        else:
            slides[i]["type"] = "content_3extra"        # slide 5

    return slides

def normalize_beautified_spec(spec: dict) -> dict:
    slides = spec.get("slides", [])
    normalized = []

    content_2_types = {"content_2", "content_2_a", "content_2_b", "content_2_c"}
    content_3_types = {"content_3extra", "content_3extra_image"}
    content_4_types = {"content_4", "content_4_a", "content_4_b"}

    content_2_counter = 0
    content_3_counter = 0
    content_4_counter = 0

    for slide in slides:
        t = _normalize_type(slide.get("type"))

        # Expand generic aliases into concrete template-supported variants
        if t == "content_2":
            t = _pick_content_variant("content_2", content_2_counter)
            content_2_counter += 1
        elif t == "content_3":
            t = _pick_content_3_variant(content_3_counter)
            content_3_counter += 1
        elif t == "content_4":
            t = _pick_content_variant("content_4", content_4_counter)
            content_4_counter += 1

        if t == "cover":
            normalized.append({
                "type": "cover",
                "topic": _clean_text(slide.get("topic") or slide.get("title", "")) or "Presentation",
                "speaker": _clean_text(slide.get("speaker", ""))
            })

        elif t == "section":
            section_name = _clean_text(slide.get("name") or slide.get("title", ""))
            if section_name:
                normalized.append({
                    "type": "section",
                    "name": section_name
                })

        elif t in {"content", "content_1"}:
            content = _expand_short_text(slide.get("content", ""))
            if not content:
                items = [_clean_text(x) for x in slide.get("items", []) if _clean_text(x)]
                content = "\n".join(items)

            normalized.append({
                "type": "content_text",  # 先給預設，後面 rebalance 還會再調整
                "title": _clean_text(slide.get("title", "")) or "重點說明",
                "content": content or "N/A",
                "image": _clean_text(slide.get("image", "")),
                "image_url": _clean_text(slide.get("image_url", "")),
                "image_path": _clean_text(slide.get("image_path", "")),
            })

        elif t in content_2_types:
            cards = slide.get("cards")
            if cards is None:
                cards = _normalize_items_as_cards(slide.get("items", []), 2)

            normalized.append({
                "type": t,
                "title": _clean_text(slide.get("title", "")) or "重點整理",
                "cards": _normalize_cards(cards, 2)
            })

        elif t in content_3_types:
            cards = slide.get("cards")
            if cards is None:
                cards = _normalize_items_as_cards(slide.get("items", []), 3)

            normalized.append({
                "type": t,
                "title": _clean_text(slide.get("title", "")) or "重點整理",
                "cards": _normalize_cards(cards, 3)
            })

        elif t in content_4_types:
            cards = slide.get("cards")
            if cards is None:
                cards = _normalize_items_as_cards(slide.get("items", []), 4)

            normalized.append({
                "type": t,
                "title": _clean_text(slide.get("title", "")) or "四大重點",
                "cards": _normalize_cards(cards, 4)
            })

        elif t == "content_image":
            normalized.append({
                "type": "content_image",
                "title": _clean_text(slide.get("title", "")) or "重點說明",
                "content": _normalize_content_image_content(slide),
                "image": _clean_text(slide.get("image", ""))
            })

        elif t == "content_text":
            content = _expand_short_text(slide.get("content", ""))
            if not content:
                items = [_clean_text(x) for x in slide.get("items", []) if _clean_text(x)]
                content = "\n".join(items)

            normalized.append({
                "type": "content_text",
                "title": _clean_text(slide.get("title", "")) or "重點說明",
                "content": content or "N/A"
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
                "steps": [_expand_short_text(x, min_chars=12) for x in slide.get("steps", []) if _clean_text(x)]
            })

        elif t == "end":
            normalized.append({"type": "end"})

        else:
            fallback = dict(slide)
            if t:
                fallback["type"] = t
            normalized.append(fallback)

    normalized = _rebalance_single_content_variants(normalized)
    normalized = _rebalance_content_3_variants(normalized)
    normalized = _ensure_cover_and_end(normalized)
    return {"slides": normalized}
