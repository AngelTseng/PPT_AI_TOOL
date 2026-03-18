def decide_target_type(slide: dict) -> str:
    t = slide.get("type", "")

    # 已知型別先保留
    if t in {
        "cover", "agenda", "section", "end",
        "table", "flow",
        "content_2", "content_3extra", "content_4",
        "content_image", "content_text"
    }:
        return t

    # unknown 頁先看結構訊號
    if slide.get("has_table"):
        return "table"

    if slide.get("has_smartart"):
        return "flow"

    text_boxes = slide.get("text_boxes", []) or []
    if len(text_boxes) >= 4:
        return "content_4"
    if len(text_boxes) == 3:
        return "content_3extra"
    if len(text_boxes) == 2:
        return "content_2"

    return "content_text"

    
def _text_boxes_to_cards(text_boxes: list[dict], limit: int) -> list[dict]:
    cards = []
    for i, box in enumerate(text_boxes[:limit], start=1):
        txt = (box.get("text") or "").strip()
        if not txt:
            continue
        cards.append({
            "item": f"Point {i}",
            "content": txt
        })
    return cards

def transform_slide_by_rules(slide: dict) -> dict:
    target_type = decide_target_type(slide)
    original_type = slide.get("type", "")
    title = slide.get("title", "")
    text_boxes = slide.get("text_boxes", []) or []

    # 已知 content 類型先保留原 cards
    if original_type == "content_2" and slide.get("cards") is not None:
        out = dict(slide)
        out["type"] = "content_2"
        return out

    if original_type == "content_3extra" and slide.get("cards") is not None:
        out = dict(slide)
        out["type"] = "content_3extra"
        return out

    if original_type == "content_4" and slide.get("cards") is not None:
        out = dict(slide)
        out["type"] = "content_4"
        return out

    # 再處理由規則轉型的頁
    if target_type == "content_2":
        return {
            "type": "content_2",
            "title": title,
            "cards": _text_boxes_to_cards(text_boxes, 2)
        }

    if target_type == "content_3extra":
        return {
            "type": "content_3extra",
            "title": title,
            "cards": _text_boxes_to_cards(text_boxes, 3)
        }

    if target_type == "content_4":
        return {
            "type": "content_4",
            "title": title,
            "cards": _text_boxes_to_cards(text_boxes, 4)
        }

    if target_type in {
        "cover", "agenda", "section", "end",
        "table", "flow", "content_image", "content_text"
    }:
        out = dict(slide)
        out["type"] = target_type
        return out

    return {
        "type": "content_text",
        "title": title,
        "content": slide.get("content", "")
    }
    
def rule_based_transform_spec(spec: dict) -> dict:
    slides = spec.get("slides", []) or []
    out_slides = []

    for slide in slides:
        out_slides.append(transform_slide_by_rules(slide))

    return {
        **spec,
        "slides": out_slides
    }