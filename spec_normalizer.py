def normalize_beautified_spec(spec: dict) -> dict:
    slides = spec.get("slides", [])
    normalized = []

    content_2_types = {"content_2", "content_2_a", "content_2_b", "content_2_c"}
    content_4_types = {"content_4", "content_4_a", "content_4_b"}

    for slide in slides:
        t = slide.get("type")

        if t == "cover":
            normalized.append({
                "type": "cover",
                "topic": slide.get("topic") or slide.get("title", ""),
                "speaker": slide.get("speaker", "")
            })

        elif t == "section":
            normalized.append({
                "type": "section",
                "name": slide.get("name") or slide.get("title", "")
            })

        elif t in content_2_types:
            cards = slide.get("cards")
            if cards is None:
                items = slide.get("items", [])
                cards = [{"item": f"Point {i}", "content": txt} for i, txt in enumerate(items[:2], start=1)]

            normalized.append({
                "type": t,
                "title": slide.get("title", ""),
                "cards": cards[:2]
            })

        elif t == "content_3extra":
            cards = slide.get("cards")
            if cards is None:
                items = slide.get("items", [])
                cards = [{"item": f"Point {i}", "content": txt} for i, txt in enumerate(items[:3], start=1)]

            normalized.append({
                "type": "content_3extra",
                "title": slide.get("title", ""),
                "cards": cards[:3]
            })

        elif t in content_4_types:
            cards = slide.get("cards")
            if cards is None:
                items = slide.get("items", [])
                cards = [{"item": f"Point {i}", "content": txt} for i, txt in enumerate(items[:4], start=1)]

            normalized.append({
                "type": t,
                "title": slide.get("title", ""),
                "cards": cards[:4]
            })

        elif t == "content_image":
            normalized.append({
                "type": "content_image",
                "title": slide.get("title", ""),
                "content": slide.get("content", ""),
                "image": slide.get("image", "")
            })

        elif t == "agenda":
            normalized.append({
                "type": "agenda",
                "title": slide.get("title", "Agenda"),
                "items": slide.get("items", [])
            })

        elif t == "table":
            normalized.append({
                "type": "table",
                "title": slide.get("title", ""),
                "columns": slide.get("columns", []),
                "rows": slide.get("rows", [])
            })

        elif t == "flow":
            normalized.append({
                "type": "flow",
                "title": slide.get("title", ""),
                "steps": slide.get("steps", [])
            })

        elif t == "end":
            normalized.append({"type": "end"})

        else:
            normalized.append(slide)

    return {"slides": normalized}