def _safe_len(text) -> int:
    return len((text or "").strip())


def evaluate_slide_quality(slide: dict) -> dict:
    t = slide.get("type", "")

    result = {
        "type": t,
        "density_score": 2,        # 0 too sparse, 1 sparse, 2 good, 3 dense, 4 too dense
        "parallel_score": 2,       # 0 poor, 1 weak, 2 good, 3 strong
        "structure_score": 2,      # 0 poor, 1 weak, 2 good, 3 strong
        "completeness_score": 2,   # 0 poor, 1 weak, 2 good, 3 strong
        "action": "keep",
        "reasons": []
    }

    if t in {"content_2", "content_3extra", "content_4"}:
        cards = slide.get("cards", []) or []

        content_lengths = [_safe_len(c.get("content")) for c in cards]
        item_lengths = [_safe_len(c.get("item")) for c in cards]

        empty_cards = sum(1 for c in cards if not (c.get("content") or "").strip())
        avg_content = sum(content_lengths) / len(content_lengths) if content_lengths else 0
        max_content = max(content_lengths) if content_lengths else 0
        min_content = min(content_lengths) if content_lengths else 0

        if empty_cards > 0:
            result["completeness_score"] = 0
            result["action"] = "enrich"
            result["reasons"].append("Some cards are empty.")

        elif avg_content < 12:
            result["density_score"] = 0
            result["action"] = "enrich"
            result["reasons"].append("Card content is too short.")

        elif max_content > 100:
            result["density_score"] = 4
            result["action"] = "compress"
            result["reasons"].append("Card content is too long.")

        if content_lengths and (max_content - min_content) > 60:
            result["parallel_score"] = 0
            if result["action"] == "keep":
                result["action"] = "rebalance"
            result["reasons"].append("Card lengths are unbalanced.")

        if any(x > 24 for x in item_lengths):
            if result["action"] == "keep":
                result["action"] = "compress"
            result["reasons"].append("Some card titles are too long.")

    elif t == "flow":
        steps = slide.get("steps", []) or []
        step_lengths = [_safe_len(s) for s in steps]

        if not steps:
            result["structure_score"] = 0
            result["action"] = "enrich"
            result["reasons"].append("No flow steps found.")
        elif len(steps) < 3:
            result["structure_score"] = 1
            result["action"] = "enrich"
            result["reasons"].append("Too few steps for a clear flow.")
        elif max(step_lengths, default=0) > 28:
            result["density_score"] = 4
            result["action"] = "compress"
            result["reasons"].append("Flow steps are too long.")
        elif min(step_lengths, default=999) < 4:
            result["density_score"] = 0
            result["action"] = "enrich"
            result["reasons"].append("Flow steps are too short.")

    elif t == "table":
        columns = slide.get("columns", []) or []
        rows = slide.get("rows", []) or []

        cell_lengths = []
        empty_cells = 0
        for row in rows:
            for cell in row:
                txt = (cell or "").strip()
                cell_lengths.append(len(txt))
                if not txt:
                    empty_cells += 1

        if empty_cells > 0:
            result["completeness_score"] = 1
            result["action"] = "enrich"
            result["reasons"].append("Some table cells are empty.")
        elif max(cell_lengths, default=0) > 36:
            result["density_score"] = 4
            result["action"] = "compress"
            result["reasons"].append("Some table cells are too long.")

        if any(_safe_len(c) > 20 for c in columns):
            if result["action"] == "keep":
                result["action"] = "compress"
            result["reasons"].append("Some column headers are too long.")

    elif t in {"content_text", "content_image"}:
        title_len = _safe_len(slide.get("title"))
        content_len = _safe_len(slide.get("content"))

        if content_len < 20:
            result["density_score"] = 0
            result["action"] = "enrich"
            result["reasons"].append("Main content is too short.")
        elif content_len > 140:
            result["density_score"] = 4
            result["action"] = "compress"
            result["reasons"].append("Main content is too long.")

        if title_len > 30:
            if result["action"] == "keep":
                result["action"] = "compress"
            result["reasons"].append("Title may be too long for one line.")

    return result


def evaluate_spec_quality(spec: dict) -> dict:
    slides = spec.get("slides", []) or []
    evaluations = [evaluate_slide_quality(slide) for slide in slides]
    return {
        "slides": evaluations
    }