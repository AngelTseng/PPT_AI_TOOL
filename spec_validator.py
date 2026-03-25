SUPPORTED_TYPES = {
    "cover",
    "agenda",
    "section",
    "content_image",
    "content_text",
    "content_2_a",
    "content_2_b",
    "content_2_c",
    "content_3extra_a",
    "content_3extra_b",
    "content_3extra_image",
    "content_4_a",
    "content_4_b",
    "table",
    "flow",
    "end",
}

CONTENT_TYPES = {
    "content_image",
    "content_text",
    "content_2_a",
    "content_2_b",
    "content_2_c",
    "content_3extra_a",
    "content_3extra_b",
    "content_3extra_image",
    "content_4_a",
    "content_4_b",
    "table",
    "flow",
}


def check_slide_diversity(spec: dict):
    warnings = []
    slides = spec.get("slides", [])

    types = [s.get("type") for s in slides if s.get("type") not in ("cover", "end")]
    unique_types = set(types)

    if len(slides) >= 5 and len(unique_types) < 3:
        warnings.append("Deck uses fewer than 3 distinct slide types; may look monotonous.")

    streak = 1
    for i in range(1, len(types)):
        if types[i] == types[i - 1]:
            streak += 1
            if streak >= 3:
                warnings.append(f"Slide type '{types[i]}' repeats 3 or more times in a row.")
                break
        else:
            streak = 1

    return warnings


def check_agenda_coverage(spec: dict):
    warnings = []
    slides = spec.get("slides", [])

    agenda_slide = next((s for s in slides if s.get("type") == "agenda"), None)
    if agenda_slide is None:
        return warnings

    items = agenda_slide.get("items", [])
    content_like = [s for s in slides if s.get("type") in CONTENT_TYPES or s.get("type") == "section"]

    if len(content_like) < len(items):
        warnings.append("Agenda items may exceed available explanatory slides.")

    return warnings


def check_section_coverage(spec: dict):
    warnings = []
    slides = spec.get("slides", [])

    for i, slide in enumerate(slides[:-1]):
        if slide.get("type") == "section":
            next_type = slides[i + 1].get("type")
            if next_type not in CONTENT_TYPES:
                warnings.append(
                    f"Section slide at position {i+1} is not immediately followed by a content slide."
                )

    return warnings


def validate_deck_spec(spec: dict):
    errors = []
    warnings = []

    slides = spec.get("slides")
    if not isinstance(slides, list) or not slides:
        errors.append("slides must be a non-empty list")
        return {"errors": errors, "warnings": warnings, "normalized_spec": spec}

    for i, slide in enumerate(slides, start=1):
        t = slide.get("type")

        if t not in SUPPORTED_TYPES:
            errors.append(f"slides[{i}] unsupported type: {t}")
            continue

        if t == "agenda":
            items = slide.get("items", [])
            if not isinstance(items, list):
                errors.append(f"slides[{i}].items must be list")
            elif len(items) > 5:
                warnings.append(f"slides[{i}].items > 5, extra ignored")

        elif t in {"content_image", "content_text"}:
            if not str(slide.get("title", "")).strip():
                errors.append(f"slides[{i}].title required")
            if not str(slide.get("content", "")).strip():
                errors.append(f"slides[{i}].content required")

        elif t in {"content_2_a", "content_2_b", "content_2_c"}:
            cards = slide.get("cards", [])
            if not isinstance(cards, list):
                errors.append(f"slides[{i}].cards must be list")
            elif len(cards) > 2:
                warnings.append(f"slides[{i}].cards > 2, extra ignored")

        elif t in {"content_3extra_a", "content_3extra_b", "content_3extra_image"}:
            cards = slide.get("cards", [])
            if not isinstance(cards, list):
                errors.append(f"slides[{i}].cards must be list")
            elif len(cards) > 3:
                warnings.append(f"slides[{i}].cards > 3, extra ignored")

        elif t in {"content_4_a", "content_4_b"}:
            cards = slide.get("cards", [])
            if not isinstance(cards, list):
                errors.append(f"slides[{i}].cards must be list")
            elif len(cards) > 4:
                warnings.append(f"slides[{i}].cards > 4, extra ignored")

        elif t == "flow":
            steps = slide.get("steps", [])
            if not isinstance(steps, list) or len(steps) < 2:
                errors.append(f"slides[{i}].steps must have >= 2")
            elif len(steps) > 6:
                warnings.append(f"slides[{i}].steps > 6, extra ignored")

            variant = slide.get("variant", "")
            if variant and variant not in {"flow_chart_1", "flow_chart_2", "flow_chart_3"}:
                errors.append(f"slides[{i}].variant invalid: {variant}")

        elif t == "table":
            columns = slide.get("columns", [])
            rows = slide.get("rows", [])
            if not isinstance(columns, list) or not columns:
                errors.append(f"slides[{i}].columns invalid")
            if not isinstance(rows, list) or not rows:
                errors.append(f"slides[{i}].rows invalid")

    warnings.extend(check_slide_diversity(spec))
    warnings.extend(check_agenda_coverage(spec))
    warnings.extend(check_section_coverage(spec))

    return {
        "errors": errors,
        "warnings": warnings,
        "normalized_spec": spec
    }
    
def check_slide_count_by_budget(spec: dict, min_slides: int, max_slides: int):
    warnings = []
    slides = spec.get("slides", [])
    count = len(slides)

    if count < min_slides:
        warnings.append(f"Slide count {count} is below recommended minimum {min_slides}.")
    if count > max_slides:
        warnings.append(f"Slide count {count} exceeds recommended maximum {max_slides}.")

    return warnings