from slide_registry import SLIDE_REGISTRY

SUPPORTED_TYPES = set(SLIDE_REGISTRY.keys())

CONTENT_2_TYPES = {"content_2_a", "content_2_b", "content_2_c"}
CONTENT_4_TYPES = {"content_4_a", "content_4_b"}
CONTENT_3_TYPES = {"content_3extra", "content_3extra_image"}
CONTENT_TYPES = CONTENT_2_TYPES | CONTENT_3_TYPES | CONTENT_4_TYPES | {"content_image", "content_text", "table", "flow"}


def _text_len(value) -> int:
    if value is None:
        return 0
    return len(str(value).strip())


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

    agenda_slide = None
    for s in slides:
        if s.get("type") == "agenda":
            agenda_slide = s
            break

    if agenda_slide is None:
        return warnings

    items = agenda_slide.get("items", [])
    content_like = [s for s in slides if s.get("type") in ({"section"} | CONTENT_TYPES)]

    if len(content_like) < len(items):
        warnings.append("Agenda items may exceed available explanatory slides.")

    return warnings


def check_section_coverage(spec: dict):
    warnings = []
    slides = spec.get("slides", [])

    for i, slide in enumerate(slides):
        if slide.get("type") != "section":
            continue

        has_content_under_section = False
        for j in range(i + 1, len(slides)):
            t = slides[j].get("type")
            if t == "section":
                break
            if t in CONTENT_TYPES:
                has_content_under_section = True
                break

        if not has_content_under_section:
            warnings.append(
                f"Section slide at position {i+1} has no dedicated content slide before next section/end."
            )

        if i + 1 < len(slides) and slides[i + 1].get("type") == "section":
            warnings.append(f"Section slide at position {i+1} is followed by another section slide.")

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
                warnings.append(f"slides[{i}].items >5 , extra ignored")

        elif t in CONTENT_2_TYPES:
            cards = slide.get("cards", [])
            if not isinstance(cards, list):
                errors.append(f"slides[{i}].cards must be list")
            elif len(cards) > 2:
                warnings.append(f"slides[{i}].cards >2 , extra ignored")
            elif len(cards) < 2:
                warnings.append(f"slides[{i}].cards <2 , layout may look incomplete")

            for c_idx, card in enumerate(cards, start=1):
                if _text_len(card.get("content", "")) < 16:
                    warnings.append(f"slides[{i}].cards[{c_idx}].content is very short")

        elif t in CONTENT_3_TYPES:
            cards = slide.get("cards", [])
            if not isinstance(cards, list):
                errors.append(f"slides[{i}].cards must be list")
            elif len(cards) != 3:
                errors.append(f"slides[{i}].cards must be exactly 3 for {t}")

            for c_idx, card in enumerate(cards, start=1):
                if _text_len(card.get("content", "")) < 16:
                    warnings.append(f"slides[{i}].cards[{c_idx}].content is very short")

        elif t in CONTENT_4_TYPES:
            cards = slide.get("cards", [])
            if not isinstance(cards, list):
                errors.append(f"slides[{i}].cards must be list")
            elif len(cards) != 4:
                errors.append(f"slides[{i}].cards must be exactly 4 for {t}")

            for c_idx, card in enumerate(cards, start=1):
                if _text_len(card.get("content", "")) < 16:
                    warnings.append(f"slides[{i}].cards[{c_idx}].content is very short")

        elif t in {"content_image", "content_text"}:
            if not slide.get("title"):
                errors.append(f"slides[{i}].title is required")
            if not slide.get("content"):
                errors.append(f"slides[{i}].content is required")
            elif _text_len(slide.get("content")) < 20:
                warnings.append(f"slides[{i}].content is very short")

        elif t == "flow":
            steps = slide.get("steps", [])
            if not isinstance(steps, list) or len(steps) < 2:
                errors.append(f"slides[{i}].steps must have >=2")
            else:
                for s_idx, step in enumerate(steps, start=1):
                    if _text_len(step) < 8:
                        warnings.append(f"slides[{i}].steps[{s_idx}] is very short")

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

    return {"errors": errors, "warnings": warnings, "normalized_spec": spec}