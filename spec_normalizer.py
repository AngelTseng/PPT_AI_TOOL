def _enforce_layout_streak_limit(slides: list, max_streak: int = 2) -> list:
    """
    Final guardrail after normalization.
    Do not allow the same final slide type variant to repeat more than max_streak times in a row.
    """
    result = []

    for slide in slides:
        result.append(slide)

        if len(result) < max_streak + 1:
            continue

        a = result[-3].get("type", "")
        b = result[-2].get("type", "")
        c = result[-1].get("type", "")

        if a == b == c:
            t = c

            if t == "content_2_a":
                result[-1]["type"] = "content_2_b"
            elif t == "content_2_b":
                result[-1]["type"] = "content_2_c"
            elif t == "content_2_c":
                result[-1]["type"] = "content_2_a"

            elif t == "content_4_a":
                result[-1]["type"] = "content_4_b"
            elif t == "content_4_b":
                result[-1]["type"] = "content_4_a"

            elif t == "content_3extra_a":
                result[-1]["type"] = "content_3extra_b"
            elif t == "content_3extra_b":
                print("[WARN] content_3extra_b repeats more than 2 times after normalization.")
            elif t == "content_3extra_image":
                print("[WARN] content_3extra_image repeats more than 2 times after normalization.")
                
            else:
                print(f"[WARN] slide type '{t}' repeats more than 2 times after normalization.")

    return result

def _build_cards_from_items(items, limit):
    cards = []
    for i, txt in enumerate((items or [])[:limit], start=1):
        cards.append({
            "item": f"Point {i}",
            "content": txt
        })
    return cards


def _avg_content_len(cards):
    vals = [len(str(c.get("content", "")).strip()) for c in (cards or [])]
    return sum(vals) / len(vals) if vals else 0


def _max_step_len(steps):
    return max((len(str(x).strip()) for x in (steps or [])), default=0)


def normalize_beautified_spec(spec: dict) -> dict:
    slides = spec.get("slides", [])
    normalized = []

    for slide in slides:
        t = slide.get("type")

        if t == "cover":
            normalized.append({
                "type": "cover",
                "topic": slide.get("topic") or slide.get("title", ""),
                "speaker": slide.get("speaker", "")
            })

        elif t == "agenda":
            normalized.append({
                "type": "agenda",
                "title": slide.get("title", "Agenda"),
                "items": list((slide.get("items") or [])[:5]),
            })

        elif t == "section":
            section_name = str(
                slide.get("name")
                or slide.get("title")
                or "Section"
            ).strip()

            normalized.append({
                "type": "section",
                "name": section_name
            })

        elif t in ("content_image", "content_text", "single"):
            variant = str(slide.get("variant", "")).strip().lower()
            resolved = "content_image" if variant == "image" else "content_text"

            title = str(slide.get("title", "") or "").strip()
            content = str(slide.get("content", "") or "").strip()

            if not title:
                title = "Title"

            if not content:
                # 先用 title 補一個最小可通過驗證的內容
                # 避免 validator 報 slides[i].content required
                content = f"{title}：請補充內容"

            normalized.append({
                "type": resolved,
                "title": title,
                "content": content
            })

        elif t in ("content_2", "content_2_a", "content_2_b", "content_2_c"):
            cards = slide.get("cards")
            if cards is None:
                cards = _build_cards_from_items(slide.get("items", []), 2)
            cards = cards[:2]

            if t in ("content_2_a", "content_2_b", "content_2_c"):
                resolved = t
            else:
                title = str(slide.get("title", "")).strip()
                avg_len = _avg_content_len(cards)
                if title:
                    resolved = "content_2_a"
                elif avg_len > 60:
                    resolved = "content_2_c"
                else:
                    resolved = "content_2_b"

            normalized.append({
                "type": resolved,
                "title": slide.get("title", ""),
                "cards": cards
            })

        elif t in (
            "content_3",
            "content_3extra",
            "content_3extra_a",
            "content_3extra_b",
            "content_3extra_image",
        ):
            cards = slide.get("cards")
            if cards is None:
                cards = _build_cards_from_items(slide.get("items", []), 3)
            cards = cards[:3]

            title = str(slide.get("title", "")).strip()
            variant = str(slide.get("variant", "") or slide.get("template_key", "")).strip().lower()

            if t == "content_3extra_image" or variant in {"image", "content_3extra_image"}:
                resolved = "content_3extra_image"
            elif t in {"content_3extra_a", "content_3extra_b"}:
                resolved = t
            elif variant == "content_3extra_b":
                resolved = "content_3extra_b"
            elif variant == "content_3extra_a":
                resolved = "content_3extra_a"
            else:
                avg_len = _avg_content_len(cards)
                resolved = "content_3extra_b" if avg_len > 60 else "content_3extra_a"

            normalized.append({
                "type": resolved,
                "title": title,
                "cards": cards
            })

        elif t in ("content_4", "content_4_a", "content_4_b"):
            cards = slide.get("cards")
            if cards is None:
                cards = _build_cards_from_items(slide.get("items", []), 4)
            cards = cards[:4]

            if t in ("content_4_a", "content_4_b"):
                resolved = t
            else:
                resolved = "content_4_a" if str(slide.get("title", "")).strip() else "content_4_b"

            normalized.append({
                "type": resolved,
                "title": slide.get("title", ""),
                "cards": cards
            })

        elif t == "table":
            normalized.append({
                "type": "table",
                "title": slide.get("title", ""),
                "columns": slide.get("columns", []),
                "rows": slide.get("rows", [])
            })

        elif t == "flow":
            steps = list(slide.get("steps", [])[:6])
            variant = str(slide.get("variant", "")).strip().lower()

            if variant not in ("flow_chart_1", "flow_chart_2", "flow_chart_3"):
                variant = "flow_chart_3" if _max_step_len(steps) >= 22 else ""

            payload = {
                "type": "flow",
                "title": slide.get("title", ""),
                "steps": steps
            }
            if variant:
                payload["variant"] = variant

            normalized.append(payload)

        elif t == "end":
            normalized.append({"type": "end"})

        else:
            normalized.append(slide)

    normalized = _enforce_layout_streak_limit(normalized, max_streak=2)
    return {"slides": normalized}