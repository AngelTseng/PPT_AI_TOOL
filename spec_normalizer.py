def normalize_beautified_spec(spec: dict) -> dict:
    slides = spec.get("slides", [])
    normalized = []

    for slide in slides:
        t = slide.get("type")

        # COVER
        if t == "cover":
            normalized.append({
                "type": "cover",
                "topic": slide.get("topic") or slide.get("title", ""),
                "speaker": slide.get("speaker", "")
            })

        # SECTION
        elif t == "section":
            normalized.append({
                "type": "section",
                "name": slide.get("name") or slide.get("title", "")
            })

        elif t == "content_2":
    
            cards = slide.get("cards")

            if cards is None:
                items = slide.get("items", [])
                cards = []

                for i, txt in enumerate(items[:2], start=1):
                    cards.append({
                        "item": f"Point {i}",
                        "content": txt
                    })

            normalized.append({
                "type": "content_2",
                "title": slide.get("title", ""),
                "cards": cards[:2]
            })
        
        # CONTENT_3EXTRA
        elif t == "content_3extra":

            cards = slide.get("cards")

            # LLM sometimes outputs "items"
            if cards is None:
                items = slide.get("items", [])

                cards = []
                for i, txt in enumerate(items[:3], start=1):
                    cards.append({
                        "item": f"Point {i}",
                        "content": txt
                    })

            normalized.append({
                "type": "content_3extra",
                "title": slide.get("title", ""),
                "cards": cards[:3]
            })
            
        elif t == "content_4":
    
            cards = slide.get("cards")

            if cards is None:
                items = slide.get("items", [])
                cards = []

                for i, txt in enumerate(items[:4], start=1):
                    cards.append({
                        "item": f"Point {i}",
                    "content": txt
                    })

            normalized.append({
                    "type": "content_4",
                    "title": slide.get("title", ""),
                    "cards": cards[:4]
                })

        # AGENDA
        elif t == "agenda":
            normalized.append({
                "type": "agenda",
                "title": slide.get("title", "Agenda"),
                "items": slide.get("items", [])
            })

        # TABLE
        elif t == "table":
            normalized.append({
                "type": "table",
                "title": slide.get("title", ""),
                "columns": slide.get("columns", []),
                "rows": slide.get("rows", [])
            })

        # FLOW
        elif t == "flow":
            normalized.append({
                "type": "flow",
                "title": slide.get("title", ""),
                "steps": slide.get("steps", [])
            })

        # END
        elif t == "end":
            normalized.append({"type": "end"})

        else:
            normalized.append(slide)

    return {"slides": normalized}