# -*- coding: utf-8 -*-

import json
from pathlib import Path


FLOW_TEMPLATE_INDEX = {
    "flow_chart_1": 14,
    "flow_chart_2": 15,
    "flow_chart_3": 16,
}


SLIDE_REGISTRY = {

    # --------------------------------------------------
    # Cover
    # --------------------------------------------------
    "cover": {
        "template_slide_index": 1,
        "required_fields": ["topic", "speaker"],
        "description": "Cover slide with title and speaker."
    },

    # --------------------------------------------------
    # Agenda
    # --------------------------------------------------
    "agenda": {
        "template_slide_index": 2,
        "required_fields": ["items"],
        "optional_fields": ["title"],
        "max_items": 5,
        "description": "Agenda slide."
    },

    # --------------------------------------------------
    # Section divider
    # --------------------------------------------------
    "section": {
        "template_slide_index": 3,
        "required_fields": ["name"],
        "description": "Section divider slide."
    },

    # --------------------------------------------------
    # 3 card slide with built-in images
    # --------------------------------------------------
    "content_3extra_image": {
        "template_slide_index": 4,
        "required_fields": ["title", "cards"],
        "max_cards": 3,
        "description": "Three card content slide with built-in template images."
    },

    # --------------------------------------------------
    # 3 card slide (text only)
    # --------------------------------------------------
    "content_3extra": {
        "template_slide_index": 5,
        "required_fields": ["title", "cards"],
        "max_cards": 3,
        "description": "Three card content slide."
    },

    # --------------------------------------------------
    # 2 card slides (three variants)
    # --------------------------------------------------
    "content_2_a": {
        "template_slide_index": 6,
        "required_fields": ["title", "cards"],
        "max_cards": 2,
        "description": "Two card slide variant A."
    },

    "content_2_b": {
        "template_slide_index": 7,
        "required_fields": ["cards"],
        "max_cards": 2,
        "description": "Two card slide variant B."
    },

    "content_2_c": {
        "template_slide_index": 11,
        "required_fields": ["cards"],
        "max_cards": 2,
        "description": "Two card slide variant C."
    },

    # --------------------------------------------------
    # Table slide
    # --------------------------------------------------
    "table": {
        "template_slide_index": 8,
        "required_fields": ["columns", "rows"],
        "optional_fields": ["title"],
        "description": "Table slide."
    },

    # --------------------------------------------------
    # One-content slide with built-in image
    # --------------------------------------------------
    "content_image": {
        "template_slide_index": 9,
        "required_fields": ["title", "content"],
        "description": "One-content slide using the template's built-in image."
    },

    # --------------------------------------------------
    # One-content text slide
    # --------------------------------------------------
    "content_text": {
        "template_slide_index": 10,
        "required_fields": ["title", "content"],
        "description": "One-content text slide."
    },

    # --------------------------------------------------
    # 4 card slides
    # --------------------------------------------------
    "content_4_a": {
        "template_slide_index": 12,
        "required_fields": ["title", "cards"],
        "max_cards": 4,
        "description": "Four card slide variant A."
    },

    "content_4_b": {
        "template_slide_index": 13,
        "required_fields": ["cards"],
        "max_cards": 4,
        "description": "Four card slide variant B."
    },

    # --------------------------------------------------
    # Flow slide
    # --------------------------------------------------
    "flow": {
        "template_slide_index": 14,
        "required_fields": ["title", "steps"],
        "description": "Flow / SmartArt slide."
    },

    # --------------------------------------------------
    # Ending slide
    # --------------------------------------------------
    "end": {
        "template_slide_index": 17,
        "required_fields": [],
        "description": "Thank you slide."
    },
}


def _detect_flow_variant_from_shapes(shapes):
    names = {str(s.get("name", "")).strip().lower() for s in shapes}
    for variant in ("flow_chart_1", "flow_chart_2", "flow_chart_3"):
        if variant in names:
            return variant
    return None


def _apply_template_map_overrides():
    """Best-effort sync of template indexes from template_map.json detected_type."""
    template_map_path = Path(__file__).resolve().parent / "template_map.json"
    if not template_map_path.exists():
        return

    try:
        data = json.loads(template_map_path.read_text(encoding="utf-8"))
    except Exception:
        return

    for slide in data:
        detected_type = slide.get("detected_type")
        slide_index = slide.get("slide_index")

        if not isinstance(slide_index, int) or slide_index <= 0:
            continue

        if detected_type in SLIDE_REGISTRY:
            SLIDE_REGISTRY[detected_type]["template_slide_index"] = slide_index

        flow_variant = _detect_flow_variant_from_shapes(slide.get("shapes", []))
        if flow_variant:
            FLOW_TEMPLATE_INDEX[flow_variant] = slide_index

    # Keep the base flow type aligned with default flow variant.
    SLIDE_REGISTRY["flow"]["template_slide_index"] = FLOW_TEMPLATE_INDEX.get(
        "flow_chart_1",
        SLIDE_REGISTRY["flow"]["template_slide_index"]
    )


_apply_template_map_overrides()