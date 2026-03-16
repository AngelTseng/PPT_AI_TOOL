# -*- coding: utf-8 -*-

import json
from pathlib import Path


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
    # 3 card slide
    # --------------------------------------------------
    "content_3extra": {
        "template_slide_index": 4,
        "required_fields": ["title", "cards"],
        "max_cards": 3,
        "description": "Three card content slide."
    },

    # --------------------------------------------------
    # 2 card slides (three variants)
    # --------------------------------------------------
    "content_2_a": {
        "template_slide_index": 5,
        "required_fields": ["title", "cards"],
        "max_cards": 2,
        "description": "Two card slide variant A."
    },

    "content_2_b": {
        "template_slide_index": 6,
        "required_fields": ["title", "cards"],
        "max_cards": 2,
        "description": "Two card slide variant B."
    },

    "content_2_c": {
        "template_slide_index": 9,
        "required_fields": ["title", "cards"],
        "max_cards": 2,
        "description": "Two card slide variant C."
    },

    # --------------------------------------------------
    # Table slide
    # --------------------------------------------------
    "table": {
        "template_slide_index": 7,
        "required_fields": ["columns", "rows"],
        "optional_fields": ["title"],
        "description": "Table slide."
    },

    # --------------------------------------------------
    # Image + text slide (previously unknown)
    # --------------------------------------------------
    "content_image": {
        "template_slide_index": 8,
        "required_fields": ["title", "content"],
        "optional_fields": ["image"],
        "description": "Single image with text slide."
    },

    # --------------------------------------------------
    # 4 card slides
    # --------------------------------------------------
    "content_4_a": {
        "template_slide_index": 10,
        "required_fields": ["title", "cards"],
        "max_cards": 4,
        "description": "Four card slide variant A."
    },

    "content_4_b": {
        "template_slide_index": 11,
        "required_fields": ["title", "cards"],
        "max_cards": 4,
        "description": "Four card slide variant B."
    },

    # --------------------------------------------------
    # Flow slide
    # --------------------------------------------------
    "flow": {
        "template_slide_index": 12,
        "required_fields": ["title", "steps"],
        "description": "Flow / SmartArt slide."
    },

    # --------------------------------------------------
    # Ending slide
    # --------------------------------------------------
    "end": {
        "template_slide_index": 13,
        "required_fields": [],
        "description": "Thank you slide."
    },
}


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

        if detected_type in SLIDE_REGISTRY and isinstance(slide_index, int) and slide_index > 0:
            SLIDE_REGISTRY[detected_type]["template_slide_index"] = slide_index


_apply_template_map_overrides()
