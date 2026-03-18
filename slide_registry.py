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
        "family": "cover",
        "variant": "default",
        "template_key": "cover",
        "template_slide_index": 1,
        "required_fields": ["topic", "speaker"],
        "density": "light",
        "image_slots": 0,
        "fit_profile": {
            "title_max_chars": 36,
            "body_max_chars": 0,
        },
        "best_for": ["cover", "opening"],
        "description": "Cover slide with title and speaker."
    },

    # --------------------------------------------------
    # Agenda
    # --------------------------------------------------
    "agenda": {
        "family": "agenda",
        "variant": "default",
        "template_key": "agenda",
        "template_slide_index": 2,
        "required_fields": ["items"],
        "optional_fields": ["title"],
        "max_items": 5,
        "density": "medium",
        "image_slots": 0,
        "fit_profile": {
            "title_max_chars": 28,
            "item_max_chars": 24,
            "max_items": 5,
        },
        "best_for": ["overview", "agenda"],
        "description": "Agenda slide."
    },

    # --------------------------------------------------
    # Section divider
    # --------------------------------------------------
    "section": {
        "family": "section",
        "variant": "default",
        "template_key": "section",
        "template_slide_index": 3,
        "required_fields": ["name"],
        "density": "light",
        "image_slots": 0,
        "fit_profile": {
            "title_max_chars": 30,
        },
        "best_for": ["divider", "chapter break", "transition"],
        "description": "Section divider slide."
    },

    # --------------------------------------------------
    # 3 card slide with built-in images
    # --------------------------------------------------
    "content_3extra_image": {
        "family": "content_3",
        "variant": "image",
        "template_key": "content_3extra_image",
        "template_slide_index": 4,
        "required_fields": ["title", "cards"],
        "max_cards": 3,
        "density": "medium",
        "image_slots": 3,
        "fit_profile": {
            "title_max_chars": 30,
            "item_max_chars": 18,
            "content_max_chars": 48,
            "max_cards": 3,
        },
        "best_for": ["three highlights", "three visual points"],
        "description": "Three card content slide with built-in template images."
    },

    # --------------------------------------------------
    # 3 card slide (text only)
    # --------------------------------------------------
    "content_3extra": {
        "family": "content_3",
        "variant": "text",
        "template_key": "content_3extra",
        "template_slide_index": 5,
        "required_fields": ["title", "cards"],
        "max_cards": 3,
        "density": "dense",
        "image_slots": 0,
        "fit_profile": {
            "title_max_chars": 30,
            "item_max_chars": 18,
            "content_max_chars": 56,
            "max_cards": 3,
        },
        "best_for": ["three grouped ideas", "three parallel points"],
        "description": "Three card content slide."
    },

    # --------------------------------------------------
    # 2 card slides (three variants)
    # --------------------------------------------------
    "content_2_a": {
        "family": "content_2",
        "variant": "a",
        "template_key": "content_2_a",
        "template_slide_index": 6,
        "required_fields": ["title", "cards"],
        "max_cards": 2,
        "density": "medium",
        "image_slots": 2,
        "fit_profile": {
            "title_max_chars": 30,
            "item_max_chars": 18,
            "content_max_chars": 64,
            "max_cards": 2,
        },
        "best_for": ["two concepts", "balanced comparison", "paired highlights"],
        "description": "Two card slide variant A."
    },

    "content_2_b": {
        "family": "content_2",
        "variant": "b",
        "template_key": "content_2_b",
        "template_slide_index": 7,
        "required_fields": ["cards"],
        "max_cards": 2,
        "density": "light",
        "image_slots": 2,
        "fit_profile": {
            "item_max_chars": 18,
            "content_max_chars": 54,
            "max_cards": 2,
        },
        "best_for": ["short two-up summary", "paired visual points"],
        "description": "Two card slide variant B."
    },

    "content_2_c": {
        "family": "content_2",
        "variant": "c",
        "template_key": "content_2_c",
        "template_slide_index": 11,
        "required_fields": ["cards"],
        "max_cards": 2,
        "density": "dense",
        "image_slots": 0,
        "fit_profile": {
            "item_max_chars": 20,
            "content_max_chars": 80,
            "max_cards": 2,
        },
        "best_for": ["two text-heavy explanations", "before after"],
        "description": "Two card slide variant C."
    },

    # --------------------------------------------------
    # Table slide
    # --------------------------------------------------
    "table": {
        "family": "table",
        "variant": "default",
        "template_key": "table",
        "template_slide_index": 8,
        "required_fields": ["columns", "rows"],
        "optional_fields": ["title"],
        "density": "dense",
        "image_slots": 0,
        "fit_profile": {
            "title_max_chars": 28,
            "max_columns": 5,
            "max_rows": 8,
            "cell_max_chars": 28,
        },
        "best_for": ["comparison", "structured facts", "matrix"],
        "description": "Table slide."
    },

    # --------------------------------------------------
    # One-content slide with built-in image
    # --------------------------------------------------
    "content_image": {
        "family": "single",
        "variant": "image",
        "template_key": "content_image",
        "template_slide_index": 9,
        "required_fields": ["title", "content"],
        "density": "light",
        "image_slots": 1,
        "fit_profile": {
            "title_max_chars": 30,
            "content_max_chars": 220,
        },
        "best_for": ["single key message", "hero explanation"],
        "description": "One-content slide using the template's built-in image."
    },

    # --------------------------------------------------
    # One-content text slide
    # --------------------------------------------------
    "content_text": {
        "family": "single",
        "variant": "text",
        "template_key": "content_text",
        "template_slide_index": 10,
        "required_fields": ["title", "content"],
        "density": "medium",
        "image_slots": 0,
        "fit_profile": {
            "title_max_chars": 30,
            "content_max_chars": 260,
        },
        "best_for": ["single explanation", "one message text"],
        "description": "One-content text slide."
    },

    # --------------------------------------------------
    # 4 card slides
    # --------------------------------------------------
    "content_4_a": {
        "family": "content_4",
        "variant": "a",
        "template_key": "content_4_a",
        "template_slide_index": 12,
        "required_fields": ["title", "cards"],
        "max_cards": 4,
        "density": "dense",
        "image_slots": 0,
        "fit_profile": {
            "title_max_chars": 28,
            "item_max_chars": 16,
            "content_max_chars": 40,
            "max_cards": 4,
        },
        "best_for": ["four capabilities", "four item summary"],
        "description": "Four card slide variant A."
    },

    "content_4_b": {
        "family": "content_4",
        "variant": "b",
        "template_key": "content_4_b",
        "template_slide_index": 13,
        "required_fields": ["cards"],
        "max_cards": 4,
        "density": "medium",
        "image_slots": 0,
        "fit_profile": {
            "item_max_chars": 16,
            "content_max_chars": 34,
            "max_cards": 4,
        },
        "best_for": ["short four-up highlights", "compact summary"],
        "description": "Four card slide variant B."
    },

    # --------------------------------------------------
    # Flow slide
    # --------------------------------------------------
    "flow": {
        "family": "flow",
        "variant": "default",
        "template_key": "flow_chart_1",
        "template_slide_index": 14,
        "required_fields": ["title", "steps"],
        "density": "medium",
        "image_slots": 0,
        "fit_profile": {
            "title_max_chars": 28,
            "step_max_chars": 26,
            "max_steps": 6,
        },
        "best_for": ["process", "sequence", "workflow"],
        "description": "Flow / SmartArt slide."
    },

    # --------------------------------------------------
    # Ending slide
    # --------------------------------------------------
    "end": {
        "family": "end",
        "variant": "default",
        "template_key": "end",
        "template_slide_index": 17,
        "required_fields": [],
        "density": "light",
        "image_slots": 0,
        "fit_profile": {},
        "best_for": ["closing"],
        "description": "Thank you slide."
    },
}


FAMILY_VARIANTS = {
    "content_2": ["content_2_a", "content_2_b", "content_2_c"],
    "content_3": ["content_3extra", "content_3extra_image"],
    "content_4": ["content_4_a", "content_4_b"],
    "single": ["content_image", "content_text"],
    "flow": ["flow"],
    "table": ["table"],
    "cover": ["cover"],
    "agenda": ["agenda"],
    "section": ["section"],
    "end": ["end"],
}


FLOW_VARIANTS = ("flow_chart_1", "flow_chart_2", "flow_chart_3")


def get_layout_config(slide_type: str) -> dict:
    return SLIDE_REGISTRY[slide_type]


def get_layout_family(slide_type: str) -> str:
    cfg = get_layout_config(slide_type)
    return cfg.get("family", slide_type)


def get_family_variants(family: str) -> list[str]:
    return list(FAMILY_VARIANTS.get(family, []))


def resolve_flow_template_key(slide_spec: dict | None = None) -> str:
    slide_spec = slide_spec or {}
    preferred = str(slide_spec.get("variant") or slide_spec.get("template_key") or "").strip().lower()
    if preferred in FLOW_TEMPLATE_INDEX:
        return preferred
    return "flow_chart_1"


def _detect_flow_variant_from_shapes(shapes):
    names = {str(s.get("name", "")).strip().lower() for s in shapes}
    for variant in FLOW_VARIANTS:
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
