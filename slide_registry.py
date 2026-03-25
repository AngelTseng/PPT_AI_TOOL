# -*- coding: utf-8 -*-
import json
from pathlib import Path


FLOW_TEMPLATE_INDEX = {
    "flow_chart_1": 14,
    "flow_chart_2": 15,
    "flow_chart_3": 16,
}

SLIDE_REGISTRY = {
    "cover": {
        "family": "cover",
        "variant": "default",
        "template_key": "cover",
        "template_slide_index": 1,
        "required_fields": ["topic", "speaker"],
        "density": "light",
        "image_slots": 0,
        "fit_profile": {"title_max_chars": 36, "body_max_chars": 0},
    },
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
        "fit_profile": {"title_max_chars": 28, "item_max_chars": 24, "max_items": 5},
    },
    "section": {
        "family": "section",
        "variant": "default",
        "template_key": "section",
        "template_slide_index": 3,
        "required_fields": ["name"],
        "density": "light",
        "image_slots": 0,
        "fit_profile": {"title_max_chars": 30},
    },

    "content_3extra_image": {
        "family": "content_3",
        "variant": "image",
        "template_key": "content_3extra_image",
        "template_slide_index": 4,
        "required_fields": ["title", "cards"],
        "max_cards": 3,
        "density": "medium",
        "image_slots": 1,
        "fit_profile": {
            "title_max_chars": 30,
            "item_max_chars": 18,
            "content_max_chars": 48,
            "max_cards": 3,
        },
    },
    "content_3extra_a": {
        "family": "content_3",
        "variant": "a",
        "template_key": "content_3extra_a",
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
    },
    "content_3extra_b": {
        "family": "content_3",
        "variant": "b",
        "template_key": "content_3extra_b",
        "template_slide_index": 6,
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
    },

    "content_2_a": {
        "family": "content_2",
        "variant": "a",
        "template_key": "content_2_a",
        "template_slide_index": 7,
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
    },
    "content_2_b": {
        "family": "content_2",
        "variant": "b",
        "template_key": "content_2_b",
        "template_slide_index": 8,
        "required_fields": ["cards"],
        "max_cards": 2,
        "density": "light",
        "image_slots": 0,
        "fit_profile": {
            "item_max_chars": 18,
            "content_max_chars": 54,
            "max_cards": 2,
        },
    },
    "content_2_c": {
        "family": "content_2",
        "variant": "c",
        "template_key": "content_2_c",
        "template_slide_index": 12,
        "required_fields": ["cards"],
        "max_cards": 2,
        "density": "dense",
        "image_slots": 0,
        "fit_profile": {
            "item_max_chars": 20,
            "content_max_chars": 80,
            "max_cards": 2,
        },
    },

    "table": {
        "family": "table",
        "variant": "default",
        "template_key": "table",
        "template_slide_index": 9,
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
    },

    "content_image": {
        "family": "single",
        "variant": "image",
        "template_key": "content_image",
        "template_slide_index": 10,
        "required_fields": ["title", "content"],
        "density": "light",
        "image_slots": 1,
        "fit_profile": {
            "title_max_chars": 30,
            "content_max_chars": 220,
        },
    },
    "content_text": {
        "family": "single",
        "variant": "text",
        "template_key": "content_text",
        "template_slide_index": 11,
        "required_fields": ["title", "content"],
        "density": "medium",
        "image_slots": 0,
        "fit_profile": {
            "title_max_chars": 30,
            "content_max_chars": 260,
        },
    },

    "content_4_a": {
        "family": "content_4",
        "variant": "a",
        "template_key": "content_4_a",
        "template_slide_index": 13,
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
    },
    "content_4_b": {
        "family": "content_4",
        "variant": "b",
        "template_key": "content_4_b",
        "template_slide_index": 14,
        "required_fields": ["cards"],
        "max_cards": 4,
        "density": "medium",
        "image_slots": 0,
        "fit_profile": {
            "item_max_chars": 16,
            "content_max_chars": 34,
            "max_cards": 4,
        },
    },

    "flow": {
        "family": "flow",
        "variant": "default",
        "template_key": "flow_chart_1",
        "template_slide_index": 15,
        "required_fields": ["title", "steps"],
        "density": "medium",
        "image_slots": 0,
        "fit_profile": {
            "title_max_chars": 28,
            "step_max_chars": 26,
            "max_steps": 6,
            "long_step_threshold": 22,
        },
    },

    "end": {
        "family": "end",
        "variant": "default",
        "template_key": "end",
        "template_slide_index": 18,
        "required_fields": [],
        "density": "light",
        "image_slots": 0,
        "fit_profile": {},
    },
}


FAMILY_VARIANTS = {
    "content_2": ["content_2_a", "content_2_b", "content_2_c"],
    "content_3": ["content_3extra_a", "content_3extra_b", "content_3extra_image"],
    "content_4": ["content_4_a", "content_4_b"],
    "single": ["content_image", "content_text"],
    "flow": ["flow"],
    "table": ["table"],
    "cover": ["cover"],
    "agenda": ["agenda"],
    "section": ["section"],
    "end": ["end"],
}


ABSTRACT_TYPE_ALIASES = {
    "content_2": "content_2_a",
    "content_3": "content_3extra_a",
    "content_4": "content_4_a",
    "single": "content_text",
}


FLOW_VARIANTS = ("flow_chart_1", "flow_chart_2", "flow_chart_3")


def get_layout_config(slide_type: str) -> dict:
    return SLIDE_REGISTRY[slide_type]


def get_layout_family(slide_type: str) -> str:
    cfg = get_layout_config(slide_type)
    return cfg.get("family", slide_type)


def get_family_variants(family: str) -> list[str]:
    return list(FAMILY_VARIANTS.get(family, []))


def normalize_registry_type(slide_type: str) -> str:
    return ABSTRACT_TYPE_ALIASES.get(slide_type, slide_type)


def resolve_flow_template_key(slide_spec: dict | None = None) -> str:
    slide_spec = slide_spec or {}
    preferred = str(slide_spec.get("variant") or slide_spec.get("template_key") or "").strip().lower()
    if preferred in FLOW_TEMPLATE_INDEX:
        return preferred

    steps = [str(x).strip() for x in slide_spec.get("steps", []) if str(x).strip()]
    max_len = max((len(x) for x in steps), default=0)

    # 第三種給長文字
    if max_len >= 22:
        return "flow_chart_3"

    text = " ".join(steps).lower()
    loop_keywords = ["iteration", "iterate", "feedback", "cycle", "optimize", "improve", "review", "迭代", "循環", "回饋", "優化", "改善", "檢討"]
    if any(k in text for k in loop_keywords):
        return "flow_chart_2"

    return "flow_chart_1"


def _detect_flow_variant_from_shapes(shapes):
    names = {str(s.get("name", "")).strip().lower() for s in shapes}
    for variant in FLOW_VARIANTS:
        if variant in names:
            return variant
    return None


def _infer_detected_type_from_shapes(shapes):
    names = {str(s.get("name", "")).strip().lower() for s in shapes}

    if {"title_content_1", "title_content_2", "content_1", "content_2"}.issubset(names) and "title" not in names:
        return "content_2_b"

    if {"item_1", "item_2", "content_1", "content_2"}.issubset(names) and "title" not in names:
        return "content_2_c"

    if {"title_content_1", "title_content_2", "content_1", "content_2", "title"}.issubset(names):
        return "content_2_a"

    if {"title", "content", "img"}.issubset(names):
        return "content_image"

    if {"title", "content"}.issubset(names) and "img" not in names:
        return "content_text"

    if {"title", "sheet_1"}.issubset(names):
        return "table"

    if {"title", "item_1", "item_2", "item_3", "content_1", "content_2", "content_3"}.issubset(names):
        if {"img", "img_1", "img_2", "img_3"}.intersection(names):
            return "content_3extra_image"
        return "content_3extra_a"

    if {"content_1", "content_2", "content_3", "content_4"}.issubset(names):
        if "title" in names:
            return "content_4_a"
        return "content_4_b"

    if "agenda_name" in names:
        return "section"
    if "outline" in names:
        return "agenda"
    if {"topic", "speaker_name"}.issubset(names):
        return "cover"
    if any(n.startswith("flow_chart_") for n in names):
        return "flow"
    return None


def resolve_content_3_template_key(slide_spec: dict | None = None) -> str:
    slide_spec = slide_spec or {}
    preferred = str(
        slide_spec.get("variant") or slide_spec.get("template_key") or ""
    ).strip().lower()

    allowed = {
        "content_3extra_a",
        "content_3extra_b",
        "content_3extra_image",
    }

    if preferred in allowed:
        return preferred

    # 預設
    return "content_3extra_a"


def _apply_template_map_overrides():
    template_map_path = Path(__file__).resolve().parent / "template_map.json"
    if not template_map_path.exists():
        return

    try:
        data = json.loads(template_map_path.read_text(encoding="utf-8"))
    except Exception:
        return

    for slide in data:
        slide_index = slide.get("slide_index")
        shapes = slide.get("shapes", [])
        detected_type = slide.get("detected_type")
        inferred_type = _infer_detected_type_from_shapes(shapes)

        if not isinstance(slide_index, int) or slide_index <= 0:
            continue

        resolved_type = detected_type if detected_type in SLIDE_REGISTRY else inferred_type

        if resolved_type in SLIDE_REGISTRY:
            SLIDE_REGISTRY[resolved_type]["template_slide_index"] = slide_index

        flow_variant = _detect_flow_variant_from_shapes(shapes)
        if flow_variant:
            FLOW_TEMPLATE_INDEX[flow_variant] = slide_index

    SLIDE_REGISTRY["flow"]["template_slide_index"] = FLOW_TEMPLATE_INDEX.get(
        "flow_chart_1",
        SLIDE_REGISTRY["flow"]["template_slide_index"]
    )


_apply_template_map_overrides()