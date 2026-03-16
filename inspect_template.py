from pathlib import Path
from pptx import Presentation
import json
import re

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE = BASE_DIR / "template.pptx"
OUT = BASE_DIR / "template_map.json"

if not TEMPLATE.exists():
    raise FileNotFoundError(f"Template not found: {TEMPLATE}")

prs = Presentation(str(TEMPLATE))
data = []


def has_table_shape(shapes, name: str) -> bool:
    return any(s["name"] == name and s["has_table"] for s in shapes)


def has_prefix(shape_names: set, prefix: str) -> bool:
    return any(n.startswith(prefix) for n in shape_names)


def count_regex(shape_names: set, pattern: str) -> int:
    rx = re.compile(pattern)
    return sum(1 for n in shape_names if rx.fullmatch(n))


def detect_template_type(shape_names: set, shapes: list, slide_index: int, total_slides: int) -> str:
    has = lambda n: n in shape_names

    content_count = count_regex(shape_names, r"content_\d+")
    item_count = count_regex(shape_names, r"item_\d+")
    img_count = count_regex(shape_names, r"img_\d+")
    flow_count = count_regex(shape_names, r"flow_chart_\d+")
    agenda_count = count_regex(shape_names, r"agenda_\d+")

    # cover
    if has("Topic") and has("speaker_name"):
        return "cover"

    # agenda
    if has("outline") and agenda_count >= 1:
        return "agenda"

    # section
    if has("agenda_name"):
        return "section"

    # table
    if has_table_shape(shapes, "sheet_1"):
        return "table"

    # end (keep near the top because the last slide may not have stable placeholder names)
    if slide_index == total_slides:
        return "end"

    # flow / smartart variants
    if flow_count >= 1:
        return "flow"

    # single image + text
    if has("title") and has("content") and has("img"):
        return "content_image"

    # title + text only (same content family, but without editable image placeholder)
    if has("title") and has("content") and not has("img"):
        return "content_text"

    # 4-card variants
    if content_count >= 4:
        if item_count >= 4:
            return "content_4_b"
        return "content_4_a"

    # 3-card variants
    if has("title") and content_count >= 3 and item_count >= 3:
        if img_count >= 1:
            return "content_3extra_image"
        return "content_3extra"

    # 2-card variants
    if has("title") and content_count >= 2 and img_count >= 2:
        return "content_2_a"

    if content_count >= 2 and has("title_content_1") and has("title_content_2") and img_count >= 0:
        # Variant with two titled columns; images may or may not exist in the template.
        if img_count >= 2:
            return "content_2_b_image"
        return "content_2_b"

    if item_count >= 2 and content_count >= 2 and item_count < 3 and content_count < 3:
        return "content_2_c"

    return "unknown"


for i, slide in enumerate(prs.slides, start=1):
    shapes = []
    for shp in slide.shapes:
        text_preview = None
        if getattr(shp, "has_text_frame", False):
            text_preview = "\n".join(p.text for p in shp.text_frame.paragraphs).strip()

        shapes.append({
            "name": getattr(shp, "name", ""),
            "is_placeholder": getattr(shp, "is_placeholder", False),
            "has_text": getattr(shp, "has_text_frame", False),
            "has_table": getattr(shp, "has_table", False),
            "has_chart": getattr(shp, "has_chart", False),
            "shape_type": str(getattr(shp, "shape_type", "")),
            "text_preview": (text_preview[:120] if text_preview else ""),
        })

    shape_names = {s["name"] for s in shapes}

    detected_type = detect_template_type(
        shape_names=shape_names,
        shapes=shapes,
        slide_index=i,
        total_slides=len(prs.slides),
    )

    data.append({
        "slide_index": i,
        "layout": slide.slide_layout.name,
        "detected_type": detected_type,
        "shape_count": len(shapes),
        "shapes": shapes,
    })

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print(f"Saved: {OUT}")
