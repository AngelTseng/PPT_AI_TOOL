from pathlib import Path
from pptx import Presentation
import json

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE = BASE_DIR / "template.pptx"
OUT = BASE_DIR / "template_map.json"

if not TEMPLATE.exists():
    raise FileNotFoundError(f"Template not found: {TEMPLATE}")

prs = Presentation(str(TEMPLATE))
data = []


def _normalize_shape_names(shape_names: set[str]) -> set[str]:
    return {str(n).strip().lower() for n in shape_names if str(n).strip()}


def detect_template_type(shape_names: set, shapes: list, slide_index: int, total_slides: int) -> str:
    names = _normalize_shape_names(shape_names)

    def has(name: str) -> bool:
        return name.strip().lower() in names

    def has_img(prefix: str = "img") -> bool:
        prefix = prefix.lower()
        return any(n == prefix or n.startswith(f"{prefix}_") for n in names)

    # cover
    if has("Topic") and has("speaker_name"):
        return "cover"

    # agenda
    if has("outline") and has("agenda_1"):
        return "agenda"

    # section
    if has("agenda_name"):
        return "section"

    # flow
    if has("flow_chart_1"):
        return "flow"

    # table
    for shp in shapes:
        if str(shp.get("name", "")).strip().lower() == "sheet_1" and shp.get("has_table"):
            return "table"

    # content_image (more tolerant for img naming)
    if has("title") and has("content") and has_img("img"):
        return "content_image"

    # content_4 variants
    if has("content_1") and has("content_2") and has("content_3") and has("content_4"):
        if has("item_1") and has("item_2") and has("item_3") and has("item_4"):
            return "content_4_b"
        return "content_4_a"

    # content_3extra
    if has("item_1") and has("item_2") and has("item_3") and has("content_1") and has("content_2") and has("content_3"):
        return "content_3extra"

    # content_2 variants
    if has("item_1") and has("item_2") and has("content_1") and has("content_2") and not has("item_3") and not has("content_3"):
        return "content_2_c"

    if has("title") and has("content_1") and has("content_2") and has("img_1") and has("img_2"):
        return "content_2_a"

    if has("content_1") and has("content_2") and has("img_1") and has("img_2"):
        return "content_2_b"

    # end (fallback to last slide if no better match)
    if slide_index == total_slides:
        return "end"

    return "unknown"


for i, slide in enumerate(prs.slides, start=1):
    shapes = []
    for shp in slide.shapes:
        shapes.append({
            "name": getattr(shp, "name", ""),
            "is_placeholder": getattr(shp, "is_placeholder", False),
            "has_text": getattr(shp, "has_text_frame", False),
            "has_table": getattr(shp, "has_table", False),
            "has_chart": getattr(shp, "has_chart", False),
            "shape_type": str(getattr(shp, "shape_type", "")),
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
        "shapes": shapes,
    })

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print(f"Saved: {OUT}")
