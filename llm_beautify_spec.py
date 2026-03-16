import json
from pathlib import Path
from openai import OpenAI

from config import OPENAI_MODEL
from slide_registry import SLIDE_REGISTRY

client = OpenAI()

BASE_DIR = Path(__file__).resolve().parent
SUPPORTED_SLIDE_TYPES = list(SLIDE_REGISTRY.keys())
INPUT_SPEC = BASE_DIR / "extracted_deck_spec.json"
OUTPUT_SPEC = BASE_DIR / "beautified_deck_spec.json"


BEAUTIFY_SCHEMA = {
    "name": "beautified_deck_spec",
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "slides": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "type": {
                            "type": "string",
                            "enum": SUPPORTED_SLIDE_TYPES
                        },
                        "topic": {"type": "string"},
                        "speaker": {"type": "string"},
                        "title": {"type": "string"},
                        "content": {"type": "string"},
                        "items": {
                            "type": "array",
                            "items": {"type": "string"},
                            "maxItems": 5
                        },
                        "name": {"type": "string"},
                        "cards": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "additionalProperties": False,
                                "properties": {
                                    "item": {"type": "string"},
                                    "content": {"type": "string"}
                                },
                                "required": ["item", "content"]
                            }
                        },
                        "columns": {
                            "type": "array",
                            "items": {"type": "string"},
                            "minItems": 1
                        },
                        "rows": {
                            "type": "array",
                            "minItems": 1,
                            "items": {
                                "type": "array",
                                "items": {"type": "string"}
                            }
                        },
                        "steps": {
                            "type": "array",
                            "items": {"type": "string"},
                            "minItems": 2
                        }
                    },
                    "required": ["type"],
                    "allOf": [
                        {
                            "if": {"properties": {"type": {"const": "cover"}}},
                            "then": {"required": ["type", "topic", "speaker"]}
                        },
                        {
                            "if": {"properties": {"type": {"const": "agenda"}}},
                            "then": {"required": ["type", "items"]}
                        },
                        {
                            "if": {"properties": {"type": {"const": "section"}}},
                            "then": {"required": ["type", "name"]}
                        },
                        {
                            "if": {"properties": {"type": {"enum": ["content_2", "content_2_a"]}}},
                            "then": {"required": ["type", "title", "cards"]}
                        },
                        {
                            "if": {"properties": {"type": {"enum": ["content_2_b", "content_2_c"]}}},
                            "then": {"required": ["type", "cards"]}
                        },
                        {
                            "if": {"properties": {"type": {"enum": ["content_4", "content_4_a"]}}},
                            "then": {"required": ["type", "title", "cards"]}
                        },
                        {
                            "if": {"properties": {"type": {"enum": ["content_4_b"]}}},
                            "then": {"required": ["type", "cards"]}
},
                        {
                            "if": {"properties": {"type": {"const": "content_image"}}},
                            "then": {"required": ["type", "title", "content"]}
                        },
                        {
                            "if": {"properties": {"type": {"const": "content_3extra"}}},
                            "then": {"required": ["type", "title", "cards"]}
                        },
                        {
                            "if": {"properties": {"type": {"const": "table"}}},
                            "then": {"required": ["type", "columns", "rows"]}
                        },
                        {
                            "if": {"properties": {"type": {"const": "flow"}}},
                            "then": {"required": ["type", "title", "steps"]}
                        },
                        {
                            "if": {"properties": {"type": {"const": "end"}}},
                            "then": {"required": ["type"]}
                        }
                    ]
                }
            }
        },
        "required": ["slides"]
    }
}


def load_extracted_spec(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def build_prompt(extracted_spec: dict) -> str:
    return f"""
You are improving an existing company-template PowerPoint deck.

You will receive an extracted deck spec from an existing presentation.
Many slides may have type="unknown". You must convert them into valid supported slide types.

Supported slide types:
- cover
- agenda
- section
- content_2 / content_2_a / content_2_b / content_2_c
- content_3extra
- content_image
- content_4 / content_4_a / content_4_b
- table
- flow
- end

Primary goals:
- preserve the original meaning
- improve readability
- shorten verbose text
- use concise PowerPoint wording
- make the deck less monotonous
- do not output unknown
- do not invent unsupported slide types

Diversity rules:
- Use varied layouts across the deck.
- Do not overuse content_3extra.
- Do not repeat content_3extra more than 2 times in a row.
- For a normal deck, use at least 3 distinct slide types excluding cover and end.
- For decks with 7 or more slides, use at least 4 distinct slide types excluding cover and end.
- Prefer a presentation rhythm with visible variation, instead of many slides with the same structure.

Agenda rules:
- Agenda is optional.
- If agenda is included, keep it concise.
- If agenda is included, each agenda item should correspond to at least one explanatory slide.

Hard layout rules:
- content_3extra must always contain exactly 3 cards.
- content_4/content_4_a/content_4_b must always contain exactly 4 cards.
- Do not use the same content_3 variant for all 3-card slides.
- Do not use the same content_4 variant for all 4-card slides.
- When the deck contains 2 or more 3-card slides, use at least 2 different 3-card layout variants.
- When the deck contains 2 or more 4-card slides, use at least 2 different 4-card layout variants.

Section coverage rules:
- Every section slide must be immediately followed by at least one explanatory content slide.
- Do not place two section slides back-to-back.
- Do not create a section slide unless there is at least one following content slide dedicated to that section.
- If a topic only needs one content slide, prefer using that content slide directly instead of adding a separate section slide.

Slide usage guidance:
- cover: title / opening page
- agenda: optional overview
- section: major topic break, chapter divider, or rhythm change
- content_2 variants (content_2/content_2_a/content_2_b/content_2_c): two grouped ideas or two-column explanation
- content_3extra: exactly three grouped ideas or three parallel highlights
- content_image: one key image with one explanatory text block
- If a slide has only one core idea/content block, prefer content_image as the one-content slide (do not force it into content_2/content_3extra/content_4).
- content_4 variants (content_4/content_4_a/content_4_b): four grouped ideas, four capabilities, or four-item summary
- table: comparisons, structured facts, grouped responsibilities, categories, tools, or capability summaries
- flow: process, sequence, collaboration stages, lifecycle, or learning path
- end: closing slide

Strong layout guidance:
- If content naturally fits 2 grouped ideas, prefer content_2.
- If content naturally fits 4 grouped ideas, prefer content_4 or table.
- Only use content_3extra when the content truly fits a 3-point grouped layout.
- If a slide reads like a chapter heading or transition, prefer section.
- If a slide has a single key message with one supporting paragraph, prefer content_image over multi-card layouts.
- If a slide contains categories, grouped facts, responsibilities, tools, or capability summaries, prefer table.
- If a slide contains stages, phases, sequence, workflow, or lifecycle, prefer flow.
- Do not default every explanatory slide to content_3extra.

Good deck rhythm examples:
- cover → section → content_2 → table → flow → end
- cover → agenda → section → content_3extra → table → section → flow → end
- cover → section → content_2 → content_4 → table → flow → end

Important:
- If a slide has raw title + multiple explanatory paragraphs, convert it into the most suitable content slide instead of defaulting to content_3extra.
- If content is too long for one slide, split it into multiple slides when needed.
- Use section slides to break the deck into meaningful chapters and improve pacing.

Here is the extracted presentation spec:
{json.dumps(extracted_spec, ensure_ascii=False, indent=2)}

Output valid deck_spec JSON only.
"""

def sanitize_slides(spec: dict) -> dict:
    allowed = set(SUPPORTED_SLIDE_TYPES)

    slides = spec.get("slides", [])
    cleaned = []

    for s in slides:
        if isinstance(s, dict) and s.get("type") in allowed:
            cleaned.append(s)
        else:
            print("[WARN] Dropping malformed slide:", s)

    spec["slides"] = cleaned
    return spec

def diversify_layouts(spec: dict) -> dict:
    
    slides = spec.get("slides", [])

    for i in range(2, len(slides)):
        t0 = slides[i - 2].get("type")
        t1 = slides[i - 1].get("type")
        t2 = slides[i].get("type")

        if t0 == t1 == t2 == "content_3extra":

            title = slides[i].get("title", "").lower()

            # heuristic conversion
            if any(k in title for k in ["tool", "能力", "skills", "技術", "能力"]):
                slides[i]["type"] = "table"

                slides[i]["columns"] = ["Category", "Details"]
                slides[i]["rows"] = [
                    ["Item 1", "Summary"],
                    ["Item 2", "Summary"],
                    ["Item 3", "Summary"]
                ]

                slides[i].pop("cards", None)

            else:
                slides[i]["type"] = "section"
                slides[i]["name"] = slides[i].get("title", "Section")

                slides[i].pop("cards", None)
                slides[i].pop("title", None)

            print(f"[INFO] diversified slide {i+1} -> {slides[i]['type']}")

    return spec

def beautify_spec(extracted_spec: dict) -> dict:
    prompt = build_prompt(extracted_spec)

    resp = client.chat.completions.create(
        model=OPENAI_MODEL,
        messages=[
            {"role": "developer", "content": "Return only valid JSON matching the schema."},
            {"role": "user", "content": prompt},
        ],
        response_format={
            "type": "json_schema",
            "json_schema": BEAUTIFY_SCHEMA,
        },
        temperature=0.3,
    )

    text = resp.choices[0].message.content
    spec = json.loads(text)
    spec = sanitize_slides(spec)

    slides = spec.get("slides", [])
    print("[INFO] beautified slide count:", len(slides))
    types = [s.get("type") for s in slides]
    print("[INFO] beautified slide types:", types)

    streak = 1
    for i in range(1, len(types)):
        if types[i] == types[i - 1] == "content_3extra":
            streak += 1
        else:
            streak = 1

        if streak >= 3:
            print("[WARN] content_3extra repeats 3 or more times in a row.")
            break

    spec = diversify_layouts(spec)
    return spec


def main():
    extracted_spec = load_extracted_spec(INPUT_SPEC)
    beautified = beautify_spec(extracted_spec)

    with open(OUTPUT_SPEC, "w", encoding="utf-8") as f:
        json.dump(beautified, f, ensure_ascii=False, indent=2)

    print(f"[INFO] Saved beautified spec to: {OUTPUT_SPEC}")


if __name__ == "__main__":
    main()