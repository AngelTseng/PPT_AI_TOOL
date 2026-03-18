import json
from pathlib import Path
from openai import OpenAI

client = OpenAI()

MODEL = "gpt-4.1-mini"

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_MAP_PATH = BASE_DIR / "template_map.json"


DECK_SPEC_SCHEMA = {
    "name": "deck_spec",
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
                            "enum": [
                                "cover",
                                "agenda",
                                "section",
                                "content_2",
                                "content_3extra",
                                "content_4",
                                "table",
                                "flow",
                                "end"
                            ]
                        },
                        "topic": {"type": "string"},
                        "speaker": {"type": "string"},
                        "title": {"type": "string"},
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
                            "if": {"properties": {"type": {"const": "content_2"}}},
                            "then": {"required": ["type", "title", "cards"]}
                        },
                        {
                            "if": {"properties": {"type": {"const": "content_4"}}},
                            "then": {"required": ["type", "title", "cards"]}
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


def load_template_map() -> list:
    if not TEMPLATE_MAP_PATH.exists():
        return []

    with open(TEMPLATE_MAP_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def build_template_summary(template_map: list) -> str:
    """
    把 template_map.json 壓縮成比較適合給 LLM 的摘要
    """
    lines = []
    for slide in template_map:
        slide_index = slide.get("slide_index")
        shapes = slide.get("shapes", [])

        shape_names = []
        for shp in shapes:
            name = shp.get("name", "").strip()
            if name:
                shape_names.append(name)

        if shape_names:
            lines.append(f"Slide {slide_index}: " + ", ".join(shape_names))

    return "\n".join(lines)

def build_developer_prompt(template_summary: str) -> str:
    return f"""
You are a PowerPoint deck planning assistant.

Your job is to generate a valid deck_spec JSON for a company-template PowerPoint tool.

Supported slide types:
- cover
- agenda
- section
- content_2
- content_3extra
- content_4
- table
- flow
- end

Rules:
- agenda.items: max 5
- content_2.cards: max 2
- content_3extra.cards: max 3
- content_4.cards: max 4
- flow.steps: at least 2
- table.columns must not be empty
- table.rows must not be empty
- use concise PowerPoint wording
- do not write long report-style paragraphs
- output JSON only
- no markdown
- no explanation text

Deck planning rules:
- Do not make the deck visually monotonous.
- Use at least 3 different slide types in a normal deck.
- For decks with 5 or more slides, use at least 3 distinct slide types excluding cover and end.
- For decks with 7 or more slides, use at least 4 distinct slide types excluding cover and end.
- Avoid repeating the same content slide type more than 2 times in a row unless necessary.
- Do not overuse content_3extra.
- Do not default every explanatory slide to content_3extra.

Agenda planning rules:
- Agenda is optional.
- Agenda may contain fewer than 5 items.
- If agenda is included, each agenda item should be followed by at least one explanatory slide.
- Do not include agenda if it makes the deck unnecessarily long.
- The number of agenda items should match the number of major content sections.

Slide usage guidance:
- Use content_2 for two grouped ideas, two parallel concepts, two-column explanations, or two balanced highlights.
- Use content_3extra for exactly three grouped ideas or three parallel highlights.
- Use content_4 for four grouped ideas, four capabilities, four categorized points, or four-item summaries.
- Use table for comparisons, specifications, structured facts, grouped responsibilities, categories, or tool/capability summaries.
- Use flow for processes, sequences, collaboration stages, development lifecycle, or learning paths.
- Use section to break major topics and improve presentation rhythm.

Layout adaptation rules:
- If content naturally fits 2 grouped ideas, prefer content_2 instead of forcing it into content_3extra.
- If content naturally fits 4 grouped ideas, prefer content_4 or table instead of forcing it into content_3extra.
- Only use content_3extra when the content truly fits a 3-point grouped layout.
- If a slide looks like a chapter heading or transition, use section.
- If content is too long for one slide, split it into multiple slides when needed.

Section coverage rules:
- Every section slide must be immediately followed by at least one explanatory content slide.
- Do not place two section slides back-to-back.
- Do not create a section slide unless there is at least one following content slide dedicated to that section.
- If a topic only needs one content slide, prefer using that content slide directly instead of adding a separate section slide.

Good presentation rhythm examples:

Example 5-slide deck:
cover
agenda
content_2 or content_3extra
flow or table
end

Example 6-slide deck:
cover
agenda
section
content_2 or content_3extra
flow or table
end

Example 7-slide deck:
cover
agenda
section
content_2
table
flow
end

Example 8-slide deck:
cover
agenda
section
content_2
content_4
table
flow
end

Always prefer varied slide layouts instead of repeating the same type.

Here is the template structure summary:
{template_summary}

Use the template structure to choose appropriate slide types and content layout.
"""

def sanitize_slides(spec: dict) -> dict:
    allowed = {
        "cover", "agenda", "section",
        "content_2", "content_3extra", "content_4",
        "table", "flow", "end"
    }

    slides = spec.get("slides", [])
    cleaned = []

    for s in slides:
        if isinstance(s, dict) and s.get("type") in allowed:
            cleaned.append(s)
        else:
            print("[WARN] Dropping malformed slide:", s)

    spec["slides"] = cleaned
    return spec

def generate_spec(user_prompt: str) -> dict:
    template_map = load_template_map()
    template_summary = build_template_summary(template_map)
    developer_prompt = build_developer_prompt(template_summary)

    try:
        resp = client.chat.completions.create(
            model=MODEL,
            messages=[
                {"role": "developer", "content": developer_prompt},
                {"role": "user", "content": user_prompt},
            ],
            response_format={
                "type": "json_schema",
                "json_schema": DECK_SPEC_SCHEMA,
            },
            temperature=0.3,
        )

        text = resp.choices[0].message.content

        print("\n[DEBUG] LLM RAW OUTPUT]\n")
        print(text)
        print("\n")

        spec = json.loads(text)
        spec = sanitize_slides(spec)
        
        slides = spec.get("slides", [])
        normalized = []

        for s in slides:
            if isinstance(s, str):
                normalized.append({"type": s})
            else:
                normalized.append(s)

        spec["slides"] = normalized

        types = [s.get("type") for s in normalized]
        print("[INFO] slide count:", len(normalized))
        print("[INFO] slide types:", types)

        return spec

    except json.JSONDecodeError as e:
        raise ValueError(f"Model output is not valid JSON: {e}")

    except Exception as e:
        raise RuntimeError(f"Failed to generate deck spec: {e}")

