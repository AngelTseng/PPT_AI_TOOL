import json
from openai import OpenAI
from slide_quality import evaluate_slide_quality

client = OpenAI()
MODEL = "gpt-4.1-mini"


def build_rewrite_prompt(slide: dict, action: str, reasons: list[str]) -> str:
    return f"""
You are rewriting text for an existing PowerPoint slide.

Rules:
- Keep the slide type exactly unchanged.
- Do not add or remove slides.
- Do not convert to another layout.
- Do not invent new numbers, facts, or commitments.
- Preserve the original meaning.
- Keep wording concise and presentation-friendly.
- Title must stay short enough for one line.
- For cards, keep item short and content balanced.
- For flow, keep each step short and parallel.
- For table, keep headers and cells brief and consistent.

Rewrite mode:
- action = {action}
- reasons = {json.dumps(reasons, ensure_ascii=False)}

Action meaning:
- keep: minimal cleanup only
- compress: shorten long text
- enrich: expand overly short text using safe elaboration only
- rebalance: make parallel sections more even and consistent

Safe enrichment rules:
- only restate or clarify the original content
- do not invent metrics, dates, claims, or roadmap promises
- you may add generic explanatory wording such as purpose, method, benefit, or summary

Return valid JSON for this single slide only.

Slide:
{json.dumps(slide, ensure_ascii=False, indent=2)}
"""


def should_call_llm(slide: dict) -> tuple[bool, dict]:
    quality = evaluate_slide_quality(slide)
    action = quality.get("action", "keep")
    return action != "keep", quality


def rewrite_slide_fields(slide: dict) -> dict:
    need_llm, quality = should_call_llm(slide)
    if not need_llm:
        return slide

    prompt = build_rewrite_prompt(
        slide=slide,
        action=quality["action"],
        reasons=quality.get("reasons", [])
    )

    resp = client.chat.completions.create(
        model=MODEL,
        messages=[
            {
                "role": "developer",
                "content": "Return only valid JSON for the input slide. Keep slide type unchanged."
            },
            {"role": "user", "content": prompt},
        ],
        response_format={"type": "json_object"},
        temperature=0.2,
    )

    text = resp.choices[0].message.content
    new_slide = json.loads(text)

    # 強制保留原版型
    new_slide["type"] = slide.get("type")
    return new_slide


def rewrite_overflow_fields_with_llm(spec: dict) -> dict:
    slides = spec.get("slides", []) or []
    out_slides = []

    for slide in slides:
        try:
            out_slides.append(rewrite_slide_fields(slide))
        except Exception as e:
            print("[WARN] rewrite failed, keep original slide:", e)
            out_slides.append(slide)

    return {
        **spec,
        "slides": out_slides
    }