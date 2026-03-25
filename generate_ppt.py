import json
from datetime import datetime

from ppt_renderer import render_deck
from spec_validator import validate_deck_spec
from spec_normalizer import normalize_beautified_spec

from llm_generate_spec import generate_spec
from extract_ppt_content import extract_ppt_to_spec


TEMPLATE = "template.pptx"
OUTPUT = f"output_{datetime.now().strftime('%H%M%S')}.pptx"

import os

def find_spec_files():
    specs = []
    for f in os.listdir("."):
        if f.lower().endswith(".json"):
            specs.append(f)
    return specs

def load_spec_from_file(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def choose_spec_file():
    specs = find_spec_files()

    if not specs:
        print("[ERROR] No spec JSON files found.")
        return None

    print("\nAvailable spec files:")

    for i, s in enumerate(specs, start=1):
        print(f"{i}. {s}")

    choice = input("\nSelect spec number: ").strip()

    try:
        idx = int(choice) - 1
        return specs[idx]
    except:
        print("Invalid selection.")
        return None

def save_spec_to_file(spec: dict, path: str):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False, indent=2)

def find_ppt_files():
    ppts = []
    for f in os.listdir("."):
        if f.lower().endswith(".pptx") and f != TEMPLATE:
            ppts.append(f)
    return ppts

def choose_ppt_file():
    ppts = find_ppt_files()

    if not ppts:
        print("[ERROR] No PPT files found.")
        return None

    print("\nAvailable PPT files:")

    for i, p in enumerate(ppts, start=1):
        print(f"{i}. {p}")

    choice = input("\nSelect PPT number: ").strip()

    try:
        idx = int(choice) - 1
        return ppts[idx]
    except:
        print("Invalid selection.")
        return None

def main():

    spec = None
        
    mode = input(
        "\nChoose mode:\n"
        "1 = Generate presentation by prompt\n"
        "2 = Beautify existing PPT\n"
        "3 = Use my own spec\n"
        "> "
    ).strip()

    if mode == "1":
        prompt = input("\nDescribe the presentation:\n").strip()
        spec = generate_spec(prompt)
        save_spec_to_file(spec, "deck_spec.generated.json")
        print("[INFO] Saved generated spec to: deck_spec.generated.json")

    elif mode == "2":
    
        ppt_path = choose_ppt_file()

        if ppt_path is None:
            return

        print(f"[INFO] Using PPT: {ppt_path}")

        extracted = extract_ppt_to_spec(ppt_path)

        save_spec_to_file(extracted, "extracted_deck_spec.json")
        print("[INFO] Saved extracted spec to: extracted_deck_spec.json")

        from rule_based_transform import rule_based_transform_spec
        from llm_beautify_spec import rewrite_overflow_fields_with_llm

        transformed = rule_based_transform_spec(extracted)
        save_spec_to_file(transformed, "transformed_deck_spec.json")
        print("[INFO] Saved transformed spec to: transformed_deck_spec.json")

        rewritten = rewrite_overflow_fields_with_llm(transformed)
        save_spec_to_file(rewritten, "beautified_deck_spec.json")
        print("[INFO] Saved beautified spec to: beautified_deck_spec.json")

        spec = rewritten

    elif mode == "3":
        spec_path = choose_spec_file()

        if spec_path is None:
            return

        print(f"[INFO] Using spec: {spec_path}")
        spec = load_spec_from_file(spec_path)

    else:
        print("Invalid mode")
        return

    print("[INFO] Original spec:")
    print(json.dumps(spec, indent=2, ensure_ascii=False))

    spec = normalize_beautified_spec(spec)

    print("[INFO] Normalized spec:")
    print(json.dumps(spec, indent=2, ensure_ascii=False))

    result = validate_deck_spec(spec)

    if result["errors"]:
        print("[ERROR]")
        for e in result["errors"]:
            print("-", e)
        return

    if result["warnings"]:
        print("[WARN]")
        for w in result["warnings"]:
            print("-", w)

    render_deck(
        template_pptx=TEMPLATE,
        deck_spec=result["normalized_spec"],
        out_pptx=OUTPUT
    )

    print(f"\n[INFO] PPT generated: {OUTPUT}")


if __name__ == "__main__":
    main()