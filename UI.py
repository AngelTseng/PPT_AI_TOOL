import json
import os
import tempfile
import traceback
from datetime import datetime
from pathlib import Path

import streamlit as st

from llm_generate_spec import generate_spec
from llm_beautify_spec import beautify_spec
from extract_ppt_content import extract_ppt_to_spec
from spec_normalizer import normalize_beautified_spec
from spec_validator import validate_deck_spec
from ppt_renderer import render_deck


# =========================
# Basic config
# =========================
st.set_page_config(
    page_title="AI PPT Tool",
    page_icon="📊",
    layout="wide"
)

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "template.pptx"
OUTPUT_DIR = BASE_DIR / "ui_outputs"
OUTPUT_DIR.mkdir(exist_ok=True)


# =========================
# Helpers
# =========================
def timestamp_str() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def save_uploaded_file(uploaded_file, target_path: Path) -> Path:
    with open(target_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return target_path


def save_json(data: dict, target_path: Path) -> Path:
    with open(target_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return target_path


def load_json(uploaded_file) -> dict:
    return json.loads(uploaded_file.getvalue().decode("utf-8"))


def read_file_bytes(path: Path) -> bytes:
    with open(path, "rb") as f:
        return f.read()


def render_from_spec(spec: dict, output_name_prefix: str = "output") -> tuple[dict, Path]:
    """
    normalize -> validate -> render
    returns: (normalized_spec, output_pptx_path)
    """
    normalized_spec = normalize_beautified_spec(spec)
    result = validate_deck_spec(normalized_spec)

    if result["errors"]:
        raise ValueError("Validation failed:\n" + "\n".join(result["errors"]))

    out_path = OUTPUT_DIR / f"{output_name_prefix}_{timestamp_str()}.pptx"

    render_deck(
        template_pptx=str(TEMPLATE_PATH),
        deck_spec=result["normalized_spec"],
        out_pptx=str(out_path)
    )

    return result["normalized_spec"], out_path


def show_warnings(spec: dict):
    result = validate_deck_spec(spec)
    if result["warnings"]:
        st.warning("Warnings:\n\n" + "\n".join(f"- {w}" for w in result["warnings"]))


def pretty_json_block(title: str, data: dict):
    with st.expander(title, expanded=False):
        st.code(json.dumps(data, ensure_ascii=False, indent=2), language="json")


def ensure_template_exists():
    if not TEMPLATE_PATH.exists():
        st.error(f"Cannot find template file: {TEMPLATE_PATH}")
        st.stop()


# =========================
# UI Header
# =========================
ensure_template_exists()

st.title("📊 AI PPT Tool")
st.caption("Generate, beautify, and render PowerPoint presentations using your company template.")

with st.sidebar:
    st.header("Mode")
    mode = st.radio(
        "Choose workflow",
        [
            "Generate from prompt",
            "Beautify existing PPT",
            "Render from spec JSON"
        ]
    )

    st.markdown("---")
    st.write("Template:")
    st.code(str(TEMPLATE_PATH.name))

    show_debug = st.checkbox("Show intermediate JSON", value=True)
    show_logs = st.checkbox("Show step status", value=True)


# =========================
# Mode 1: Generate from prompt
# =========================
if mode == "Generate from prompt":
    st.subheader("Mode 1 · Generate presentation from prompt")

    prompt = st.text_area(
        "Describe the presentation",
        height=180,
        placeholder="Example: Create a 7-slide presentation in Chinese introducing FPGA basics to junior hardware engineers. Use varied layouts and include one table and one flow slide."
    )

    col1, col2 = st.columns([1, 5])
    with col1:
        run_generate = st.button("Generate PPT", use_container_width=True)

    if run_generate:
        if not prompt.strip():
            st.error("Please enter a prompt first.")
        else:
            try:
                if show_logs:
                    st.info("Step 1/4: Generating deck spec with LLM...")

                generated_spec = generate_spec(prompt)

                if show_debug:
                    pretty_json_block("Generated spec", generated_spec)

                if show_logs:
                        st.info("Step 2/4: Normalizing + validating + rendering...")

                normalized_spec, output_pptx = render_from_spec(
                    generated_spec,
                    output_name_prefix="generated"
                )

                if show_debug:
                    pretty_json_block("Normalized spec", normalized_spec)

                show_warnings(normalized_spec)

                if show_logs:
                    st.success("Step 3/4: Rendering complete")
                    st.success("Step 4/4: PPT ready")

                st.success(f"Generated: {output_pptx.name}")

                if show_logs:
                    st.success("Step 3/4: Rendering complete")
                    st.success("Step 4/4: PPT ready")

                st.success(f"Generated: {output_pptx.name}")

                st.download_button(
                    label="Download PPT",
                    data=read_file_bytes(output_pptx),
                    file_name=output_pptx.name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

            except Exception as e:
                st.error(f"Generate failed: {e}")
                with st.expander("Error details"):
                    st.code(traceback.format_exc())


# =========================
# Mode 2: Beautify existing PPT
# =========================
elif mode == "Beautify existing PPT":
    st.subheader("Mode 2 · Beautify an existing PPT")

    uploaded_ppt = st.file_uploader(
        "Upload a PPTX file",
        type=["pptx"]
    )

    col1, col2 = st.columns([1, 5])
    with col1:
        run_beautify = st.button("Beautify PPT", use_container_width=True)

    if run_beautify:
        if uploaded_ppt is None:
            st.error("Please upload a PPTX file first.")
        else:
            try:
                input_ppt_path = OUTPUT_DIR / f"uploaded_{timestamp_str()}_{uploaded_ppt.name}"
                save_uploaded_file(uploaded_ppt, input_ppt_path)

                if show_logs:
                    st.info("Step 1/5: Extracting content from uploaded PPT...")

                extracted_spec = extract_ppt_to_spec(str(input_ppt_path))

                if show_debug:
                    pretty_json_block("Extracted spec", extracted_spec)

                if show_logs:
                    st.info("Step 2/5: Beautifying spec with LLM...")

                beautified_spec = beautify_spec(extracted_spec)

                if show_debug:
                    pretty_json_block("Beautified spec", beautified_spec)

                if show_logs:
                    st.info("Step 3/5: Normalizing spec...")

                normalized_spec, output_pptx = render_from_spec(
                    beautified_spec,
                    output_name_prefix="beautified"
                )

                if show_debug:
                    pretty_json_block("Normalized spec", normalized_spec)

                show_warnings(normalized_spec)

                # save intermediate files for convenience
                extracted_json_path = OUTPUT_DIR / f"extracted_spec_{timestamp_str()}.json"
                beautified_json_path = OUTPUT_DIR / f"beautified_spec_{timestamp_str()}.json"
                save_json(extracted_spec, extracted_json_path)
                save_json(beautified_spec, beautified_json_path)

                if show_logs:
                    st.success("Step 4/5: Rendering complete")
                    st.success("Step 5/5: Beautified PPT ready")

                st.success(f"Generated: {output_pptx.name}")

                c1, c2, c3 = st.columns(3)

                with c1:
                    st.download_button(
                        label="Download PPT",
                        data=read_file_bytes(output_pptx),
                        file_name=output_pptx.name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )

                with c2:
                    st.download_button(
                        label="Download Extracted Spec",
                        data=read_file_bytes(extracted_json_path),
                        file_name=extracted_json_path.name,
                        mime="application/json"
                    )

                with c3:
                    st.download_button(
                        label="Download Beautified Spec",
                        data=read_file_bytes(beautified_json_path),
                        file_name=beautified_json_path.name,
                        mime="application/json"
                    )

            except Exception as e:
                st.error(f"Beautify failed: {e}")
                with st.expander("Error details"):
                    st.code(traceback.format_exc())


# =========================
# Mode 3: Render from spec JSON
# =========================
elif mode == "Render from spec JSON":
    st.subheader("Mode 3 · Render PPT from your own spec")

    uploaded_spec = st.file_uploader(
        "Upload a spec JSON file",
        type=["json"]
    )

    col1, col2 = st.columns([1, 5])
    with col1:
        run_render = st.button("Render PPT", use_container_width=True)

    if run_render:
        if uploaded_spec is None:
            st.error("Please upload a spec JSON file first.")
        else:
            try:
                original_spec = load_json(uploaded_spec)

                if show_debug:
                    pretty_json_block("Original spec", original_spec)

                if show_logs:
                    st.info("Step 1/3: Normalizing spec...")

                normalized_spec, output_pptx = render_from_spec(
                    original_spec,
                    output_name_prefix="spec_rendered"
                )

                if show_debug:
                    pretty_json_block("Normalized spec", normalized_spec)

                show_warnings(normalized_spec)

                if show_logs:
                    st.success("Step 2/3: Rendering complete")
                    st.success("Step 3/3: PPT ready")

                st.success(f"Generated: {output_pptx.name}")

                st.download_button(
                    label="Download PPT",
                    data=read_file_bytes(output_pptx),
                    file_name=output_pptx.name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

            except Exception as e:
                st.error(f"Render failed: {e}")
                with st.expander("Error details"):
                    st.code(traceback.format_exc())