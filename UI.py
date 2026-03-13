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

class StepRunner:
    def __init__(self, title: str, total_steps: int, show_logs: bool = True):
        self.title = title
        self.total_steps = total_steps
        self.current_step = 0
        self.show_logs = show_logs

        self.status = st.status(title, expanded=True)
        self.progress = st.progress(0, text="Preparing...")

    def update(self, label: str, kind: str = "info"):
        self.current_step += 1
        pct = int(self.current_step / self.total_steps * 100)
        step_text = f"Step {self.current_step}/{self.total_steps}: {label}"

        self.progress.progress(pct, text=step_text)

        if self.show_logs:
            box_class = {
                "info": "status-info",
                "warn": "status-warn",
                "ok": "status-ok",
                "err": "status-err",
            }.get(kind, "status-info")

            self.status.markdown(
                f'<div class="{box_class}"> {step_text}</div>',
                unsafe_allow_html=True
            )

    def success(self, message: str = "Completed"):
        self.progress.progress(100, text=message)
        self.status.update(label=f" {message}", state="complete", expanded=False)
        st.toast(message)

    def error(self, message: str):
        self.status.update(label=f" {message}", state="error", expanded=True)
        st.error(message)

def run_with_spinner(message: str, func, *args, **kwargs):
    with st.spinner(message):
        return func(*args, **kwargs)

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

def show_result_box(title: str, result: dict | None, clear_key: str):
    if not result:
        return

    with st.container(key=f"result_box_{clear_key}"):
        st.success(title)

        if result.get("summary"):
            st.caption(result["summary"])

        files = result.get("files", [])
        if files:
            cols = st.columns(len(files))
            for col, file_info in zip(cols, files):
                with col:
                    st.download_button(
                        label=f"⬇ {file_info['label']}",
                        data=read_file_bytes(Path(file_info["path"])),
                        file_name=file_info["name"],
                        mime=file_info["mime"],
                        use_container_width=True,
                        key=f"download_{file_info['name']}"
                    )

        if st.button("🗑 Clear result", key=clear_key, use_container_width=True):
            return "clear"

# =========================
# UI Header
# =========================
ensure_template_exists()

# Session state
if "generate_result" not in st.session_state:
    st.session_state.generate_result = None

if "beautify_result" not in st.session_state:
    st.session_state.beautify_result = None

if "render_result" not in st.session_state:
    st.session_state.render_result = None

st.title("📊 AI PPT Tool")
st.caption("Generate, beautify, and render PowerPoint presentations using your company template.")

st.markdown("""
    <style>

    /* ===== Primary Action Buttons ===== */
    .stButton > button[kind="primary"] {
        background-color: #2b6cff;
        color: white;
        border-radius: 8px;
        border: none;
        font-weight: 600;
    }

    .stButton > button[kind="primary"]:hover {
        background-color: #1f4ed8;
    }

    /* ===== Download Buttons ===== */
    div[data-testid="stDownloadButton"] button {
        background-color: #f1f3f5;
        color: #333333;
        border: 1px solid #d0d5dd;
        border-radius: 8px;
    }

    div[data-testid="stDownloadButton"] button:hover {
        background-color: #e6e8eb;
    }

    /* ===== Clear Button (scoped to result box only) ===== */
    [class*="st-key-result_box_clear_"] .stButton > button[kind="secondary"] {
        background-color: #ffeaea;
        color: #c53030;
        border: 1px solid #f5b5b5;
        border-radius: 8px;
    }

    [class*="st-key-result_box_clear_"] .stButton > button[kind="secondary"]:hover {
        background-color: #ffd6d6;
    }


    /* ===== Success Box ===== */
    div[data-testid="stAlert"] {
        border-radius: 10px;
    }

    /* ===== Status Cards ===== */
    .status-info {
        padding: 10px 14px;
        border-radius: 8px;
        background: #eef4ff;
        border-left: 5px solid #2b6cff;
    }

    .status-warn {
        padding: 10px 14px;
        border-radius: 8px;
        background: #fff8e6;
        border-left: 5px solid #f0ad4e;
    }

    .status-ok {
        padding: 10px 14px;
        border-radius: 8px;
        background: #ecf9f0;
        border-left: 5px solid #28a745;
    }

    .status-err {
        padding: 10px 14px;
        border-radius: 8px;
        background: #fff0f0;
        border-left: 5px solid #dc3545;
    }

    </style>
    """, unsafe_allow_html=True)

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
    
    st.write("Template:")
    st.code(str(TEMPLATE_PATH.name))

    show_debug = st.checkbox("Show intermediate JSON", value=True)
    show_logs = st.checkbox("Show step status", value=True)

# =========================
# Mode 1: Generate from prompt
# =========================

if mode == "Generate from prompt":
    st.subheader("Mode 1 · Generate PPT from prompt")

    prompt = st.text_area(
        "Describe the presentation you want",
        height=220,
        placeholder="Example: Create a 5-slide presentation about AI networking trends..."
    )

    run_generate = st.button(
        "Generate PPT",
        type="primary",
        use_container_width=True,
        key="generate_btn"
    )

    if run_generate:
        if not prompt.strip():
            st.error("Please enter a prompt first.")
        else:
            try:
                runner = StepRunner("Generating presentation...", total_steps=4, show_logs=show_logs)

                runner.update("Generating deck spec with LLM")
                generated_spec = run_with_spinner(
                    "LLM is generating slide spec...",
                    generate_spec,
                    prompt
                )

                if show_debug:
                    pretty_json_block("Generated spec", generated_spec)

                runner.update("Normalizing spec")
                normalized_spec = normalize_beautified_spec(generated_spec)

                if show_debug:
                    pretty_json_block("Normalized spec", normalized_spec)

                runner.update("Validating spec", kind="warn")
                result = validate_deck_spec(normalized_spec)
                if result["errors"]:
                    raise ValueError("Validation failed:\n" + "\n".join(result["errors"]))

                show_warnings(result["normalized_spec"])

                runner.update("Rendering PPT", kind="ok")
                out_path = OUTPUT_DIR / f"generated_{timestamp_str()}.pptx"
                run_with_spinner(
                    "Rendering PowerPoint...",
                    render_deck,
                    template_pptx=str(TEMPLATE_PATH),
                    deck_spec=result["normalized_spec"],
                    out_pptx=str(out_path)
                )

                runner.success("PPT ready")

                st.session_state.generate_result = {
                    "summary": f"Latest generated file: {out_path.name}",
                    "files": [
                        {
                            "label": "Download PPT",
                            "path": str(out_path),
                            "name": out_path.name,
                            "mime": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        }
                    ]
                }

            except Exception as e:
                if "runner" in locals():
                    runner.error(f"Generate failed: {e}")
                else:
                    st.error(f"Generate failed: {e}")

                with st.expander("Error details"):
                    st.code(traceback.format_exc())

    clear_action = show_result_box(
        "Generate result ready",
        st.session_state.generate_result,
        clear_key="clear_generate"
    )

    if clear_action == "clear":
        st.session_state.generate_result = None
        st.rerun()
                            
# =========================
# Mode 2: Beautify existing PPT
# =========================

elif mode == "Beautify existing PPT":
    st.subheader("Mode 2 · Beautify an existing PPT")

    uploaded_ppt = st.file_uploader(
        "Upload a PPTX file",
        type=["pptx"],
        key="upload_ppt_mode2"
    )

    run_beautify = st.button(
        "Beautify PPT",
        type="primary",
        use_container_width=True,
        key="beautify_btn"
    )

    if run_beautify:
        if uploaded_ppt is None:
            st.error("Please upload a PPTX file first.")
        else:
            try:
                runner = StepRunner("Beautifying presentation...", total_steps=5, show_logs=show_logs)

                input_ppt_path = OUTPUT_DIR / f"uploaded_{timestamp_str()}_{uploaded_ppt.name}"
                save_uploaded_file(uploaded_ppt, input_ppt_path)

                runner.update("Extracting content from uploaded PPT")
                extracted_spec = run_with_spinner(
                    "Extracting slide content...",
                    extract_ppt_to_spec,
                    str(input_ppt_path)
                )

                if show_debug:
                    pretty_json_block("Extracted spec", extracted_spec)

                runner.update("Beautifying spec with LLM")
                beautified_spec = run_with_spinner(
                    "LLM is improving your presentation structure...",
                    beautify_spec,
                    extracted_spec
                )

                if show_debug:
                    pretty_json_block("Beautified spec", beautified_spec)

                runner.update("Normalizing spec")
                normalized_spec = normalize_beautified_spec(beautified_spec)

                if show_debug:
                    pretty_json_block("Normalized spec", normalized_spec)

                runner.update("Validating spec", kind="warn")
                result = validate_deck_spec(normalized_spec)
                if result["errors"]:
                    raise ValueError("Validation failed:\n" + "\n".join(result["errors"]))

                show_warnings(result["normalized_spec"])

                runner.update("Rendering PPT", kind="ok")
                output_pptx = OUTPUT_DIR / f"beautified_{timestamp_str()}.pptx"
                run_with_spinner(
                    "Rendering beautified PowerPoint...",
                    render_deck,
                    template_pptx=str(TEMPLATE_PATH),
                    deck_spec=result["normalized_spec"],
                    out_pptx=str(output_pptx)
                )

                extracted_json_path = OUTPUT_DIR / f"extracted_spec_{timestamp_str()}.json"
                beautified_json_path = OUTPUT_DIR / f"beautified_spec_{timestamp_str()}.json"
                save_json(extracted_spec, extracted_json_path)
                save_json(beautified_spec, beautified_json_path)

                runner.success("Beautified PPT ready")

                st.session_state.beautify_result = {
                    "summary": f"Latest beautified file: {output_pptx.name}",
                    "files": [
                        {
                            "label": "Download PPT",
                            "path": str(output_pptx),
                            "name": output_pptx.name,
                            "mime": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        },
                        {
                            "label": "Download Extracted Spec",
                            "path": str(extracted_json_path),
                            "name": extracted_json_path.name,
                            "mime": "application/json"
                        },
                        {
                            "label": "Download Beautified Spec",
                            "path": str(beautified_json_path),
                            "name": beautified_json_path.name,
                            "mime": "application/json"
                        }
                    ]
                }

            except Exception as e:
                if "runner" in locals():
                    runner.error(f"Beautify failed: {e}")
                else:
                    st.error(f"Beautify failed: {e}")

                with st.expander("Error details"):
                    st.code(traceback.format_exc())

    clear_action = show_result_box(
        "Beautify result ready",
        st.session_state.beautify_result,
        clear_key="clear_beautify"
    )

    if clear_action == "clear":
        st.session_state.beautify_result = None
        st.rerun()
                        
# =========================
# Mode 3: Render from spec JSON
# =========================

elif mode == "Render from spec JSON":
    st.subheader("Mode 3 · Render PPT from your own spec")

    uploaded_spec = st.file_uploader(
        "Upload a spec JSON file",
        type=["json"],
        key="upload_spec_mode3"
    )

    run_render = st.button(
        "Render PPT",
        type="primary",
        use_container_width=True,
        key="render_btn"
    )

    if run_render:
        if uploaded_spec is None:
            st.error("Please upload a spec JSON file first.")
        else:
            try:
                runner = StepRunner("Rendering from spec...", total_steps=3, show_logs=show_logs)

                runner.update("Loading JSON spec")
                original_spec = load_json(uploaded_spec)

                if show_debug:
                    pretty_json_block("Original spec", original_spec)

                runner.update("Normalizing + validating spec", kind="warn")
                normalized_spec = normalize_beautified_spec(original_spec)
                result = validate_deck_spec(normalized_spec)
                if result["errors"]:
                    raise ValueError("Validation failed:\n" + "\n".join(result["errors"]))

                if show_debug:
                    pretty_json_block("Normalized spec", result["normalized_spec"])

                show_warnings(result["normalized_spec"])

                runner.update("Rendering PPT", kind="ok")
                output_pptx = OUTPUT_DIR / f"spec_rendered_{timestamp_str()}.pptx"
                run_with_spinner(
                    "Rendering PowerPoint...",
                    render_deck,
                    template_pptx=str(TEMPLATE_PATH),
                    deck_spec=result["normalized_spec"],
                    out_pptx=str(output_pptx)
                )

                runner.success("PPT ready")

                st.session_state.render_result = {
                    "summary": f"Latest rendered file: {output_pptx.name}",
                    "files": [
                        {
                            "label": "Download PPT",
                            "path": str(output_pptx),
                            "name": output_pptx.name,
                            "mime": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        }
                    ]
                }

            except Exception as e:
                if "runner" in locals():
                    runner.error(f"Render failed: {e}")
                else:
                    st.error(f"Render failed: {e}")

                with st.expander("Error details"):
                    st.code(traceback.format_exc())

    clear_action = show_result_box(
        "Render result ready",
        st.session_state.render_result,
        clear_key="clear_render"
    )

    if clear_action == "clear":
        st.session_state.render_result = None
        st.rerun()
        
    
