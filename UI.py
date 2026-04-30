import json
import os
import tempfile
import traceback
from datetime import datetime
from pathlib import Path
import streamlit as st
from llm_generate_spec import generate_spec
from extract_ppt_content import extract_ppt_to_spec
from spec_normalizer import normalize_beautified_spec
from spec_validator import validate_deck_spec
from ppt_renderer import render_deck
from llm_beautify_spec import beautify_spec
from extract_word_content import extract_word_to_payload
from extract_excel_content import extract_excel_to_payload
from excel_to_spec import excel_payload_to_spec
from extract_pdf_content import extract_pdf_to_payload
import platform  
import shutil   
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
def export_ppt_preview_images(pptx_path: Path) -> list[str]:
    """
    Export a PPTX into per-slide JPG preview images (Windows + PowerPoint COM only).
    Returns image paths; returns empty list when preview export is unavailable.
    """
    if platform.system().lower() != "windows":
        return []
    try:
        import pythoncom
        import win32com.client as win32
    except Exception:
        return []
    preview_dir = OUTPUT_DIR / f"{pptx_path.stem}_preview"
    if preview_dir.exists():
        shutil.rmtree(preview_dir, ignore_errors=True)
    preview_dir.mkdir(parents=True, exist_ok=True)
    app = None
    pres = None
    try:
        pythoncom.CoInitialize()
        app = win32.Dispatch("PowerPoint.Application")
        app.Visible = True
        pres = app.Presentations.Open(str(pptx_path.resolve()), WithWindow=False)
        # 17 = ppSaveAsJPG
        pres.SaveAs(str(preview_dir.resolve()), 17)
    except Exception:
        return []
    finally:
        if pres is not None:
            pres.Close()
        if app is not None:
            app.Quit()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    # 只抓一次，並做去重
    images = []
    seen = set()
    jpg_files = [p for p in preview_dir.iterdir() if p.is_file() and p.suffix.lower() == ".jpg"]
    def _slide_sort_key(path_obj: Path):
        digits = "".join(ch for ch in path_obj.stem if ch.isdigit())
        return (int(digits) if digits else 10**9, path_obj.name.lower())
    for p in sorted(jpg_files, key=_slide_sort_key):
        key = str(p.resolve()).lower()
        if key not in seen:
            seen.add(key)
            images.append(str(p))
    return images
def show_warnings(result: dict):
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
        preview_images = result.get("preview_images", [])
        if preview_images:
            preview_state_key = f"{clear_key}_preview_index"
            if preview_state_key not in st.session_state:
                st.session_state[preview_state_key] = 0
            max_idx = len(preview_images) - 1
            current_idx = st.session_state[preview_state_key]
            if current_idx < 0:
                current_idx = 0
            if current_idx > max_idx:
                current_idx = max_idx
            st.session_state[preview_state_key] = current_idx
            st.markdown("#### 👀 PPT Preview")
            st.caption("固定預覽視窗，可用上一頁/下一頁切換。")
            st.markdown(
                "<div style='border:1px solid #d0d5dd; border-radius:10px; padding:12px; background:#fafafa;'>",
                unsafe_allow_html=True,
            )
            st.image(
                preview_images[st.session_state[preview_state_key]],
                caption=f"Slide {st.session_state[preview_state_key] + 1}",
                use_container_width=True,
            )
            st.markdown("</div>", unsafe_allow_html=True)
            nav_left, nav_mid, nav_right = st.columns([1, 2, 1])
            with nav_left:
                if st.button("⬅ 上一頁", key=f"{clear_key}_prev", use_container_width=True, type="primary"):
                    st.session_state[preview_state_key] = max(0, st.session_state[preview_state_key] - 1)
                    st.rerun()
            with nav_mid:
                st.markdown(
                    f"<div style='text-align:center; font-weight:600; color:#2b6cff;'>Slide {st.session_state[preview_state_key] + 1} / {len(preview_images)}</div>",
                    unsafe_allow_html=True,
                )
            with nav_right:
                if st.button("下一頁 ➡", key=f"{clear_key}_next", use_container_width=True, type="primary"):
                    st.session_state[preview_state_key] = min(max_idx, st.session_state[preview_state_key] + 1)
                    st.rerun()
        if st.button("🗑 Clear result", key=clear_key, use_container_width=True):
            preview_state_key = f"{clear_key}_preview_index"
            if preview_state_key in st.session_state:
                del st.session_state[preview_state_key]
            return "clear"
def build_word_prompt(word_payload: dict, duration_mode: str) -> str:
    if duration_mode == "10-15 min":
        min_slides, target, max_slides = 12, 14, 20
    else:
        min_slides, target, max_slides = 22, 26, 35
    return f"""
You are generating a PowerPoint deck from a Word report.
Report title:
{word_payload.get("title", "")}
Report content:
{word_payload.get("raw_text", "")}
Presentation constraints:
- Duration: {duration_mode}
- Slide count minimum: {min_slides}
- Preferred target: {target}
- Maximum: {max_slides}
Rules:
- Must include cover and end slide
- Each slide should contain ONE key idea
- Split long content into multiple slides
- Avoid empty slides
- Avoid overly dense slides
- Use structured bullet points
"""
# =========================
# UI Header
# =========================
ensure_template_exists()
# Session state
if "generate_result" not in st.session_state:
    st.session_state.generate_result = None
if "beautify_result" not in st.session_state:
    st.session_state.beautify_result = None
if "word_generate_result" not in st.session_state:
    st.session_state.word_generate_result = None
if "excel_generate_result" not in st.session_state:
    st.session_state.excel_generate_result = None
if "pdf_generate_result" not in st.session_state:
    st.session_state.pdf_generate_result = None
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
    st.header("Settings")
    st.write("Template:")
    st.code(str(TEMPLATE_PATH.name))
    with st.expander("Advanced / Debug", expanded=False):
        show_debug = st.checkbox("Show intermediate JSON", value=False)
        show_logs = st.checkbox("Show step status", value=False)
        st.caption("Template information")
        st.code(str(TEMPLATE_PATH))


def summarize_uploaded_files(file_groups: dict[str, list]):
    labels = {
        "word": "Word content",
        "excel": "Excel data",
        "pdf": "PDF reference",
        "pptx": "PPT reference / existing deck",
        "txt": "Text notes",
    }
    st.markdown("### Detected files summary")
    has_any = False
    for group_key in ["word", "excel", "pdf", "pptx", "txt"]:
        for uploaded in file_groups[group_key]:
            has_any = True
            st.write(f"- {uploaded.name} → {labels[group_key]}")
    if not has_any:
        st.caption("No files uploaded yet.")


def build_integrated_prompt(base_prompt: str, payload_texts: list[str]) -> str:
    sections = []
    if base_prompt.strip():
        sections.append(f"User requirement:\n{base_prompt.strip()}")
    if payload_texts:
        sections.append("Source materials:\n" + "\n\n".join(payload_texts))
    if not sections:
        sections.append("Create a professional presentation based on uploaded materials.")
    return "\n\n".join(sections)


st.subheader("Create PPT with unified input flow")
prompt = st.text_area(
    "1) Describe your presentation",
    height=220,
    placeholder="Example: Create a 10-slide technical presentation based on the uploaded Word report and Excel data."
)
uploaded_files = st.file_uploader(
    "2) Upload materials",
    type=["pptx", "docx", "pdf", "xlsx", "xls", "txt"],
    accept_multiple_files=True,
    key="upload_unified_inputs"
)
st.caption("Support: PowerPoint, Word, PDF, Excel, TXT. You may upload multiple files and the tool will organize them into one PPT.")

file_groups = {
    "pptx": [],
    "word": [],
    "pdf": [],
    "excel": [],
    "txt": [],
    "unsupported": [],
}

for uploaded in (uploaded_files or []):
    ext = uploaded.name.lower().split(".")[-1] if "." in uploaded.name else ""
    if ext == "pptx":
        file_groups["pptx"].append(uploaded)
    elif ext == "docx":
        file_groups["word"].append(uploaded)
    elif ext == "pdf":
        file_groups["pdf"].append(uploaded)
    elif ext in ["xlsx", "xls"]:
        file_groups["excel"].append(uploaded)
    elif ext == "txt":
        file_groups["txt"].append(uploaded)
    else:
        file_groups["unsupported"].append(uploaded)

if file_groups["unsupported"]:
    st.warning("Unsupported files are skipped: " + ", ".join(f.name for f in file_groups["unsupported"]))

summarize_uploaded_files(file_groups)

has_prompt = bool(prompt.strip())
processable_count = sum(len(file_groups[k]) for k in ["pptx", "word", "pdf", "excel", "txt"])
only_single_pptx = len(file_groups["pptx"]) == 1 and processable_count == 1

if only_single_pptx and not has_prompt:
    workflow = "beautify_existing_ppt"
    st.info("Detected workflow: Beautify existing PPT")
else:
    workflow = "generate_integrated_ppt"
    st.info("Detected workflow: Generate integrated PPT")

run_main = st.button("Generate PPT", type="primary", use_container_width=True, key="generate_unified_btn")

if run_main:
    if processable_count == 0 and not has_prompt:
        st.error("Please enter requirements or upload at least one supported file.")
    else:
        try:
            if workflow == "beautify_existing_ppt":
                uploaded_ppt = file_groups["pptx"][0]
                runner = StepRunner("Beautifying PPT.", total_steps=6, show_logs=show_logs)
                input_ppt_path = OUTPUT_DIR / f"uploaded_{timestamp_str()}_{uploaded_ppt.name}"
                save_uploaded_file(uploaded_ppt, input_ppt_path)
                runner.update("Extracting content from PPT")
                extracted_spec = run_with_spinner("Reading PPT content.", extract_ppt_to_spec, str(input_ppt_path))
                if show_debug:
                    pretty_json_block("Extracted spec", extracted_spec)
                runner.update("Beautifying content with LLM")
                beautified_spec = run_with_spinner("LLM is improving the slide content.", beautify_spec, extracted_spec)
                if show_debug:
                    pretty_json_block("Beautified spec", beautified_spec)
                runner.update("Normalizing spec")
                normalized_spec = normalize_beautified_spec(beautified_spec)
                runner.update("Validating spec", kind="warn")
                result = validate_deck_spec(normalized_spec)
                if result["errors"]:
                    raise ValueError("Validation failed:\n" + "\n".join(result["errors"]))
                runner.update("Rendering PPT", kind="ok")
                output_pptx = OUTPUT_DIR / f"beautified_{timestamp_str()}.pptx"
                run_with_spinner("Rendering PowerPoint.", render_deck, template_pptx=str(TEMPLATE_PATH), deck_spec=result["normalized_spec"], out_pptx=str(output_pptx))
                runner.success("Beautified PPT ready")
            else:
                runner = StepRunner("Generating integrated PPT.", total_steps=6, show_logs=show_logs)
                payload_sections = []
                runner.update("Processing uploaded materials")
                for uploaded_docx in file_groups["word"]:
                    pth = OUTPUT_DIR / f"uploaded_{timestamp_str()}_{uploaded_docx.name}"
                    save_uploaded_file(uploaded_docx, pth)
                    word_payload = run_with_spinner("Reading Word content.", extract_word_to_payload, str(pth))
                    payload_sections.append(f"[WORD] {uploaded_docx.name}\n{word_payload.get('raw_text', '')}")
                for uploaded_pdf in file_groups["pdf"]:
                    pth = OUTPUT_DIR / f"uploaded_{timestamp_str()}_{uploaded_pdf.name}"
                    save_uploaded_file(uploaded_pdf, pth)
                    pdf_payload = run_with_spinner("Reading PDF content.", extract_pdf_to_payload, str(pth))
                    payload_sections.append(f"[PDF] {uploaded_pdf.name}\n{pdf_payload.get('raw_text', '')}")
                for uploaded_excel in file_groups["excel"]:
                    pth = OUTPUT_DIR / f"uploaded_{timestamp_str()}_{uploaded_excel.name}"
                    save_uploaded_file(uploaded_excel, pth)
                    excel_payload = run_with_spinner("Reading Excel content.", extract_excel_to_payload, str(pth))
                    payload_sections.append(f"[EXCEL] {uploaded_excel.name}\n" + json.dumps(excel_payload, ensure_ascii=False)[:8000])
                for uploaded_txt in file_groups["txt"]:
                    txt_content = uploaded_txt.getvalue().decode("utf-8", errors="ignore")
                    payload_sections.append(f"[TXT] {uploaded_txt.name}\n{txt_content}")
                for uploaded_ppt in file_groups["pptx"]:
                    pth = OUTPUT_DIR / f"uploaded_{timestamp_str()}_{uploaded_ppt.name}"
                    save_uploaded_file(uploaded_ppt, pth)
                    ppt_spec = run_with_spinner("Reading PPT content.", extract_ppt_to_spec, str(pth))
                    payload_sections.append(f"[PPTX] {uploaded_ppt.name}\n" + json.dumps(ppt_spec, ensure_ascii=False)[:8000])
                if show_debug:
                    pretty_json_block("Material summary", {"sections": payload_sections})
                runner.update("Generating draft spec with LLM")
                integrated_prompt = build_integrated_prompt(prompt, payload_sections)
                generated_spec = run_with_spinner("LLM is generating slide spec.", generate_spec, integrated_prompt)
                if show_debug:
                    pretty_json_block("Generated draft spec", generated_spec)
                runner.update("Beautifying spec with LLM")
                beautified_spec = run_with_spinner("LLM is refining the slide spec.", beautify_spec, generated_spec)
                runner.update("Normalizing spec")
                normalized_spec = normalize_beautified_spec(beautified_spec)
                runner.update("Validating spec", kind="warn")
                result = validate_deck_spec(normalized_spec)
                if result["errors"]:
                    raise ValueError("Validation failed:\n" + "\n".join(result["errors"]))
                runner.update("Rendering PPT", kind="ok")
                output_pptx = OUTPUT_DIR / f"integrated_{timestamp_str()}.pptx"
                run_with_spinner("Rendering PowerPoint.", render_deck, template_pptx=str(TEMPLATE_PATH), deck_spec=result["normalized_spec"], out_pptx=str(output_pptx))
                runner.success("Integrated PPT ready")

            st.session_state.generate_result = {
                "summary": f"Latest generated file: {output_pptx.name}",
                "files": [{
                    "label": "Download PPT",
                    "path": str(output_pptx),
                    "name": output_pptx.name,
                    "mime": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                }],
                "preview_images": export_ppt_preview_images(output_pptx)
            }
        except Exception as e:
            if "runner" in locals():
                runner.error(f"Workflow failed: {e}")
            else:
                st.error(f"Workflow failed: {e}")
            with st.expander("Error details"):
                st.code(traceback.format_exc())

clear_action = show_result_box("PPT result ready", st.session_state.generate_result, clear_key="clear_unified_generate")
if clear_action == "clear":
    st.session_state.generate_result = None
    st.rerun()
