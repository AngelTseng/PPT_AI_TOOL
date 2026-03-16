# PPT AI Tool

A PowerPoint generation tool built with **Streamlit + OpenAI + PowerPoint COM**.
This project converts prompts or existing PPT content into a structured `deck_spec`, then renders a `.pptx` based on your company template.

---

## 1) System Positioning (Frontend / Backend)

Although this is a single Python repository, it can be viewed as two layers:

- **Frontend (UI layer)**: `UI.py` (Streamlit)
  - Provides 3 modes:
    1. Generate PPT from prompt
    2. Beautify uploaded PPT
    3. Render from uploaded spec JSON
  - Displays progress, warnings/errors, and downloadable outputs.

- **Backend (service logic layer)**: multiple Python modules
  - LLM generation / beautification: `llm_generate_spec.py`, `llm_beautify_spec.py`
  - Spec normalization: `spec_normalizer.py`
  - Spec validation: `spec_validator.py`
  - PPT rendering: `ppt_renderer.py` (via Windows PowerPoint COM)
  - Slide/template registry: `slide_registry.py`
  - Existing PPT extraction: `extract_ppt_content.py`

In short: `UI.py` handles user interaction; backend modules handle planning, normalization, validation, and rendering.

---

## 2) Architecture Overview

```text
[User / Browser]
      |
      v
[Streamlit UI: UI.py]
      |
      +--> generate_spec() / beautify_spec()   (OpenAI)
      |
      +--> normalize_beautified_spec()
      |
      +--> validate_deck_spec()
      |
      +--> render_deck()   (PowerPoint COM on Windows)
      v
[Generated PPTX + JSON artifacts]
```

### 2.1 Core Data Model: `deck_spec`

Most workflows revolve around a JSON object:

- `slides: []`
- Each slide includes a `type` (`cover`, `section`, `content_2_a`, `content_4_a`, `flow`, `table`, `end`, etc.)
- Different slide types require different fields (for example, `cards`, `steps`, `columns/rows`)

This format is simultaneously:
- the LLM output target,
- the validator input,
- and the renderer input.

---

## 3) Frontend (UI) Details

### File
- `UI.py`

### Features
- Streamlit-based interactive app
- Download panel for generated artifacts
- Optional intermediate JSON display (debug mode)
- Three workflow modes (Generate / Beautify / Render from JSON)

### Output location
- Default output folder: `ui_outputs/`

---

## 4) Backend Module Responsibilities

### 4.1 LLM Spec Generation
- `llm_generate_spec.py`
- Builds `deck_spec` from user prompt + template summary (`template_map.json`)
- Uses OpenAI structured output (`response_format=json_schema`)

### 4.2 LLM Spec Beautification
- `llm_beautify_spec.py`
- Converts extracted/rough specs into cleaner, renderable specs

### 4.3 Spec Normalization
- `spec_normalizer.py`
- Cleans strings and fills defaults
- Normalizes slide type variants
- Improves low-information content blocks
- Ensures deck contains a first `cover` and final `end` slide

### 4.4 Spec Validation
- `spec_validator.py`
- Checks required fields and structural constraints
- Returns `errors` and `warnings`
- Prevents rendering of invalid specs

### 4.5 PPT Rendering
- `ppt_renderer.py`
- Applies normalized spec content to named template shapes
- Uses `win32com.client` to call desktop PowerPoint
- Handles text, table, and SmartArt filling

### 4.6 Template Metadata
- `slide_registry.py`: maps slide types to template slide indexes and required fields
- `template_map.json`: shape metadata summary used by generation prompt

---

## 5) Environment Requirements

### OS / Runtime
- **Windows** (required for PowerPoint COM rendering)
- Python 3.10+
- Microsoft PowerPoint desktop installed and COM-accessible

### Python packages
From `requirements.txt`:
- `streamlit`
- `pywin32`
- `python-pptx`
- `pydantic`
- `openai`
- `tqdm`

> Note: In Linux containers, spec-related steps can run, but COM-based PPT rendering usually cannot.

---

## 6) Environment Variables

Required:
- `OPENAI_API_KEY`

Optional:
- `OPENAI_MODEL` (default in `config.py`: `gpt-4.1-mini`)

PowerShell example:

```powershell
$env:OPENAI_API_KEY="sk-..."
$env:OPENAI_MODEL="gpt-4.1-mini"
```

CMD example:

```cmd
set OPENAI_API_KEY=sk-...
set OPENAI_MODEL=gpt-4.1-mini
```

---

## 7) Installation and Startup

### 7.1 Install dependencies

```bash
pip install -r requirements.txt
```

Or use the provided batch file:
- `insatll.bat` (current filename in repo)

### 7.2 Start UI

```bash
python -m streamlit run UI.py
```

Or use:
- `run_ui.bat`

Default address:
- `http://localhost:8501`

---

## 8) Recommended Workflow

1. Prepare `template.pptx` with correct shape naming
2. Launch UI
3. Choose a mode:
   - Generate from prompt
   - Beautify uploaded PPT
   - Render from spec JSON
4. Review warnings and intermediate spec if needed
5. Download final PPT/JSON outputs

---

## 9) Template Design Notes

This project strongly depends on stable shape names. Keep key names consistent, for example:
- `Topic`, `speaker_name`
- `agenda_name`, `title`, `item_1`, `content_1`
- `flow_chart_1`, `sheet_1`

If template naming changes, update these together:
- `slide_registry.py`
- renderer/extractor mapping logic
- `template_map.json` (rebuild recommended)

---

## 10) FAQ

### Q1: Why does rendering fail?
Common causes:
- non-Windows environment
- PowerPoint not installed
- template shape names do not match expected names

### Q2: Why is some content still short or sparse?
Spec passes through normalizer and validator first. The system adds defaults where possible, but better input prompt/spec content yields better output.

### Q3: Can I use this as a JSON spec pipeline without rendering?
Yes. You can use `generate_spec / beautify_spec / normalize / validate` without invoking COM rendering.

---

## 11) Main File Index

- `UI.py`: Streamlit frontend entrypoint
- `config.py`: shared runtime config
- `llm_generate_spec.py`: prompt -> spec generation
- `llm_beautify_spec.py`: extracted spec beautification
- `spec_normalizer.py`: spec normalization
- `spec_validator.py`: spec validation
- `ppt_renderer.py`: PowerPoint rendering
- `extract_ppt_content.py`: PPT extraction
- `slide_registry.py`: slide type/index registry
- `template_map.json`: template shape summary

---

## 12) Suggested Future Improvements

- Abstract COM renderer to support non-Windows rendering strategies
- Move validator rules into configurable policy files
- Add unit tests (normalizer/validator/renderer mocks)
- Add CI pipeline (lint + tests + schema checks)

