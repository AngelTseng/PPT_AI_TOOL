import os
import json
import shutil
import pythoncom
import win32com.client as win32

from slide_registry import (
    FLOW_TEMPLATE_INDEX,
    SLIDE_REGISTRY,
    normalize_registry_type,
    resolve_flow_template_key,
    resolve_content_3_template_key,
)

from renderers_content import (
    render_content_image,
    render_content_2_a,
    render_content_2_b,
    render_content_2_c,
    render_content_3extra,
    render_content_4_a,
    render_content_4_b,
)

SLIDE_RENDERERS = {}

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_MAP_PATH = os.path.join(BASE_DIR, "template_map.json")

MsoAutoSizeTextToFitShape = 2
MsoTrue = -1
MsoFalse = 0


# ============================================================
# Template map / role helpers
# ============================================================

def _load_template_map():
    try:
        with open(TEMPLATE_MAP_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {}

    slide_map = {}
    for slide_info in data:
        slide_index = slide_info.get("slide_index")
        if not slide_index:
            continue

        by_name = {}
        for shp in slide_info.get("shapes", []):
            name = str(shp.get("name", "")).strip()
            if not name:
                continue
            by_name[name.lower()] = shp

        slide_map[int(slide_index)] = {
            "detected_type": slide_info.get("detected_type"),
            "shapes": by_name,
        }
    return slide_map


TEMPLATE_ROLE_MAP = _load_template_map()


# ============================================================
# Registry helpers
# ============================================================

def register_renderer(name):
    def wrapper(func):
        SLIDE_RENDERERS[name] = func
        return func
    return wrapper

# content renderers from renderers_content.py
SLIDE_RENDERERS["content_image"] = render_content_image
SLIDE_RENDERERS["content_2_a"] = render_content_2_a
SLIDE_RENDERERS["content_2_b"] = render_content_2_b
SLIDE_RENDERERS["content_2_c"] = render_content_2_c
SLIDE_RENDERERS["content_3extra"] = render_content_3extra
SLIDE_RENDERERS["content_4_a"] = render_content_4_a
SLIDE_RENDERERS["content_4_b"] = render_content_4_b

# ============================================================
# PowerPoint / slide helpers
# ============================================================

def duplicate_to_presentation(src_slide, dst_pres, max_retries: int = 10):
    import time

    last_err = None

    for attempt in range(1, max_retries + 1):
        try:
            src_slide.Copy()
            pythoncom.PumpWaitingMessages()
            time.sleep(0.2 * attempt)

            paste_result = dst_pres.Slides.Paste(dst_pres.Slides.Count + 1)
            pythoncom.PumpWaitingMessages()
            time.sleep(0.1)

            try:
                return paste_result.Item(1)
            except Exception:
                return dst_pres.Slides(dst_pres.Slides.Count)

        except Exception as e:
            last_err = e
            print(f"[WARN] Paste failed attempt {attempt}/{max_retries}: {e}")
            pythoncom.PumpWaitingMessages()
            time.sleep(0.3 * attempt)

    raise RuntimeError(f"Failed to paste slide after {max_retries} retries: {last_err}")

def delete_all_slides(pres):
    for i in range(pres.Slides.Count, 0, -1):
        pres.Slides(i).Delete()


def shape_by_name(slide, name: str):
    for i in range(1, slide.Shapes.Count + 1):
        shp = slide.Shapes(i)
        try:
            if shp.Name == name:
                return shp
        except Exception:
            pass
    return None


def _find_template_slide_index_by_shape(src_pres, shape_name: str):
    target = str(shape_name).strip().lower()
    for i in range(1, src_pres.Slides.Count + 1):
        slide = src_pres.Slides(i)
        for j in range(1, slide.Shapes.Count + 1):
            shp = slide.Shapes(j)
            try:
                if str(shp.Name).strip().lower() == target:
                    return i
            except Exception:
                continue
    return None


# ============================================================
# Flow variant helpers
# ============================================================

def choose_flow_variant(slide_spec: dict) -> str:
    explicit = str(slide_spec.get("variant", "")).strip().lower()
    if explicit in FLOW_TEMPLATE_INDEX:
        return explicit

    steps = [str(s).strip() for s in slide_spec.get("steps", []) if str(s).strip()]
    max_len = max((len(s) for s in steps), default=0)

    # 第三種固定給長文字
    if max_len >= 22:
        return "flow_chart_3"

    title = str(slide_spec.get("title", "")).lower()
    text = " ".join([title] + [s.lower() for s in steps])

    loop_keywords = [
        "iteration", "iterate", "feedback", "cycle", "optimize", "improve", "review", 
        "迭代", "循環", "回饋", "優化", "改善", "檢討"
    ]

    if any(kw in text for kw in loop_keywords):
        return "flow_chart_2"

    return "flow_chart_1"


def _resolve_flow_prefer_name(slide, slide_spec):
    variant = choose_flow_variant(slide_spec)

    # If the chosen variant does not exist on this slide, use first SmartArt.
    shp = shape_by_name(slide, variant)
    try:
        if shp is not None and getattr(shp, "HasSmartArt", False):
            return variant
    except Exception:
        pass

    for i in range(1, slide.Shapes.Count + 1):
        candidate = slide.Shapes(i)
        try:
            if getattr(candidate, "HasSmartArt", False):
                return str(candidate.Name)
        except Exception:
            pass

    return variant


# ============================================================
# Text fit / overlap helpers
# ============================================================

def _fit_text_to_shape(shape):
    """
    Enable wrap and disable PowerPoint auto-resize.
    Keep original template font size first.
    Further shrinking is handled later.
    """
    try:
        tf2 = shape.TextFrame2
        tf2.WordWrap = MsoTrue
        tf2.AutoSize = 0
    except Exception:
        pass

    try:
        tf = shape.TextFrame
        tf.WordWrap = True
    except Exception:
        pass


def _set_wordwrap_and_autosize(shape, no_wrap: bool = False):
    """
    Control wrap behavior explicitly.
    no_wrap=True  -> single-line preference
    no_wrap=False -> allow wrapping
    """
    wrap = MsoFalse if no_wrap else MsoTrue

    try:
        shape.TextFrame2.WordWrap = wrap
        shape.TextFrame2.AutoSize = 0
    except Exception:
        pass

    try:
        shape.TextFrame.WordWrap = bool(not no_wrap)
    except Exception:
        pass


def _clamp_shape_within_slide(slide, shape):
    """
    Generic clamp helper.
    NOTE:
    This function may move shape.Top / shape.Left.
    Do NOT use it inside text-growth logic if you want
    the textbox to keep its original anchor position.
    """
    try:
        sw = float(slide.Parent.PageSetup.SlideWidth)
        sh = float(slide.Parent.PageSetup.SlideHeight)

        left = float(shape.Left)
        top = float(shape.Top)
        width = float(shape.Width)
        height = float(shape.Height)

        max_left = max(0.0, sw - width)
        max_top = max(0.0, sh - height)

        new_left = min(max(left, 0.0), max_left)
        new_top = min(max(top, 0.0), max_top)

        if abs(new_left - left) > 0.1:
            shape.Left = new_left
        if abs(new_top - top) > 0.1:
            shape.Top = new_top
    except Exception:
        pass


def _expand_shape_height_to_fit_text(
    slide,
    shape,
    max_extra_height: float = 220.0,
    step: float = 12.0,
    single_line: bool = False,
):
    """
    Expand textbox downward only.
    Keep original Top / Left fixed.
    Never call generic clamp here, otherwise the box may be pushed upward.
    """
    try:
        _set_wordwrap_and_autosize(shape, no_wrap=single_line)
    except Exception:
        pass

    try:
        tr2 = shape.TextFrame2.TextRange
        if tr2.Length <= 0:
            return
    except Exception:
        return

    try:
        original_left = float(shape.Left)
        original_top = float(shape.Top)
        original_height = float(shape.Height)
        slide_height = float(slide.Parent.PageSetup.SlideHeight)
    except Exception:
        return

    # Only allow growth downward within slide bottom bound.
    max_height_by_slide = max(0.0, slide_height - original_top)
    max_height = min(original_height + max_extra_height, max_height_by_slide)

    safety = 30
    while safety > 0:
        safety -= 1

        try:
            # Re-fetch every round so COM layout/bounds can refresh
            tr2 = shape.TextFrame2.TextRange
            if tr2.Length <= 0:
                break

            bw = float(tr2.BoundWidth)
            bh = float(tr2.BoundHeight)
            w = float(shape.Width)
            h = float(shape.Height)
        except Exception:
            break

        fits = (bw <= w and bh <= h)

        if single_line:
            try:
                line_count = tr2.Lines().Count
            except Exception:
                line_count = 1
            fits = fits and (line_count <= 1)

        if fits:
            break

        try:
            current_height = float(shape.Height)
            new_height = min(current_height + step, max_height)

            if new_height <= current_height + 0.1:
                break

            shape.Height = new_height

            # Force anchor position to stay fixed
            if abs(float(shape.Top) - original_top) > 0.1:
                shape.Top = original_top
            if abs(float(shape.Left) - original_left) > 0.1:
                shape.Left = original_left

        except Exception:
            break


def _text_still_overflows_shape(shape, single_line: bool = False) -> bool:
    """
    Check whether text still overflows its own textbox.
    """
    try:
        tr2 = shape.TextFrame2.TextRange
        if tr2.Length <= 0:
            return False

        bw = float(tr2.BoundWidth)
        bh = float(tr2.BoundHeight)
        w = float(shape.Width)
        h = float(shape.Height)

        fits = (bw <= w and bh <= h)

        if single_line:
            try:
                line_count = tr2.Lines().Count
            except Exception:
                line_count = 1
            fits = fits and (line_count <= 1)

        return not fits
    except Exception:
        return False


def _rect(shape):
    return (
        float(shape.Left),
        float(shape.Top),
        float(shape.Left + shape.Width),
        float(shape.Top + shape.Height),
    )


def _intersects(a, b) -> bool:
    l1, t1, r1, b1 = a
    l2, t2, r2, b2 = b
    return not (r1 <= l2 or r2 <= l1 or b1 <= t2 or b2 <= t1)


def normalize_slide_type(slide_type: str, slide_spec: dict | None = None) -> str:
    slide_type = str(slide_type or "").strip()
    slide_spec = slide_spec or {}

    if slide_type == "flow":
        return "flow"

    if slide_type == "content_3":
        return resolve_content_3_template_key(slide_spec)

    if slide_type in SLIDE_REGISTRY:
        return slide_type

    return normalize_registry_type(slide_type)


def _get_slide_type_for_runtime_slide(slide):
    try:
        shape_names = set()
        for i in range(1, slide.Shapes.Count + 1):
            shp = slide.Shapes(i)
            try:
                shape_names.add(str(shp.Name).strip().lower())
            except Exception:
                pass

        for _, info in TEMPLATE_ROLE_MAP.items():
            template_names = set(info.get("shapes", {}).keys())
            if not template_names:
                continue
            if shape_names & template_names:
                detected_type = info.get("detected_type")
                if detected_type and detected_type != "unknown":
                    return detected_type
    except Exception:
        pass
    return None


def _get_shape_role(slide, shape):
    try:
        slide_type = _get_slide_type_for_runtime_slide(slide)
        if not slide_type:
            return "protected"

        for _, info in TEMPLATE_ROLE_MAP.items():
            if info.get("detected_type") != slide_type:
                continue

            shp_meta = info.get("shapes", {}).get(str(shape.Name).strip().lower())
            if shp_meta:
                return shp_meta.get("role", "protected")
    except Exception:
        pass

    return "protected"


def _should_ignore_overlap(slide, other_shape):
    role = _get_shape_role(slide, other_shape)
    return role == "background"


def _find_overlaps(slide, shape):
    overlaps = []
    target = _rect(shape)

    for i in range(1, slide.Shapes.Count + 1):
        other = slide.Shapes(i)

        try:
            if other.Name == shape.Name:
                continue
        except Exception:
            continue

        try:
            if _should_ignore_overlap(slide, other):
                continue
        except Exception:
            pass

        try:
            if _intersects(target, _rect(other)):
                overlaps.append(other)
        except Exception:
            pass

    return overlaps


# ============================================================
# Color helpers
# ============================================================

def rgb_to_tuple(rgb):
    r = rgb & 255
    g = (rgb >> 8) & 255
    b = (rgb >> 16) & 255
    return r, g, b


def brightness(rgb):
    r, g, b = rgb_to_tuple(rgb)
    return 0.299 * r + 0.587 * g + 0.114 * b


def detect_slide_text_color(slide):
    try:
        bg = slide.Background.Fill.ForeColor.RGB
        if brightness(bg) > 160:
            return 0
        else:
            return 16777215
    except Exception:
        return 0


# ============================================================
# Table helpers
# ============================================================

def try_add_columns(tbl, target_cols: int):
    safety = 50
    while tbl.Columns.Count < target_cols and safety > 0:
        safety -= 1
        before = tbl.Columns.Count
        try:
            tbl.Columns.Add()
        except Exception:
            return False
        if tbl.Columns.Count == before:
            return False
    return tbl.Columns.Count >= target_cols


def try_add_rows(tbl, target_rows: int):
    safety = 200
    while tbl.Rows.Count < target_rows and safety > 0:
        safety -= 1
        before = tbl.Rows.Count
        try:
            tbl.Rows.Add()
        except Exception:
            return False
        if tbl.Rows.Count == before:
            return False
    return tbl.Rows.Count >= target_rows


def ensure_table_size_safe(tbl, need_rows: int, need_cols: int):
    ok_cols = True
    ok_rows = True

    if need_cols > tbl.Columns.Count:
        ok_cols = try_add_columns(tbl, need_cols)

    if need_rows > tbl.Rows.Count:
        ok_rows = try_add_rows(tbl, need_rows)

    return ok_rows, ok_cols


def _set_wordwrap(shape):
    try:
        shape.TextFrame2.WordWrap = True
        return
    except Exception:
        pass
    try:
        shape.TextFrame.WordWrap = True
    except Exception:
        pass


def enable_wordwrap_for_table(tbl):
    for r in range(1, tbl.Rows.Count + 1):
        for c in range(1, tbl.Columns.Count + 1):
            _set_wordwrap(tbl.Cell(r, c).Shape)


def set_table_column_widths_by_text(tbl, col_text_lens, total_width):
    weights = [max(4, int(x)) for x in col_text_lens]
    s = sum(weights) or 1

    min_frac, max_frac = 0.10, 0.50
    fracs = [w / s for w in weights]
    fracs = [min(max(f, min_frac), max_frac) for f in fracs]
    fs = sum(fracs) or 1
    fracs = [f / fs for f in fracs]

    for i, f in enumerate(fracs, start=1):
        tbl.Columns(i).Width = total_width * f


def try_delete_extra_columns(tbl, keep_cols: int):
    """
    Delete columns from right to left, keeping only keep_cols columns.
    Return True if final column count <= keep_cols.
    """
    safety = 50
    while tbl.Columns.Count > keep_cols and safety > 0:
        safety -= 1
        before = tbl.Columns.Count
        try:
            tbl.Columns(tbl.Columns.Count).Delete()
        except Exception:
            return False

        if tbl.Columns.Count == before:
            return False

    return tbl.Columns.Count <= keep_cols


def try_delete_extra_rows(tbl, keep_rows: int):
    """
    Delete rows from bottom to top, keeping only keep_rows rows.
    Return True if final row count <= keep_rows.
    """
    safety = 200
    while tbl.Rows.Count > keep_rows and safety > 0:
        safety -= 1
        before = tbl.Rows.Count
        try:
            tbl.Rows(tbl.Rows.Count).Delete()
        except Exception:
            return False

        if tbl.Rows.Count == before:
            return False

    return tbl.Rows.Count <= keep_rows

def fill_table(slide, table_name: str, columns, rows):
    if not columns:
        print("[WARN] Table columns empty.")
        return False

    shp = shape_by_name(slide, table_name)
    if shp is None:
        print(f"[WARN] Table shape not found: {table_name}")
        return False

    try:
        if not shp.HasTable:
            print(f"[WARN] Shape is not a table: {table_name}")
            return False

        tbl = shp.Table
        need_rows = 1 + len(rows)   # header + data rows
        need_cols = len(columns)

        # --------------------------------------------------
        # Step 1: 先補足不足的 rows / cols
        # --------------------------------------------------
        ok_rows, ok_cols = ensure_table_size_safe(tbl, need_rows, need_cols)
        if not ok_rows or not ok_cols:
            print(f"[WARN] Table resize incomplete: rows_ok={ok_rows}, cols_ok={ok_cols}")

        # --------------------------------------------------
        # Step 2: 刪掉多餘 columns / rows
        # --------------------------------------------------
        if tbl.Columns.Count > need_cols:
            deleted_cols_ok = try_delete_extra_columns(tbl, need_cols)
            if not deleted_cols_ok:
                print(f"[WARN] Failed to delete extra columns. current={tbl.Columns.Count}, need={need_cols}")

        if tbl.Rows.Count > need_rows:
            deleted_rows_ok = try_delete_extra_rows(tbl, need_rows)
            if not deleted_rows_ok:
                print(f"[WARN] Failed to delete extra rows. current={tbl.Rows.Count}, need={need_rows}")

        # --------------------------------------------------
        # Step 3: 清空目前表格內容
        # --------------------------------------------------
        for r in range(1, tbl.Rows.Count + 1):
            for c in range(1, tbl.Columns.Count + 1):
                try:
                    tbl.Cell(r, c).Shape.TextFrame.TextRange.Text = ""
                except Exception:
                    pass

        # --------------------------------------------------
        # Step 4: 寫 header
        # --------------------------------------------------
        for c, text in enumerate(columns, start=1):
            if c <= tbl.Columns.Count:
                try:
                    tbl.Cell(1, c).Shape.TextFrame.TextRange.Text = str(text)
                except Exception:
                    pass

        # --------------------------------------------------
        # Step 5: 寫 data rows
        # --------------------------------------------------
        for r, row in enumerate(rows, start=2):
            if r > tbl.Rows.Count:
                break

            for c, text in enumerate(row[:need_cols], start=1):
                if c <= tbl.Columns.Count:
                    try:
                        tbl.Cell(r, c).Shape.TextFrame.TextRange.Text = str(text)
                    except Exception:
                        pass

        # --------------------------------------------------
        # Step 6: 啟用表格自動換行
        # --------------------------------------------------
        enable_wordwrap_for_table(tbl)

        # --------------------------------------------------
        # Step 7: 依文字長度重新分配欄寬
        # --------------------------------------------------
        col_text_lens = [len(str(col)) for col in columns]

        for row in rows:
            for i, cell in enumerate(row[:need_cols]):
                col_text_lens[i] = max(col_text_lens[i], len(str(cell)))

        try:
            total_width = shp.Width
            set_table_column_widths_by_text(tbl, col_text_lens, total_width)
        except Exception:
            pass

        return True

    except Exception as e:
        print(f"[WARN] Fill table failed ({table_name}): {e}")
        return False


# ============================================================
# SmartArt helpers
# ============================================================

def find_smartart_shape(slide, prefer_name=None):
    if prefer_name:
        shp = shape_by_name(slide, prefer_name)
        try:
            if shp is not None and getattr(shp, "HasSmartArt", False):
                return shp
        except Exception:
            pass

    for i in range(1, slide.Shapes.Count + 1):
        shp = slide.Shapes(i)
        try:
            if getattr(shp, "HasSmartArt", False):
                return shp
        except Exception:
            pass

    return None


def ensure_smartart_nodes(slide, desired_count, prefer_name=None):
    smart_shape = find_smartart_shape(slide, prefer_name=prefer_name)
    if smart_shape is None:
        print("[WARN] No SmartArt found to expand.")
        return None, 0

    try:
        nodes = smart_shape.SmartArt.AllNodes
    except Exception as e:
        print(f"[WARN] Cannot access SmartArt nodes: {e}")
        return smart_shape, 0

    if nodes.Count < desired_count:
        print(f"[WARN] SmartArt nodes ({nodes.Count}) < desired steps ({desired_count})")

    return smart_shape, nodes.Count

def reduce_smartart_nodes(slide, desired_count, prefer_name=None):
    smart_shape = find_smartart_shape(slide, prefer_name=prefer_name)
    if smart_shape is None:
        print("[WARN] No SmartArt found to reduce.")
        return None, 0

    try:
        nodes = smart_shape.SmartArt.AllNodes
    except Exception as e:
        print(f"[WARN] Cannot access SmartArt nodes: {e}")
        return smart_shape, 0

    safety = 50
    while nodes.Count > desired_count and safety > 0:
        safety -= 1
        deleted = False

        try:
            nodes(nodes.Count).Delete()
            deleted = True
        except Exception:
            pass

        if not deleted:
            print("[WARN] SmartArt layout refused to delete node.")
            break

        try:
            nodes = smart_shape.SmartArt.AllNodes
        except Exception:
            break

    return smart_shape, nodes.Count

def fill_smartart_steps(slide, steps, prefer_name=None):
    def _is_smartart_shape(shp):
        try:
            return bool(getattr(shp, "HasSmartArt", False))
        except Exception:
            return False

    def _try_write_node(node, text):
        try:
            node.TextFrame2.TextRange.Text = text
            return True
        except Exception:
            pass
        try:
            node.TextFrame.TextRange.Text = text
            return True
        except Exception:
            pass
        try:
            node.Shapes(1).TextFrame2.TextRange.Text = text
            return True
        except Exception:
            pass
        try:
            node.Shapes(1).TextFrame.TextRange.Text = text
            return True
        except Exception:
            pass
        return False

    smart_shape = None

    if prefer_name:
        try:
            shp = shape_by_name(slide, prefer_name)
            if shp is not None and _is_smartart_shape(shp):
                smart_shape = shp
        except Exception:
            smart_shape = None

    if smart_shape is None:
        for i in range(1, slide.Shapes.Count + 1):
            shp = slide.Shapes(i)
            if _is_smartart_shape(shp):
                smart_shape = shp
                break

    if smart_shape is None:
        print("[WARN] No SmartArt found on flow slide.")
        return False

    try:
        nodes = smart_shape.SmartArt.AllNodes
        node_count = nodes.Count
        n = min(len(steps), node_count)

        print(f"[DEBUG] SmartArt shape='{smart_shape.Name}', nodes={node_count}, fill={n}")

        for i in range(1, n + 1):
            node = nodes(i)
            text = str(steps[i - 1])
            ok = _try_write_node(node, text)
            if not ok:
                print(f"[WARN] Cannot write SmartArt node #{i}")

        return True

    except Exception as e:
        print(f"[WARN] Fill SmartArt failed: {e}")
        return False


# ============================================================
# Content cleanup helpers
# ============================================================

def _shape_has_textframe(shp):
    try:
        return bool(shp.HasTextFrame)
    except Exception:
        return False


def _is_text_placeholder_candidate(shp):
    try:
        if getattr(shp, "HasSmartArt", False):
            return False
    except Exception:
        pass

    try:
        if shp.HasTable:
            return False
    except Exception:
        pass

    return _shape_has_textframe(shp)


def delete_unupdated_content_shapes(slide, slide_type, keep_names):
    keep_names_lower = {str(x).strip().lower() for x in (keep_names or set())}

    for i in range(slide.Shapes.Count, 0, -1):
        shp = slide.Shapes(i)

        try:
            name = str(shp.Name).strip()
            name_lower = name.lower()
        except Exception:
            continue

        try:
            if getattr(shp, "HasSmartArt", False):
                continue
        except Exception:
            pass

        if name_lower in keep_names_lower:
            continue

        role = _get_shape_role(slide, shp)
        if role in {"background", "protected", "image", "table"}:
            continue

        try:
            if _is_text_placeholder_candidate(shp):
                shp.Delete()
        except Exception:
            try:
                if _shape_has_textframe(shp):
                    shp.TextFrame.TextRange.Text = ""
            except Exception:
                pass


# ============================================================
# Text writer
# ============================================================

def set_text(
    slide,
    shape_name: str,
    text: str,
    bold=None,
    auto_color=False,
    no_wrap: bool = False,
    single_line: bool = False,
):
    shp = shape_by_name(slide, shape_name)

    if shp is None:
        print(f"[WARN] Shape not found: {shape_name}")
        return False

    if not shp.HasTextFrame:
        return False

    clean_text = "" if text is None else str(text).strip()
    if clean_text == "":
        try:
            shp.Delete()
            return True
        except Exception:
            try:
                shp.TextFrame.TextRange.Text = ""
                return True
            except Exception:
                return False

    tr = shp.TextFrame.TextRange
    tr.Text = clean_text

    _set_wordwrap_and_autosize(shp, no_wrap=no_wrap)
    _clamp_shape_within_slide(slide, shp)
    _fit_text_to_shape(shp)

    # 只允許往下拉高文字框，不允許縮字
    if not single_line and not no_wrap:
        _expand_shape_height_to_fit_text(slide, shp, single_line=single_line)

    overlaps = _find_overlaps(slide, shp)

    # 若有重疊，再嘗試拉高一次；仍重疊就保留預設字體，只記警告
    if overlaps:
        if not single_line and not no_wrap:
            _expand_shape_height_to_fit_text(slide, shp, single_line=single_line)

        overlaps = _find_overlaps(slide, shp)
        if overlaps:
            print(f"[WARN] Text overlap remains for shape: {shape_name}")

    # 若無重疊但仍 overflow，也只記警告，不縮字
    elif _text_still_overflows_shape(shp, single_line=single_line):
        print(f"[WARN] Text still overflows shape: {shape_name}")

    _clamp_shape_within_slide(slide, shp)

    if bold is not None:
        try:
            tr.Font.Bold = bool(bold)
        except Exception:
            pass

    if auto_color:
        color = detect_slide_text_color(slide)
        try:
            tr.Font.Color.RGB = color
        except Exception:
            pass

    return True

# ============================================================
# Renderers
# ============================================================

@register_renderer("cover")
def render_cover(slide, slide_spec):
    keep_names = {"Topic", "speaker_name"}
    set_text(slide, "Topic", str(slide_spec.get("topic", "")), no_wrap=True, single_line=True)
    set_text(slide, "speaker_name", str(slide_spec.get("speaker", "")))
    return keep_names


@register_renderer("agenda")
def render_agenda(slide, slide_spec):
    keep_names = {"outline"}
    set_text(slide, "outline", slide_spec.get("title", "Agenda"), bold=True, no_wrap=True, single_line=True)

    items = slide_spec.get("items", [])
    for i in range(1, 6):
        name = f"agenda_{i}"
        keep_names.add(name)
        text = items[i - 1] if i <= len(items) else ""
        set_text(slide, name, text, bold=True, auto_color=True)

    return keep_names


@register_renderer("section")
def render_section(slide, slide_spec):
    keep_names = {"agenda_name"}

    section_name = str(
        slide_spec.get("name")
        or slide_spec.get("title")
        or "Section"
    ).strip()

    set_text(
        slide,
        "agenda_name",
        section_name,
        bold=True,
        auto_color=True
    )
    return keep_names


@register_renderer("table")
def render_table_slide(slide, slide_spec):
    keep_names = {"title", "sheet_1"}
    set_text(slide, "title", str(slide_spec.get("title", "")), no_wrap=True, single_line=True)
    fill_table(slide, "sheet_1", slide_spec.get("columns", []), slide_spec.get("rows", []))
    return keep_names


@register_renderer("flow")
def render_flow(slide, slide_spec):
    keep_names = {"title"}
    set_text(slide, "title", str(slide_spec.get("title", "")), no_wrap=True, single_line=True)
    steps = slide_spec.get("steps", [])

    prefer_name = _resolve_flow_prefer_name(slide, slide_spec)
    if prefer_name:
        keep_names.add(prefer_name)

    smart_shape, current_count = ensure_smartart_nodes(slide, len(steps), prefer_name=prefer_name)

    print(f"[DEBUG] Flow variant selected: {prefer_name}")
    print(f"[DEBUG] Flow template nodes before adjust: {current_count}")
    print(f"[DEBUG] Flow steps requested: {len(steps)}")

    if smart_shape is not None and current_count > len(steps):
        _, final_count = reduce_smartart_nodes(slide, len(steps), prefer_name=prefer_name)
        print(f"[DEBUG] Flow template nodes after reduce: {final_count}")

    fill_smartart_steps(slide, steps, prefer_name=prefer_name)
    return keep_names


@register_renderer("end")
def render_end(slide, slide_spec):
    return set()

# compatibility aliases
SLIDE_RENDERERS["content_text"] = render_content_image
SLIDE_RENDERERS["content_3extra_a"] = render_content_3extra
SLIDE_RENDERERS["content_3extra_b"] = render_content_3extra
SLIDE_RENDERERS["content_3extra_image"] = render_content_3extra


# ============================================================
# Slide rendering entrypoints
# ============================================================

def render_slide(slide, slide_spec):
    slide_type = slide_spec.get("type")
    normalized_type = normalize_slide_type(slide_type, slide_spec)

    fn = SLIDE_RENDERERS.get(normalized_type)
    if fn is None:
        print(f"[WARN] Unsupported slide type in render_slide: {slide_type}")
        
        return

    keep_names = fn(slide, slide_spec) or set()
    #delete_unupdated_content_shapes(slide, normalized_type, keep_names)


def get_template_slide_index(slide_type, src_pres, slide_spec=None):
    normalized_type = normalize_slide_type(slide_type, slide_spec)

    if normalized_type == "flow":
        variant = resolve_flow_template_key(slide_spec or {})
        if variant not in FLOW_TEMPLATE_INDEX:
            variant = choose_flow_variant(slide_spec or {})

        variant_idx = FLOW_TEMPLATE_INDEX.get(variant)
        if isinstance(variant_idx, int) and 1 <= variant_idx <= src_pres.Slides.Count:
            return variant_idx

        scanned = _find_template_slide_index_by_shape(src_pres, variant)
        if isinstance(scanned, int):
            return scanned

    cfg = SLIDE_REGISTRY[normalized_type]
    idx = cfg["template_slide_index"]
    if idx == "LAST":
        return src_pres.Slides.Count
    return idx

# ============================================================
# Public API
# ============================================================

def render_deck(template_pptx: str, deck_spec: dict, out_pptx: str):
    pythoncom.CoInitialize()

    app = None
    src = None
    dst = None

    try:
        app = win32.Dispatch("PowerPoint.Application")
        app.Visible = True

        work_pptx = os.path.abspath(out_pptx)
        shutil.copyfile(template_pptx, work_pptx)

        src = app.Presentations.Open(os.path.abspath(template_pptx), WithWindow=False)
        dst = app.Presentations.Open(work_pptx, WithWindow=False)

        delete_all_slides(dst)

        slides = list(deck_spec.get("slides", []) or [])

        # ensure cover
        if not slides or slides[0].get("type") != "cover":
            slides.insert(0, {
                "type": "cover",
                "topic": deck_spec.get("title", "Presentation Title"),
                "speaker": deck_spec.get("speaker", "")
            })

        # ensure end
        if slides[-1].get("type") != "end":
            slides.append({"type": "end"})

        for idx, slide_spec in enumerate(slides, start=1):
            slide_type = slide_spec.get("type")
            normalized_type = normalize_slide_type(slide_type, slide_spec)

            print(f"[DEBUG] Start render slide #{idx}: raw={slide_type}, normalized={normalized_type}")

            if normalized_type not in SLIDE_REGISTRY and normalized_type not in SLIDE_RENDERERS:
                print(f"[WARN] Skip unsupported slide type at #{idx}: {slide_type}")
                continue

            src_idx = get_template_slide_index(normalized_type, src, slide_spec)
            print(f"[DEBUG] Template slide index for #{idx}: {src_idx}")

            src_slide = src.Slides(src_idx)
            new_slide = duplicate_to_presentation(src_slide, dst)
            print(f"[DEBUG] Duplicated slide #{idx}")

            render_slide(new_slide, slide_spec)
            print(f"[DEBUG] Finished render slide #{idx}")

        dst.Save()
        print(f"[INFO] PPT generated: {work_pptx}")

    finally:
        if src is not None:
            src.Close()
        if dst is not None:
            dst.Close()
        if app is not None:
            app.Quit()

        pythoncom.CoUninitialize()


# ============================================================
# Optional CLI
# ============================================================

def _load_json(path: str):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def main():
    import sys

    if len(sys.argv) != 4:
        print("Usage: python ppt_renderer.py <template.pptx> <deck_spec.json> <output.pptx>")
        raise SystemExit(1)

    template_pptx = sys.argv[1]
    spec_json = sys.argv[2]
    out_pptx = sys.argv[3]

    deck_spec = _load_json(spec_json)
    render_deck(template_pptx, deck_spec, out_pptx)


if __name__ == "__main__":
    main()
