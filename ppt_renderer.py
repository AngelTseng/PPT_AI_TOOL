import os
import time
import json
import shutil
import win32com.client as win32
from slide_registry import FLOW_TEMPLATE_INDEX, SLIDE_REGISTRY
import pythoncom

SLIDE_RENDERERS = {}

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_MAP_PATH = os.path.join(BASE_DIR, "template_map.json")

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

MsoAutoSizeTextToFitShape = 2
MsoTrue = -1
MsoFalse = 0


def _fit_text_to_shape(shape):
    """Best-effort: keep text inside the textbox instead of overflowing."""
    try:
        shape.TextFrame2.WordWrap = MsoTrue
        shape.TextFrame2.AutoSize = MsoAutoSizeTextToFitShape
        return
    except Exception:
        pass

    try:
        shape.TextFrame.WordWrap = MsoTrue
        shape.TextFrame.AutoSize = 1
    except Exception:
        pass


def _clamp_shape_within_slide(slide, shape):
    """If a shape is moved outside the page bounds, clamp it back."""
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




def choose_flow_variant(slide_spec: dict) -> str:
    title = str(slide_spec.get("title", "")).lower()
    steps = [str(s).lower() for s in slide_spec.get("steps", [])]
    text = " ".join([title] + steps)

    loop_keywords = [
        "iteration", "iterate", "feedback", "cycle", "optimize", "improve", "review",
        "迭代", "循環", "回饋", "優化", "改善", "檢討"
    ]
    linear_keywords = [
        "first", "then", "next", "finally", "phase", "stage", "process",
        "首先", "接著", "然後", "最後", "階段", "流程"
    ]
    framework_keywords = [
        "module", "stream", "framework", "layer", "pillar", "architecture",
        "模組", "工作流", "架構", "層", "面向", "支柱"
    ]

    loop_score = sum(1 for kw in loop_keywords if kw in text)
    linear_score = sum(1 for kw in linear_keywords if kw in text)
    framework_score = sum(1 for kw in framework_keywords if kw in text)

    if loop_score >= max(linear_score, framework_score) and loop_score > 0:
        return "flow_chart_2"
    if framework_score >= max(linear_score, loop_score) and framework_score > 0:
        return "flow_chart_3"
    return "flow_chart_1"


def _resolve_flow_prefer_name(slide, slide_spec: dict) -> str | None:
    candidate = choose_flow_variant(slide_spec)

    for name in (candidate, "flow_chart_1", "flow_chart_2", "flow_chart_3"):
        shp = shape_by_name(slide, name)
        try:
            if shp is not None and getattr(shp, "HasSmartArt", False):
                return name
        except Exception:
            pass

    return None


def _find_template_slide_index_by_shape(src_pres, shape_name: str) -> int | None:
    for i in range(1, src_pres.Slides.Count + 1):
        slide = src_pres.Slides(i)
        for j in range(1, slide.Shapes.Count + 1):
            shp = slide.Shapes(j)
            try:
                if str(shp.Name).strip().lower() == shape_name.lower():
                    return i
            except Exception:
                continue
    return None



def _set_wordwrap_and_autosize(shape, no_wrap: bool = False):
    wrap = MsoFalse if no_wrap else MsoTrue
    try:
        shape.TextFrame2.WordWrap = wrap
        shape.TextFrame2.AutoSize = MsoAutoSizeTextToFitShape
    except Exception:
        pass
    try:
        shape.TextFrame.WordWrap = bool(not no_wrap)
    except Exception:
        pass

def _shrink_text_to_fit_shape(
    shape,
    min_font_size: float = 10.0,
    single_line: bool = False,
):
    """Reduce font size until text fits inside shape bounds.
    If single_line=True, keep shrinking until text stays on one line.
    """
    try:
        tr2 = shape.TextFrame2.TextRange
        if tr2.Length <= 0:
            return

        try:
            current_size = float(tr2.Font.Size)
        except Exception:
            current_size = 18.0

        safety = 60
        while safety > 0:
            safety -= 1

            try:
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

            current_size -= 0.5
            if current_size < min_font_size:
                break

            try:
                tr2.Font.Size = current_size
            except Exception:
                break
    except Exception:
        pass

def register_renderer(name):
    def wrapper(func):
        SLIDE_RENDERERS[name] = func
        return func
    return wrapper

def set_table_column_widths_by_text(tbl, col_text_lens, total_width):
    weights = [max(4, int(x)) for x in col_text_lens]
    s = sum(weights) or 1

    min_frac, max_frac = 0.10, 0.50
    fracs = [w / s for w in weights]
    fracs = [min(max(f, min_frac), max_frac) for f in fracs]
    fs = sum(fracs) or 1
    fracs = [f / fs for f in fracs]

    for i, f in enumerate(fracs, start=1):  # COM is 1-based
        tbl.Columns(i).Width = total_width * f

def try_add_columns(tbl, target_cols: int):
    """
    Try to add columns until reaching target_cols.
    Returns True if reached, False otherwise.
    """
    # 有些表格 Columns.Add() 會 throw 或無效，所以每次都檢查 Count
    safety = 50
    while tbl.Columns.Count < target_cols and safety > 0:
        safety -= 1
        before = tbl.Columns.Count
        try:
            tbl.Columns.Add()
        except Exception:
            return False
        # 驗證是否真的增加
        if tbl.Columns.Count == before:
            return False
    return tbl.Columns.Count >= target_cols

def try_add_rows(tbl, target_rows: int):
    """
    Try to add rows until reaching target_rows.
    Returns True if reached, False otherwise.
    """
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
    """
    Best-effort resize table to (need_rows, need_cols).
    Returns (ok_rows, ok_cols).
    """
    ok_cols = True
    ok_rows = True

    if need_cols > tbl.Columns.Count:
        ok_cols = try_add_columns(tbl, need_cols)

    if need_rows > tbl.Rows.Count:
        ok_rows = try_add_rows(tbl, need_rows)

    return ok_rows, ok_cols


def _set_wordwrap(shape):
    # Prefer TextFrame2; fallback to TextFrame
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
  
def find_smartart_shape(slide, prefer_name=None):
    # 先用 prefer_name 鎖定
    if prefer_name:
        shp = shape_by_name(slide, prefer_name)
        try:
            if shp is not None and getattr(shp, "HasSmartArt", False):
                return shp
        except Exception:
            pass

    # 找第一個 SmartArt
    for i in range(1, slide.Shapes.Count + 1):
        shp = slide.Shapes(i)
        try:
            if getattr(shp, "HasSmartArt", False):
                return shp
        except Exception:
            pass

    return None


def ensure_smartart_nodes(slide, desired_count, prefer_name=None):
    """
    Try to increase SmartArt node count to desired_count (same slide).
    Returns (smart_shape, final_count).
    """
    smart_shape = find_smartart_shape(slide, prefer_name=prefer_name)
    if smart_shape is None:
        print("[WARN] No SmartArt found to expand.")
        return None, 0

    # 有些版本需要 select 才可編輯
    try:
        slide.Select()
    except Exception:
        pass
    try:
        smart_shape.Select()
    except Exception:
        pass

    try:
        nodes = smart_shape.SmartArt.AllNodes
    except Exception as e:
        print(f"[WARN] Cannot access SmartArt nodes: {e}")
        return smart_shape, 0
    
    if nodes.Count < desired_count:
        print(f"[WARN] SmartArt nodes ({nodes.Count}) < desired steps ({desired_count})")

    return smart_shape, nodes.Count

def reduce_smartart_nodes(slide, desired_count, prefer_name=None):
    """
    Reduce SmartArt nodes to desired_count (if possible).
    Returns (smart_shape, final_count).
    """
    smart_shape = find_smartart_shape(slide, prefer_name=prefer_name)
    if smart_shape is None:
        print("[WARN] No SmartArt found to reduce.")
        return None, 0

    # 有些版本需要 select
    try:
        slide.Select()
    except Exception:
        pass
    try:
        smart_shape.Select()
    except Exception:
        pass

    try:
        nodes = smart_shape.SmartArt.AllNodes
    except Exception as e:
        print(f"[WARN] Cannot access SmartArt nodes: {e}")
        return smart_shape, 0

    safety = 50
    while nodes.Count > desired_count and safety > 0:
        safety -= 1
        deleted = False

        # 嘗試刪最後一個 node（最安全）
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
    """
    Fill SmartArt node texts on a slide.
    - prefer_name: optional shape name to target a specific SmartArt first.
    Returns True if it attempted to fill; False if no SmartArt found.
    """

    def _is_smartart_shape(shp):
        try:
            return bool(getattr(shp, "HasSmartArt", False))
        except Exception:
            return False

    def _try_write_node(node, text):
        """
        Try multiple write paths. Return True if any succeeds.
        """
        # 1) node.TextFrame2
        try:
            node.TextFrame2.TextRange.Text = text
            return True
        except Exception:
            pass

        # 2) node.TextFrame
        try:
            node.TextFrame.TextRange.Text = text
            return True
        except Exception:
            pass

        # 3) node.Shapes(1).TextFrame2
        try:
            node.Shapes(1).TextFrame2.TextRange.Text = text
            return True
        except Exception:
            pass

        # 4) node.Shapes(1).TextFrame
        try:
            node.Shapes(1).TextFrame.TextRange.Text = text
            return True
        except Exception:
            pass

        return False

    smart_shape = None

    # 1) Prefer targeting by name, if provided
    if prefer_name:
        try:
            shp = shape_by_name(slide, prefer_name)  # uses your existing helper
            if shp is not None and _is_smartart_shape(shp):
                smart_shape = shp
        except Exception:
            smart_shape = None

    # 2) Otherwise pick the first SmartArt on the slide
    if smart_shape is None:
        for i in range(1, slide.Shapes.Count + 1):
            shp = slide.Shapes(i)
            if _is_smartart_shape(shp):
                smart_shape = shp
                break

    if smart_shape is None:
        print("[WARN] No SmartArt found on flow slide.")
        return False

    # Some environments need selection to allow editing SmartArt
    try:
        slide.Select()
    except Exception:
        pass
    try:
        smart_shape.Select()
    except Exception:
        pass

    # Fill nodes
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
    
def shape_by_name(slide, name: str):
    for i in range(1, slide.Shapes.Count + 1):
        shp = slide.Shapes(i)
        if shp.Name == name:
            return shp
    return None

def set_text(slide, shape_name: str, text: str, bold=None, auto_color=False, no_wrap: bool = False, single_line: bool = False,):
    
    shp = shape_by_name(slide, shape_name)

    if shp is None:
        print(f"[WARN] Shape not found: {shape_name}")
        return False

    if not shp.HasTextFrame:
        return False

    clean_text = "" if text is None else str(text).strip()
    if clean_text == "":
        # Requirement: remove unfilled textboxes directly.
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
    _shrink_text_to_fit_shape(shp, single_line = single_line)
    _clamp_shape_within_slide(slide, shp)

    # Keep text within textbox and avoid visual overflow when content is longer.
    _fit_text_to_shape(shp)
    _clamp_shape_within_slide(slide, shp)
    
    _fit_text_to_shape(shp)
    _resolve_overlap_or_fit(slide, shp)
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

CONTENT_CLEANUP_NAMES = {
    "content_image": {"title", "title_content", "item", "content"},
    "content_text": {"title", "title_content", "item", "content"},

    "content_2_a": {"title", "title_content_1", "title_content_2", "item_1", "item_2", "content_1", "content_2"},
    "content_2_b": {"title_content_1", "title_content_2", "item_1", "item_2", "content_1", "content_2"},
    "content_2_c": {"item_1", "item_2", "content_1", "content_2"},

    "content_3extra": {"title", "item_1", "item_2", "item_3", "content_1", "content_2", "content_3"},
    "content_3extra_image": {"title", "item_1", "item_2", "item_3", "content_1", "content_2", "content_3"},

    "content_4_a": {"title", "title_content_1", "title_content_2", "title_content_3", "title_content_4", "item_1", "item_2", "item_3", "item_4", "content_1", "content_2", "content_3", "content_4"},
    "content_4_b": {"title_content_1", "title_content_2", "title_content_3", "title_content_4", "item_1", "item_2", "item_3", "item_4", "content_1", "content_2", "content_3", "content_4"},
}

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

def _get_slide_type_for_runtime_slide(slide):
    try:
        shape_names = set()
        for i in range(1, slide.Shapes.Count + 1):
            shp = slide.Shapes(i)
            try:
                shape_names.add(str(shp.Name).strip().lower())
            except Exception:
                pass

        for slide_index, info in TEMPLATE_ROLE_MAP.items():
            template_names = set(info.get("shapes", {}).keys())
            if not template_names:
                continue

            # runtime slide 與 template slide 的命名 shape 有交集就視為同版型
            if shape_names & template_names:
                detected_type = info.get("detected_type")
                if detected_type:
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
            # 背景物件直接忽略，不參與 overlap
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

def _resolve_overlap_or_fit(slide, shape, max_expand: float = 80.0):
    # 先左右擴張
    step = 6.0
    expanded = 0.0

    while expanded < max_expand:
        overlaps = _find_overlaps(slide, shape)
        if not overlaps:
            return

        moved = False
        try:
            if float(shape.Left) > step:
                original_left = float(shape.Left)
                original_width = float(shape.Width)

                shape.Left = original_left - step / 2
                shape.Width = original_width + step
                moved = True

                # 若擴張後反而更糟，就還原
                if _find_overlaps(slide, shape):
                    shape.Left = original_left
                    shape.Width = original_width
                    moved = False
        except Exception:
            pass

        if not moved:
            break

        expanded += step

    # 再縮字
    _shrink_text_to_fit_shape(shape)

def delete_unupdated_content_shapes(slide, slide_type: str, keep_names: set[str]):
    deletable_names = CONTENT_CLEANUP_NAMES.get(slide_type, set())
    if not deletable_names:
        return

    for i in range(slide.Shapes.Count, 0, -1):
        shp = slide.Shapes(i)

        try:
            name = str(shp.Name)
        except Exception:
            continue

        if name not in deletable_names:
            continue

        if name in keep_names:
            continue

        try:
            if shp.HasTable:
                continue
        except Exception:
            pass

        try:
            if getattr(shp, "HasSmartArt", False):
                continue
        except Exception:
            pass

        try:
            if shp.HasTextFrame:
                shp.Delete()
        except Exception:
            try:
                if shp.HasTextFrame:
                    shp.TextFrame.TextRange.Text = ""
            except Exception:
                pass

def duplicate_to_presentation(src_slide, dst_pres, max_retries: int = 10):
    """
    Robust slide copy/paste across presentations.
    """
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
                # SlideRange -> first pasted slide
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
    # 從後面刪，最安全
    for i in range(pres.Slides.Count, 0, -1):
        pres.Slides(i).Delete()

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
        need_rows = 1 + len(rows)   # header + body
        need_cols = len(columns)

        ok_rows, ok_cols = ensure_table_size_safe(tbl, need_rows, need_cols)
        if not ok_rows or not ok_cols:
            print(f"[WARN] Table resize incomplete: rows_ok={ok_rows}, cols_ok={ok_cols}")

        # header
        for c, text in enumerate(columns, start=1):
            tbl.Cell(1, c).Shape.TextFrame.TextRange.Text = str(text)

        # body
        for r, row in enumerate(rows, start=2):
            for c, text in enumerate(row[:need_cols], start=1):
                tbl.Cell(r, c).Shape.TextFrame.TextRange.Text = str(text)

        enable_wordwrap_for_table(tbl)

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
            return 0        # black
        else:
            return 16777215 # white
    except:
        return 0

@register_renderer("cover")
def render_cover(slide, slide_spec):
    keep_names = {"Topic", "speaker_name"}

    set_text(slide, "Topic", str(slide_spec.get("topic", "")), no_wrap=True)
    set_text(slide, "speaker_name", str(slide_spec.get("speaker", "")))

    return keep_names

@register_renderer("agenda")
def render_agenda(slide, slide_spec):
    keep_names = {"outline"}

    set_text(slide, "outline", slide_spec.get("title", "Agenda"), bold=True, no_wrap=True)

    items = slide_spec.get("items", [])

    for i in range(1, 6):
        name = f"agenda_{i}"
        keep_names.add(name)

        text = items[i-1] if i <= len(items) else ""

        set_text(
            slide,
            name,
            text,
            bold=True,
            auto_color=True
        )

    return keep_names

@register_renderer("section")
def render_section(slide, slide_spec):
    keep_names = {"agenda_name"}

    set_text(
        slide,
        "agenda_name",
        slide_spec.get("name", ""),
        bold=True,
        auto_color=True
    )

    return keep_names

@register_renderer("content_2_a")
def render_content_2_a(slide, slide_spec):
    keep_names = set()

    if shape_by_name(slide, "title"):
        keep_names.add("title")
        set_text(slide, "title", str(slide_spec.get("title", "")), no_wrap=True, single_line=True)

    cards = slide_spec.get("cards", [])

    for i in range(1, 3):
        title_content_name = f"title_content_{i}"
        item_name = f"item_{i}"
        content_name = f"content_{i}"

        card = cards[i - 1] if i <= len(cards) else {}
        item_text = str(card.get("item", ""))
        content_text = str(card.get("content", ""))

        if shape_by_name(slide, title_content_name):
            keep_names.add(title_content_name)
            set_text(slide, title_content_name, item_text)
        elif shape_by_name(slide, item_name):
            keep_names.add(item_name)
            set_text(slide, item_name, item_text)

        if shape_by_name(slide, content_name):
            keep_names.add(content_name)
            set_text(slide, content_name, content_text)

    return keep_names

@register_renderer("content_2_b")
def render_content_2_b(slide, slide_spec):
    keep_names = set()
    cards = slide_spec.get("cards", [])

    for i in range(1, 3):
        title_content_name = f"title_content_{i}"
        item_name = f"item_{i}"
        content_name = f"content_{i}"

        card = cards[i - 1] if i <= len(cards) else {}
        item_text = str(card.get("item", ""))
        content_text = str(card.get("content", ""))

        if shape_by_name(slide, title_content_name):
            keep_names.add(title_content_name)
            set_text(slide, title_content_name, item_text)
        elif shape_by_name(slide, item_name):
            keep_names.add(item_name)
            set_text(slide, item_name, item_text)

        if shape_by_name(slide, content_name):
            keep_names.add(content_name)
            set_text(slide, content_name, content_text)

    return keep_names
             
@register_renderer("content_2_c")
def render_content_2_c(slide, slide_spec):
    keep_names = set()
    cards = slide_spec.get("cards", [])

    for i in range(1, 3):
        item_name = f"item_{i}"
        content_name = f"content_{i}"

        card = cards[i - 1] if i <= len(cards) else {}

        if shape_by_name(slide, item_name):
            keep_names.add(item_name)
            set_text(slide, item_name, str(card.get("item", "")))

        if shape_by_name(slide, content_name):
            keep_names.add(content_name)
            set_text(slide, content_name, str(card.get("content", "")))

    return keep_names
             
@register_renderer("content_4_a")
def render_content_4_a(slide, slide_spec):
    keep_names = set()

    if shape_by_name(slide, "title"):
        keep_names.add("title")
        set_text(slide, "title", str(slide_spec.get("title", "")), no_wrap=True, single_line=True)

    cards = slide_spec.get("cards", [])

    for i in range(1, 5):
        item_name = f"item_{i}"
        content_name = f"content_{i}"
        title_content_name = f"title_content_{i}"
        card = cards[i - 1] if i <= len(cards) else {}

        if shape_by_name(slide, title_content_name):
            keep_names.add(title_content_name)
            set_text(slide, title_content_name, str(card.get("item", "")))
        elif shape_by_name(slide, item_name):
            keep_names.add(item_name)
            set_text(slide, item_name, str(card.get("item", "")))

        if shape_by_name(slide, content_name):
            keep_names.add(content_name)
            set_text(slide, content_name, str(card.get("content", "")))

    return keep_names
                
@register_renderer("content_4_b")
def render_content_4_b(slide, slide_spec):
    keep_names = set()
    cards = slide_spec.get("cards", [])

    for i in range(1, 5):
        item_name = f"item_{i}"
        content_name = f"content_{i}"
        title_content_name = f"title_content_{i}"
        card = cards[i - 1] if i <= len(cards) else {}

        if shape_by_name(slide, title_content_name):
            keep_names.add(title_content_name)
            set_text(slide, title_content_name, str(card.get("item", "")))
        elif shape_by_name(slide, item_name):
            keep_names.add(item_name)
            set_text(slide, item_name, str(card.get("item", "")))

        if shape_by_name(slide, content_name):
            keep_names.add(content_name)
            set_text(slide, content_name, str(card.get("content", "")))

    return keep_names

@register_renderer("content_3extra")
def render_content_3extra(slide, slide_spec):
    keep_names = {"title"}

    set_text(slide, "title", str(slide_spec.get("title", "")), no_wrap=True, single_line=True)
    cards = slide_spec.get("cards", [])

    for i in range(1, 4):
        item_name = f"item_{i}"
        content_name = f"content_{i}"
        keep_names.add(item_name)
        keep_names.add(content_name)

        if i <= len(cards):
            card = cards[i - 1]
            set_text(slide, item_name, str(card.get("item", "")))
            set_text(slide, content_name, str(card.get("content", "")))
        else:
            set_text(slide, item_name, "")
            set_text(slide, content_name, "")

    return keep_names


@register_renderer("table")
def render_table_slide(slide, slide_spec):
    keep_names = {"title", "sheet_1"}

    set_text(slide, "title", str(slide_spec.get("title", "")), no_wrap=True, single_line=True)
    fill_table(
        slide,
        "sheet_1",
        slide_spec.get("columns", []),
        slide_spec.get("rows", [])
    )

    return keep_names


@register_renderer("flow")
def render_flow(slide, slide_spec):
    keep_names = {"title"}

    set_text(slide, "title", str(slide_spec.get("title", "")), no_wrap=True, single_line=True)
    steps = slide_spec.get("steps", [])

    prefer_name = _resolve_flow_prefer_name(slide, slide_spec)
    if prefer_name:
        keep_names.add(prefer_name)

    smart_shape, current_count = ensure_smartart_nodes(
        slide,
        len(steps),
        prefer_name=prefer_name
    )

    print(f"[DEBUG] Flow variant selected: {prefer_name}")
    print(f"[DEBUG] Flow template nodes before adjust: {current_count}")
    print(f"[DEBUG] Flow steps requested: {len(steps)}")

    if smart_shape is not None and current_count > len(steps):
        _, final_count = reduce_smartart_nodes(
            slide,
            len(steps),
            prefer_name=prefer_name
        )
        print(f"[DEBUG] Flow template nodes after reduce: {final_count}")

    fill_smartart_steps(slide, steps, prefer_name=prefer_name)

    return keep_names

@register_renderer("content_image")
def render_content_image(slide, slide_spec):
    keep_names = set()

    if shape_by_name(slide, "title"):
        keep_names.add("title")
        set_text(slide, "title", str(slide_spec.get("title", "")), no_wrap=True, single_line=True)

    if shape_by_name(slide, "content"):
        keep_names.add("content")
        set_text(slide, "content", str(slide_spec.get("content", "")))

    return keep_names

SLIDE_RENDERERS["content_text"] = render_content_image

@register_renderer("end")
def render_end(slide, slide_spec):
    return set()

SLIDE_RENDERERS["content_3extra_image"] = render_content_3extra

def render_slide(slide, slide_spec):
    t = slide_spec.get("type")
    fn = SLIDE_RENDERERS.get(t)

    if fn is None:
        print(f"[WARN] Unsupported slide type in render_slide: {t}")
        return

    keep_names = fn(slide, slide_spec) or set()
    delete_unupdated_content_shapes(slide, t, keep_names)

def get_template_slide_index(slide_type, src_pres, slide_spec=None):
    if slide_type == "flow":
        variant = choose_flow_variant(slide_spec or {})
        variant_idx = FLOW_TEMPLATE_INDEX.get(variant)
        if isinstance(variant_idx, int) and 1 <= variant_idx <= src_pres.Slides.Count:
            return variant_idx

        # fallback: scan template by SmartArt shape name if mapping is stale
        scanned = _find_template_slide_index_by_shape(src_pres, variant)
        if isinstance(scanned, int):
            return scanned

    cfg = SLIDE_REGISTRY[slide_type]
    idx = cfg["template_slide_index"]
    if idx == "LAST":
        return src_pres.Slides.Count
    return idx

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

        slides = deck_spec.get("slides", [])

        for idx, slide_spec in enumerate(slides, start=1):
            slide_type = slide_spec.get("type")
            if slide_type not in SLIDE_REGISTRY:
                print(f"[WARN] Skip unsupported slide type: {slide_type}")
                continue

            src_idx = get_template_slide_index(slide_type, src, slide_spec)
            src_slide = src.Slides(src_idx)

            new_slide = duplicate_to_presentation(src_slide, dst)
            render_slide(new_slide, slide_spec)

        dst.Save()

    finally:
        if src is not None:
            src.Close()
        if dst is not None:
            dst.Close()
        if app is not None:
            app.Quit()

        pythoncom.CoUninitialize()