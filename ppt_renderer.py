import os
import shutil
import win32com.client as win32
from slide_registry import SLIDE_REGISTRY
import pythoncom

SLIDE_RENDERERS = {}

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

def set_text(slide, shape_name: str, text: str, bold=None, auto_color=False):
    
    shp = shape_by_name(slide, shape_name)

    if shp is None:
        print(f"[WARN] Shape not found: {shape_name}")
        return False

    if not shp.HasTextFrame:
        return False

    tr = shp.TextFrame.TextRange
    tr.Text = text

    if bold is not None:
        try:
            tr.Font.Bold = bool(bold)
        except:
            pass

    if auto_color:
        color = detect_slide_text_color(slide)
        try:
            tr.Font.Color.RGB = color
        except:
            pass

    return True

def clear_textboxes_except(slide, keep_names: set):
    for i in range(1, slide.Shapes.Count + 1):
        shp = slide.Shapes(i)

        # Always keep SmartArt (avoid accidental wiping)
        try:
            if getattr(shp, "HasSmartArt", False):
                continue
        except Exception:
            pass

        if shp.Name in keep_names:
            continue

        try:
            if shp.HasTextFrame:
                shp.TextFrame.TextRange.Text = ""
        except Exception:
            pass


def duplicate_to_presentation(src_slide, dst_pres):
    """
    Copy a slide into another presentation, keeping design as much as possible.
    """
    src_slide.Copy()
    dst_pres.Slides.Paste(dst_pres.Slides.Count + 1)
    return dst_pres.Slides(dst_pres.Slides.Count)

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
    set_text(slide, "Topic", str(slide_spec.get("topic", "")))
    set_text(slide, "speaker_name", str(slide_spec.get("speaker", "")))


@register_renderer("agenda")
def render_agenda(slide, slide_spec):

    set_text(slide, "outline", slide_spec.get("title", "Agenda"), bold=True)

    items = slide_spec.get("items", [])

    for i in range(1, 6):

        text = items[i-1] if i <= len(items) else ""

        set_text(
            slide,
            f"agenda_{i}",
            text,
            bold=True,
            auto_color=True
        )
        
@register_renderer("section")
def render_section(slide, slide_spec):

    set_text(
        slide,
        "agenda_name",
        slide_spec.get("name", ""),
        bold=True,
        auto_color=True
    )

@register_renderer("content_2")
def render_content_2(slide, slide_spec):
    if shape_by_name(slide, "title"):
        set_text(slide, "title", str(slide_spec.get("title", "")))

    cards = slide_spec.get("cards", [])

    for i in range(1, 3):
        if i <= len(cards):
            card = cards[i - 1]
            if shape_by_name(slide, f"item_{i}"):
                set_text(slide, f"item_{i}", str(card.get("item", "")))
            if shape_by_name(slide, f"content_{i}"):
                set_text(slide, f"content_{i}", str(card.get("content", "")))
        else:
            if shape_by_name(slide, f"item_{i}"):
                set_text(slide, f"item_{i}", "")
            if shape_by_name(slide, f"content_{i}"):
                set_text(slide, f"content_{i}", "")
                
SLIDE_RENDERERS["content_2_a"] = render_content_2
SLIDE_RENDERERS["content_2_b"] = render_content_2
SLIDE_RENDERERS["content_2_c"] = render_content_2

@register_renderer("content_4")
def render_content_4(slide, slide_spec):
    if shape_by_name(slide, "title"):
        set_text(slide, "title", str(slide_spec.get("title", "")))

    cards = slide_spec.get("cards", [])

    for i in range(1, 5):
        if i <= len(cards):
            card = cards[i - 1]
            if shape_by_name(slide, f"item_{i}"):
                set_text(slide, f"item_{i}", str(card.get("item", "")))
            if shape_by_name(slide, f"content_{i}"):
                set_text(slide, f"content_{i}", str(card.get("content", "")))
        else:
            if shape_by_name(slide, f"item_{i}"):
                set_text(slide, f"item_{i}", "")
            if shape_by_name(slide, f"content_{i}"):
                set_text(slide, f"content_{i}", "")
                
SLIDE_RENDERERS["content_4_a"] = render_content_4
SLIDE_RENDERERS["content_4_b"] = render_content_4

@register_renderer("content_3extra")
def render_content_3extra(slide, slide_spec):
    set_text(slide, "title", str(slide_spec.get("title", "")))
    cards = slide_spec.get("cards", [])

    for i in range(1, 4):
        if i <= len(cards):
            card = cards[i - 1]
            set_text(slide, f"item_{i}", str(card.get("item", "")))
            set_text(slide, f"content_{i}", str(card.get("content", "")))
        else:
            set_text(slide, f"item_{i}", "")
            set_text(slide, f"content_{i}", "")


@register_renderer("table")
def render_table_slide(slide, slide_spec):
    set_text(slide, "title", str(slide_spec.get("title", "")))
    fill_table(
        slide,
        "sheet_1",
        slide_spec.get("columns", []),
        slide_spec.get("rows", [])
    )


@register_renderer("flow")
def render_flow(slide, slide_spec):
    set_text(slide, "title", str(slide_spec.get("title", "")))
    steps = slide_spec.get("steps", [])

    smart_shape, current_count = ensure_smartart_nodes(
        slide,
        len(steps),
        prefer_name="flow_chart_1"
    )

    print(f"[DEBUG] Flow template nodes before adjust: {current_count}")
    print(f"[DEBUG] Flow steps requested: {len(steps)}")

    if smart_shape is not None and current_count > len(steps):
        _, final_count = reduce_smartart_nodes(
            slide,
            len(steps),
            prefer_name="flow_chart_1"
        )
        print(f"[DEBUG] Flow template nodes after reduce: {final_count}")

    fill_smartart_steps(slide, steps, prefer_name="flow_chart_1")

@register_renderer("content_image")
def render_content_image(slide, slide_spec):
    if shape_by_name(slide, "title"):
        set_text(slide, "title", str(slide_spec.get("title", "")))
    if shape_by_name(slide, "content"):
        set_text(slide, "content", str(slide_spec.get("content", "")))

@register_renderer("end")
def render_end(slide, slide_spec):
    pass

def render_slide(slide, slide_spec):
    t = slide_spec.get("type")
    fn = SLIDE_RENDERERS.get(t)

    if fn is None:
        print(f"[WARN] Unsupported slide type in render_slide: {t}")
        return

    fn(slide, slide_spec)

def get_template_slide_index(slide_type, src_pres):
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

            src_idx = get_template_slide_index(slide_type, src)
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