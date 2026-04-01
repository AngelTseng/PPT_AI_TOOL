# ============================================================
# Renderer Helper (Shared Utilities)
# ============================================================
import os
MsoTrue = -1
MsoFalse = 0


# ============================================================
# Shape Cache
# ============================================================

def build_shape_cache(slide):
    cache = {}
    for i in range(1, slide.Shapes.Count + 1):
        try:
            shp = slide.Shapes(i)
            cache[str(shp.Name).strip().lower()] = shp
        except Exception:
            pass
    return cache


def shape_by_name(slide, name: str):
    if not name:
        return None

    key = str(name).strip().lower()

    try:
        cache = getattr(slide, "_shape_cache", None)
    except Exception:
        cache = None

    if isinstance(cache, dict):
        return cache.get(key)

    for i in range(1, slide.Shapes.Count + 1):
        shp = slide.Shapes(i)
        try:
            if str(shp.Name).strip().lower() == key:
                return shp
        except Exception:
            pass

    return None


# ============================================================
# Text Layout Helpers
# ============================================================

def _fit_text_to_shape(shape):
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
    try:
        sw = float(slide.Parent.PageSetup.SlideWidth)
        sh = float(slide.Parent.PageSetup.SlideHeight)

        left = float(shape.Left)
        top = float(shape.Top)
        width = float(shape.Width)
        height = float(shape.Height)

        if left < 0:
            shape.Left = 0
        if top < 0:
            shape.Top = 0
        if left + width > sw:
            shape.Left = max(0, sw - width)
        if top + height > sh:
            shape.Top = max(0, sh - height)
    except Exception:
        pass


def _expand_shape_height_to_fit_text(
    slide,
    shape,
    max_extra_height: float = 220.0,
    step: float = 12.0,
    single_line: bool = False,
):
    try:
        original_height = float(shape.Height)
    except Exception:
        return False

    max_height = original_height + max_extra_height

    for _ in range(30):
        try:
            tr2 = shape.TextFrame2.TextRange
            bw = float(tr2.BoundWidth)
            bh = float(tr2.BoundHeight)

            if bh <= float(shape.Height) + 1:
                return True
        except Exception:
            break

        try:
            new_height = min(float(shape.Height) + step, max_height)
            if new_height <= float(shape.Height) + 0.1:
                break
            shape.Height = new_height
            _clamp_shape_within_slide(slide, shape)
        except Exception:
            break

    return False


def _text_still_overflows_shape(shape, single_line: bool = False) -> bool:
    try:
        tr2 = shape.TextFrame2.TextRange
        bw = float(tr2.BoundWidth)
        bh = float(tr2.BoundHeight)

        width = float(shape.Width)
        height = float(shape.Height)

        if bw > width + 1:
            return True
        if not single_line and bh > height + 1:
            return True

        return False
    except Exception:
        return False


# ============================================================
# Color Helpers
# ============================================================

def rgb_to_tuple(rgb: int):
    return (rgb & 0xFF, (rgb >> 8) & 0xFF, (rgb >> 16) & 0xFF)


def brightness(rgb: int) -> float:
    r, g, b = rgb_to_tuple(rgb)
    return 0.299 * r + 0.587 * g + 0.114 * b


def detect_slide_text_color(slide) -> int:
    try:
        bg = slide.Background.Fill.ForeColor.RGB
        return 0xFFFFFF if brightness(bg) < 128 else 0x000000
    except Exception:
        return 0x000000


# ============================================================
# Main Text Setter (NO overlap handling here)
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

    try:
        if not shp.HasTextFrame:
            return False
    except Exception:
        return False

    clean_text = "" if text is None else str(text).strip()

    # 空字串 -> 刪除 shape（或清空）
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

    if not single_line and not no_wrap:
        _expand_shape_height_to_fit_text(slide, shp, single_line=single_line)

    # ⚠️ 不在這裡做 overlap，只檢查 overflow
    if _text_still_overflows_shape(shp, single_line=single_line):
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


def replace_picture(slide, target_shape_name: str, image_path: str) -> bool:
    if not image_path or not os.path.exists(image_path):
        return False

    target = shape_by_name(slide, target_shape_name)
    if target is None:
        return False

    try:
        left = target.Left
        top = target.Top
        width = target.Width
        height = target.Height
    except Exception:
        return False

    try:
        target.Delete()
    except Exception:
        pass

    try:
        new_pic = slide.Shapes.AddPicture(
            FileName=os.path.abspath(image_path),
            LinkToFile=False,
            SaveWithDocument=True,
            Left=left,
            Top=top,
            Width=width,
            Height=height,
        )
        new_pic.Name = target_shape_name
        return True
    except Exception as e:
        print(f"[WARN] Failed to replace picture '{target_shape_name}': {e}")
        return False


def apply_images_to_placeholders(slide, slide_spec: dict, placeholder_names: list[str]) -> set:
    kept = set()
    images = slide_spec.get("images", []) or []
    if not isinstance(images, list):
        return kept

    unused = []
    by_shape_name = {}
    for img in images:
        if not isinstance(img, dict):
            continue
        path = str(img.get("image_path", "")).strip()
        if not path:
            continue
        img_shape_name = str(img.get("shape_name", "")).strip().lower()
        if img_shape_name:
            by_shape_name[img_shape_name] = path
        unused.append(path)

    for shape_name in placeholder_names:
        if not shape_by_name(slide, shape_name):
            continue

        path = by_shape_name.get(str(shape_name).strip().lower())
        if not path and unused:
            path = unused.pop(0)

        if path and replace_picture(slide, shape_name, path):
            kept.add(shape_name)

    return kept