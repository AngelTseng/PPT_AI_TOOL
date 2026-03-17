import json
import os
import win32com.client as win32
import pythoncom

from com_utils import com_session


def shape_by_name(slide, name: str):
    for i in range(1, slide.Shapes.Count + 1):
        shp = slide.Shapes(i)
        if shp.Name == name:
            return shp
    return None

def extract_shape_names(slide):
    names = []

    for i in range(1, slide.Shapes.Count + 1):
        try:
            names.append(slide.Shapes(i).Name)
        except Exception:
            pass

    return names

def get_text_from_shape(slide, shape_name: str) -> str:
    shp = shape_by_name(slide, shape_name)
    if shp is None:
        return ""

    try:
        if not shp.HasTextFrame:
            return ""
        return shp.TextFrame.TextRange.Text.strip()
    except Exception:
        return ""

def extract_text_boxes(slide):
    text_boxes = []

    for i in range(1, slide.Shapes.Count + 1):
        shp = slide.Shapes(i)

        try:
            name = shp.Name
        except Exception:
            name = f"shape_{i}"

        try:
            if shp.HasTextFrame:
                text = shp.TextFrame.TextRange.Text.strip()
                if text:
                    text_boxes.append({
                        "shape_name": name,
                        "text": text,
                    })
        except Exception:
            pass

    return text_boxes

def get_table_data(slide, table_name: str):
    shp = shape_by_name(slide, table_name)
    if shp is None:
        return [], []

    try:
        if not shp.HasTable:
            return [], []

        tbl = shp.Table
        row_count = tbl.Rows.Count
        col_count = tbl.Columns.Count

        if row_count < 1 or col_count < 1:
            return [], []

        columns = []
        for c in range(1, col_count + 1):
            text = tbl.Cell(1, c).Shape.TextFrame.TextRange.Text.strip()
            columns.append(text)

        rows = []
        for r in range(2, row_count + 1):
            row_data = []
            has_any_text = False

            for c in range(1, col_count + 1):
                text = tbl.Cell(r, c).Shape.TextFrame.TextRange.Text.strip()
                if text:
                    has_any_text = True
                row_data.append(text)

            if has_any_text:
                rows.append(row_data)

        return columns, rows

    except Exception:
        return [], []

def extract_images(slide):
    images = []

    for i in range(1, slide.Shapes.Count + 1):
        shp = slide.Shapes(i)

        try:
            name = shp.Name
        except Exception:
            name = f"shape_{i}"

        try:
            shape_type = shp.Type
        except Exception:
            shape_type = None

        # msoLinkedPicture=11, msoPicture=13
        if shape_type in (11, 13):
            try:
                images.append({
                    "shape_name": name,
                    "shape_type": shape_type,
                    "left": shp.Left,
                    "top": shp.Top,
                    "width": shp.Width,
                    "height": shp.Height,
                })
            except Exception:
                images.append({
                    "shape_name": name,
                    "shape_type": shape_type,
                })

    return images

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


def get_smartart_steps(slide, prefer_name=None):
    shp = find_smartart_shape(slide, prefer_name=prefer_name)
    if shp is None:
        return []

    steps = []

    try:
        nodes = shp.SmartArt.AllNodes
        for i in range(1, nodes.Count + 1):
            node = nodes(i)
            text = ""

            try:
                text = node.TextFrame2.TextRange.Text.strip()
            except Exception:
                pass

            if not text:
                try:
                    text = node.TextFrame.TextRange.Text.strip()
                except Exception:
                    pass

            if not text:
                try:
                    text = node.Shapes(1).TextFrame2.TextRange.Text.strip()
                except Exception:
                    pass

            if not text:
                try:
                    text = node.Shapes(1).TextFrame.TextRange.Text.strip()
                except Exception:
                    pass

            if text:
                steps.append(text)

    except Exception:
        return []

    return steps




def detect_flow_prefer_name(slide):
    for name in ("flow_chart_1", "flow_chart_2", "flow_chart_3"):
        shp = shape_by_name(slide, name)
        try:
            if shp is not None and getattr(shp, "HasSmartArt", False):
                return name
        except Exception:
            pass
    return None

def detect_slide_type(slide) -> str:
    # cover
    if shape_by_name(slide, "Topic") and shape_by_name(slide, "speaker_name"):
        return "cover"

    # agenda
    if shape_by_name(slide, "outline") and shape_by_name(slide, "agenda_1"):
        return "agenda"

    # section
    if shape_by_name(slide, "agenda_name"):
        return "section"

    # content_image
    if shape_by_name(slide, "title") and shape_by_name(slide, "content") and shape_by_name(slide, "img"):
        return "content_image"

    # content_text
    if (
        shape_by_name(slide, "title")
        and shape_by_name(slide, "content")
        and not shape_by_name(slide, "img")
        and not shape_by_name(slide, "img_1")
        and not shape_by_name(slide, "main_image")
    ):
        return "content_text"

    # content_4 variants
    if shape_by_name(slide, "content_1") and shape_by_name(slide, "content_2") and shape_by_name(slide, "content_3") and shape_by_name(slide, "content_4"):
        if shape_by_name(slide, "item_1") and shape_by_name(slide, "item_2") and shape_by_name(slide, "item_3") and shape_by_name(slide, "item_4"):
            return "content_4_b"
        return "content_4_a"

    # content_3extra
    if shape_by_name(slide, "item_1") and shape_by_name(slide, "item_2") and shape_by_name(slide, "item_3") and shape_by_name(slide, "content_1") and shape_by_name(slide, "content_2") and shape_by_name(slide, "content_3"):
        return "content_3extra"

    # content_2 variants
    if shape_by_name(slide, "item_1") and shape_by_name(slide, "item_2") and shape_by_name(slide, "content_1") and shape_by_name(slide, "content_2") and not shape_by_name(slide, "item_3") and not shape_by_name(slide, "content_3"):
        return "content_2_c"

    if shape_by_name(slide, "title") and shape_by_name(slide, "content_1") and shape_by_name(slide, "content_2") and shape_by_name(slide, "img_1") and shape_by_name(slide, "img_2"):
        return "content_2_a"

    if shape_by_name(slide, "content_1") and shape_by_name(slide, "content_2") and shape_by_name(slide, "img_1") and shape_by_name(slide, "img_2"):
        return "content_2_b"

    # table
    shp = shape_by_name(slide, "sheet_1")
    try:
        if shp is not None and shp.HasTable:
            return "table"
    except Exception:
        pass

    # flow (support three variants)
    if detect_flow_prefer_name(slide):
        return "flow"

    return "unknown"


def extract_cover(slide):
    return {
        "type": "cover",
        "topic": get_text_from_shape(slide, "Topic"),
        "speaker": get_text_from_shape(slide, "speaker_name"),
        "text_boxes": extract_text_boxes(slide),
        "images": extract_images(slide),
    }


def extract_agenda(slide):
    items = []
    for i in range(1, 6):
        text = get_text_from_shape(slide, f"agenda_{i}")
        if text:
            items.append(text)

    return {
        "type": "agenda",
        "title": get_text_from_shape(slide, "outline") or "Agenda",
        "items": items,
        "text_boxes": extract_text_boxes(slide),
        "images": extract_images(slide),
    }


def extract_section(slide):
    return {
        "type": "section",
        "name": get_text_from_shape(slide, "agenda_name"),
    }


def extract_content_3extra(slide):
    cards = []

    for i in range(1, 4):
        item = get_text_from_shape(slide, f"item_{i}")
        content = get_text_from_shape(slide, f"content_{i}")

        if item or content:
            cards.append({
                "item": item,
                "content": content,
            })

    return {
        "type": "content_3extra",
        "title": get_text_from_shape(slide, "title"),
        "cards": cards,
    }


def extract_content_2(slide, variant: str):
    cards = []
    for i in range(1, 3):
        item = get_text_from_shape(slide, f"item_{i}") or f"Point {i}"
        content = get_text_from_shape(slide, f"content_{i}")
        if item or content:
            cards.append({"item": item, "content": content})

    return {
        "type": variant,
        "title": get_text_from_shape(slide, "title"),
        "cards": cards,
    }


def extract_content_4(slide, variant: str):
    cards = []
    for i in range(1, 5):
        item = get_text_from_shape(slide, f"item_{i}") or f"Point {i}"
        content = get_text_from_shape(slide, f"content_{i}")
        if item or content:
            cards.append({"item": item, "content": content})

    return {
        "type": variant,
        "title": get_text_from_shape(slide, "title"),
        "cards": cards,
    }


def extract_content_image(slide):
    return {
        "type": "content_image",
        "title": get_text_from_shape(slide, "title"),
        "content": get_text_from_shape(slide, "content"),
        "images": extract_images(slide),
    }

def extract_content_text(slide):
    return {
        "type": "content_text",
        "title": get_text_from_shape(slide, "title"),
        "content": get_text_from_shape(slide, "content"),
    }

def extract_table(slide):
    columns, rows = get_table_data(slide, "sheet_1")
    return {
        "type": "table",
        "title": get_text_from_shape(slide, "title"),
        "columns": columns,
        "rows": rows,
    }


def extract_flow(slide):
    prefer_name = detect_flow_prefer_name(slide)
    return {
        "type": "flow",
        "title": get_text_from_shape(slide, "title"),
        "steps": get_smartart_steps(slide, prefer_name=prefer_name),
    }

def extract_unknown(slide):
    has_table = False
    has_smartart = False

    for i in range(1, slide.Shapes.Count + 1):
        shp = slide.Shapes(i)

        try:
            if shp.HasTable:
                has_table = True
        except Exception:
            pass

        try:
            if getattr(shp, "HasSmartArt", False):
                has_smartart = True
        except Exception:
            pass

    return {
        "type": "unknown",
        "slide_index": slide.SlideIndex,
        "shape_names": extract_shape_names(slide),
        "text_boxes": extract_text_boxes(slide),
        "images": extract_images(slide),
        "has_table": has_table,
        "has_smartart": has_smartart,
    }
    

def extract_slide(slide):
    slide_type = detect_slide_type(slide)

    if slide_type == "cover":
        return extract_cover(slide)
    elif slide_type == "agenda":
        return extract_agenda(slide)
    elif slide_type == "section":
        return extract_section(slide)
    
    elif slide_type == "content_3extra":
        return extract_content_3extra(slide)
    elif slide_type in ("content_2", "content_2_a", "content_2_b", "content_2_c"):
        return extract_content_2(slide, slide_type)
    elif slide_type in ("content_4", "content_4_a", "content_4_b"):
        return extract_content_4(slide, slide_type)
    elif slide_type == "content_image":
        return extract_content_image(slide)
    elif slide_type == "content_text":
        return extract_content_text(slide)
    elif slide_type == "table":
        return extract_table(slide)
    
    elif slide_type == "flow":
        return extract_flow(slide)
    else:
        return extract_unknown(slide)

def extract_ppt_to_spec(input_pptx: str) -> dict:
    with com_session():
        app = None
        pres = None

        try:
            app = win32.Dispatch("PowerPoint.Application")
            app.Visible = True

            pres = app.Presentations.Open(os.path.abspath(input_pptx), WithWindow=False)

            slides = []
            for i in range(1, pres.Slides.Count + 1):
                slide = pres.Slides(i)
                slides.append(extract_slide(slide))

            return {"slides": slides}

        finally:
            if pres is not None:
                pres.Close()
            if app is not None:
                app.Quit()
                
                
import sys

def main():
    
    if len(sys.argv) < 2:
        print("Usage: python extract_ppt_content.py input.pptx [output.json]")
        return

    input_pptx = sys.argv[1]

    if len(sys.argv) >= 3:
        output_json = sys.argv[2]
    else:
        output_json = "extracted_deck_spec.json"

    spec = extract_ppt_to_spec(input_pptx)

    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False, indent=2)

    print(f"[INFO] Extracted spec saved to: {output_json}")


if __name__ == "__main__":
    main()