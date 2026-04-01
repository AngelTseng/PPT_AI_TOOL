from renderer_helper import apply_images_to_placeholders, set_text, shape_by_name

def _adjust_title_background(slide):
    title = shape_by_name(slide, "title")
    bg = shape_by_name(slide, "矩形: 剪去同側角落 ")  

    if not title or not bg:
        return

    try:
        padding = 2

        # 👉 目標底部
        target_bottom = title.Top + title.Width + padding

        # 👉 新高度（Top 不動，只往下長）
        new_height = target_bottom - bg.Top

        # 👉 只允許變大（避免縮回去）
        if new_height > bg.Height:
            bg.Height = new_height

    except Exception as e:
        print("[WARN] adjust title background failed:", e)

def render_content_image(slide, slide_spec):
    keep_names = set()

    if shape_by_name(slide, "title"):
        keep_names.add("title")
        set_text(slide, "title", str(slide_spec.get("title", "")), no_wrap=True, single_line=True)

    if shape_by_name(slide, "content"):
        keep_names.add("content")
        set_text(slide, "content", str(slide_spec.get("content", "")))

    keep_names |= apply_images_to_placeholders(
        slide,
        slide_spec,
        ["img", "main_image", "img_1", "img_2"],
    )

    return keep_names


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

    keep_names |= apply_images_to_placeholders(slide, slide_spec, ["img_1", "img_2", "img"])
    return keep_names


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

    keep_names |= apply_images_to_placeholders(slide, slide_spec, ["img_1", "img_2", "img"])
    return keep_names


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


def render_content_3extra(slide, slide_spec):
    keep_names = set()

    if shape_by_name(slide, "title"):
        keep_names.add("title")
        set_text(slide, "title", str(slide_spec.get("title", "")), no_wrap=True, single_line=True)

    cards = slide_spec.get("cards", [])
    for i in range(1, 4):
        item_name = f"item_{i}"
        content_name = f"content_{i}"

        if shape_by_name(slide, item_name):
            keep_names.add(item_name)
            if i <= len(cards):
                set_text(slide, item_name, str(cards[i - 1].get("item", "")))
            else:
                set_text(slide, item_name, "")

        if shape_by_name(slide, content_name):
            keep_names.add(content_name)
            if i <= len(cards):
                set_text(slide, content_name, str(cards[i - 1].get("content", "")))
            else:
                set_text(slide, content_name, "")

    return keep_names


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