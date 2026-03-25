from renderer_helper import set_text


def render_cover(slide, slide_spec):
    keep_names = {"Topic", "speaker_name"}
    set_text(slide, "Topic", str(slide_spec.get("topic", "")), no_wrap=True, single_line=True)
    set_text(slide, "speaker_name", str(slide_spec.get("speaker", "")))
    return keep_names


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


def render_end(slide, slide_spec):
    return set()