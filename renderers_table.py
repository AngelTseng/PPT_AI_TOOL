def render_table_slide(slide, slide_spec, set_text, fill_table):
    keep_names = {"title", "sheet_1"}
    set_text(slide, "title", str(slide_spec.get("title", "")), no_wrap=True, single_line=True)
    fill_table(slide, "sheet_1", slide_spec.get("columns", []), slide_spec.get("rows", []))
    return keep_names