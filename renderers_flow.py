def render_flow(slide, slide_spec, set_text, _resolve_flow_prefer_name, ensure_smartart_nodes, reduce_smartart_nodes, fill_smartart_steps):
    keep_names = {"title"}
    set_text(slide, "title", str(slide_spec.get("title", "")), no_wrap=True, single_line=True)
    steps = slide_spec.get("steps", [])

    prefer_name = _resolve_flow_prefer_name(slide, slide_spec)
    if prefer_name:
        keep_names.add(prefer_name)

    smart_shape, current_count = ensure_smartart_nodes(slide, len(steps), prefer_name=prefer_name)

    if smart_shape is not None and current_count > len(steps):
        _, final_count = reduce_smartart_nodes(slide, len(steps), prefer_name=prefer_name)

    fill_smartart_steps(slide, steps, prefer_name=prefer_name)
    return keep_names