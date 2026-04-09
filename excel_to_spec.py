from __future__ import annotations

from pathlib import Path

from excel_block_classifier import classify_excel_block

MAX_BLOCKS = 10
MAX_TABLE_COLUMNS = 6
MAX_TABLE_ROWS = 12
MAX_FLOW_STEPS = 6


FLOW_KEYS = ["step", "stage", "phase", "milestone", "process", "流程", "階段"]


def _safe_str(value) -> str:
    return str(value or "").strip()


def _table_slide(title: str, columns: list[str], rows: list[list[str]]) -> dict:
    clipped_columns = [_safe_str(c) or f"Column {i+1}" for i, c in enumerate(columns[:MAX_TABLE_COLUMNS])]
    clipped_rows = []
    for row in rows[:MAX_TABLE_ROWS]:
        clipped_rows.append([_safe_str(x) for x in row[: len(clipped_columns)]])

    if not clipped_columns:
        clipped_columns = ["Column 1"]
    if not clipped_rows:
        clipped_rows = [[""]]

    return {
        "type": "table",
        "title": title,
        "columns": clipped_columns,
        "rows": clipped_rows,
    }


def _kpi_cards(block: dict) -> list[dict]:
    rows = block.get("rows", []) or []
    columns = block.get("columns", []) or []

    cards = []
    for row in rows:
        if not isinstance(row, list) or not row:
            continue

        item = _safe_str(row[0])
        value = _safe_str(row[1]) if len(row) > 1 else ""

        if not item and len(columns) >= 1:
            item = _safe_str(columns[0])
        if not value and len(row) > 1:
            value = _safe_str(" / ".join(row[1:]))

        if item or value:
            cards.append({"item": item or "Metric", "content": value or "N/A"})

    return cards


def _kpi_slide(title: str, cards: list[dict]) -> dict:
    n = len(cards)
    if n <= 2:
        return {"type": "content_2", "title": title, "cards": cards[:2]}
    if n == 3:
        return {"type": "content_3", "title": title, "cards": cards[:3]}
    return {"type": "content_4", "title": title, "cards": cards[:4]}


def _extract_flow_steps(block: dict) -> list[str]:
    columns = [_safe_str(c).lower() for c in block.get("columns", []) or []]
    rows = block.get("rows", []) or []

    step_col_idx = None
    for idx, c in enumerate(columns):
        if any(k in c for k in FLOW_KEYS):
            step_col_idx = idx
            break

    steps = []
    if step_col_idx is not None:
        for row in rows:
            if step_col_idx < len(row):
                val = _safe_str(row[step_col_idx])
                if val:
                    steps.append(val)
    else:
        for row in rows:
            if row:
                val = _safe_str(row[0])
                if val:
                    steps.append(val)

    unique_steps = []
    seen = set()
    for s in steps:
        key = s.lower()
        if key not in seen:
            seen.add(key)
            unique_steps.append(s)

    return unique_steps[:MAX_FLOW_STEPS]


def _text_fallback_slide(title: str, block: dict) -> dict:
    raw = block.get("raw_matrix", []) or []
    lines = []
    for row in raw:
        line = " ".join(_safe_str(c) for c in row if _safe_str(c))
        if line:
            lines.append(line)

    content = "\n".join(lines[:8]).strip() or "內容摘要"

    return {
        "type": "content_text",
        "title": title,
        "content": content,
    }


def excel_payload_to_spec(payload: dict) -> dict:
    workbook_name = _safe_str(payload.get("workbook_name")) or "Workbook"
    topic = Path(workbook_name).stem or "Excel Report"

    slides = [
        {
            "type": "cover",
            "topic": topic,
            "speaker": "",
        }
    ]

    processed_blocks = 0

    for sheet in payload.get("sheets", []):
        sheet_name = _safe_str(sheet.get("sheet_name")) or "Sheet"

        for block in sheet.get("blocks", []):
            if processed_blocks >= MAX_BLOCKS:
                break

            block_title = _safe_str(block.get("title")) or sheet_name
            cls = classify_excel_block(block)
            block_type = cls.get("block_type", "unknown")

            if block_type == "table":
                slide = _table_slide(block_title, block.get("columns", []), block.get("rows", []))
                slide["debug"] = {"block_id": block.get("block_id"), "block_type": block_type, "confidence": cls.get("confidence")}
                slides.append(slide)

            elif block_type == "kpi":
                cards = _kpi_cards(block)
                if len(cards) < 2:
                    slide = _table_slide(block_title, block.get("columns", []), block.get("rows", []))
                else:
                    slide = _kpi_slide(block_title, cards)
                slide["debug"] = {"block_id": block.get("block_id"), "block_type": block_type, "confidence": cls.get("confidence")}
                slides.append(slide)

            elif block_type == "flow":
                steps = _extract_flow_steps(block)
                if len(steps) >= 2:
                    slide = {
                        "type": "flow",
                        "title": block_title,
                        "steps": steps,
                        "debug": {"block_id": block.get("block_id"), "block_type": block_type, "confidence": cls.get("confidence")},
                    }
                else:
                    slide = _table_slide(block_title, block.get("columns", []), block.get("rows", []))
                    slide["debug"] = {"block_id": block.get("block_id"), "block_type": "fallback_table", "confidence": cls.get("confidence")}
                slides.append(slide)

            elif block_type == "test_result":
                table_slide = _table_slide(block_title, block.get("columns", []), block.get("rows", []))
                table_slide["debug"] = {"block_id": block.get("block_id"), "block_type": block_type, "confidence": cls.get("confidence")}
                slides.append(table_slide)

                # Optional summary for pass/fail counts
                all_text = " ".join(" ".join(_safe_str(c).lower() for c in row) for row in block.get("rows", []))
                pass_count = all_text.count("pass") + all_text.count("通過")
                fail_count = all_text.count("fail") + all_text.count("失敗")
                if pass_count + fail_count > 0:
                    slides.append(
                        {
                            "type": "content_2",
                            "title": f"{block_title} Summary",
                            "cards": [
                                {"item": "Pass", "content": str(pass_count)},
                                {"item": "Fail", "content": str(fail_count)},
                            ],
                            "debug": {"block_id": block.get("block_id"), "block_type": "test_result_summary"},
                        }
                    )

            elif block_type in {"text", "unknown"}:
                fallback_slide = _text_fallback_slide(block_title, block)
                fallback_slide["debug"] = {"block_id": block.get("block_id"), "block_type": block_type, "confidence": cls.get("confidence")}
                slides.append(fallback_slide)

            else:
                fallback_table = _table_slide(block_title, block.get("columns", []), block.get("rows", []))
                fallback_table["debug"] = {"block_id": block.get("block_id"), "block_type": "fallback_table"}
                slides.append(fallback_table)

            processed_blocks += 1

        if processed_blocks >= MAX_BLOCKS:
            break

    slides.append({"type": "end"})
    return {"slides": slides}
