from __future__ import annotations

from typing import Any

FLOW_KEYWORDS = {
    "step", "stage", "phase", "milestone", "process", "流程", "階段"
}

TEST_RESULT_KEYWORDS = {
    "result", "pass", "fail", "status", "limit", "spec", "lower", "upper", "測量值", "規範", "結果"
}

KPI_HINTS = {"item", "value", "metric", "kpi", "result"}


def _normalize_text(value: Any) -> str:
    return str(value or "").strip().lower()


def _flatten_text(columns: list[str], rows: list[list[str]], raw_matrix: list[list[str]]) -> list[str]:
    flattened = []
    for col in columns or []:
        flattened.append(_normalize_text(col))
    for row in rows or []:
        for cell in row:
            flattened.append(_normalize_text(cell))
    for row in raw_matrix or []:
        for cell in row:
            flattened.append(_normalize_text(cell))
    return [x for x in flattened if x]


def _is_numeric(text: str) -> bool:
    try:
        float(text.replace(",", ""))
        return True
    except Exception:
        return False


def classify_excel_block(block: dict) -> dict:
    columns = block.get("columns", []) or []
    rows = block.get("rows", []) or []
    raw_matrix = block.get("raw_matrix", []) or []
    header_detected = bool(block.get("header_detected", False))

    reasons = []
    texts = _flatten_text(columns, rows, raw_matrix)
    text_blob = " ".join(texts)

    row_lengths = [len(r) for r in rows if isinstance(r, list) and r]
    uniformity = 0.0
    if row_lengths:
        uniformity = 1.0 - (max(row_lengths) - min(row_lengths)) / max(1, max(row_lengths))

    # flow detection
    if any(k in text_blob for k in FLOW_KEYWORDS):
        reasons.append("flow keywords detected")
        return {
            "block_type": "flow",
            "confidence": 0.85,
            "reasons": reasons,
        }

    # test result detection
    if any(k in text_blob for k in TEST_RESULT_KEYWORDS):
        reasons.append("test-result keywords detected")
        return {
            "block_type": "test_result",
            "confidence": 0.82,
            "reasons": reasons,
        }

    # kpi detection (2~6 rows, mostly 2 columns and value-like second col)
    if 2 <= len(rows) <= 6 and (len(columns) == 2 or all(len(r) <= 3 for r in rows)):
        second_col_values = []
        for r in rows:
            if len(r) >= 2:
                second_col_values.append(_normalize_text(r[1]))

        numeric_ratio = 0.0
        if second_col_values:
            numeric_ratio = sum(1 for v in second_col_values if _is_numeric(v)) / len(second_col_values)

        col_hint = any(_normalize_text(c) in KPI_HINTS for c in columns)
        if numeric_ratio >= 0.3 or col_hint:
            reasons.append("small item/value style dataset")
            return {
                "block_type": "kpi",
                "confidence": 0.78,
                "reasons": reasons,
            }

    # table detection
    if header_detected and len(rows) >= 2 and uniformity >= 0.65:
        reasons.extend(["header detected", "uniform columns", f"{len(rows)} data rows"])
        return {
            "block_type": "table",
            "confidence": 0.8,
            "reasons": reasons,
        }

    # text detection
    if raw_matrix:
        nonempty_cells = [str(c).strip() for row in raw_matrix for c in row if str(c).strip()]
        text_cells = [c for c in nonempty_cells if not _is_numeric(c.lower())]
        text_ratio = len(text_cells) / max(1, len(nonempty_cells))
        if text_ratio >= 0.8 and len(nonempty_cells) <= 20 and uniformity < 0.5:
            reasons.append("high text ratio with weak table structure")
            return {
                "block_type": "text",
                "confidence": 0.7,
                "reasons": reasons,
            }

    if rows and len(columns) >= 2:
        reasons.append("fallback to table-like structure")
        return {
            "block_type": "table",
            "confidence": 0.55,
            "reasons": reasons,
        }

    return {
        "block_type": "unknown",
        "confidence": 0.4,
        "reasons": ["insufficient signals"],
    }
