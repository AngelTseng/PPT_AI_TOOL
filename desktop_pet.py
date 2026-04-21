from __future__ import annotations

import ctypes
import json
import math
import queue
import random
import sys
import threading
import tkinter as tk
from tkinter import ttk
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from classifier import classify_email
from classifier_gate import should_use_llm
from draft_generator import generate_draft
from eapprove_tracker import init_tracker_db, search_eapprove_docs, upsert_eapprove_tracking
from llm_classifier import classify_with_llm
from outlook_draft_writer import write_draft_from_record
from outlook_reader import read_recent_emails
from preprocess import preprocess_email


# ============================================================
# Settings
# ============================================================
AUTO_WRITE_DRAFT = False
WRITE_ONLY_INTERNAL_PROCESS = True
MAX_AUTO_DRAFTS_PER_CYCLE = 10
SCAN_WINDOW_LIMIT = 30
POLL_RANGE_SECONDS = (30, 60)

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    APP_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
    APP_DIR = BASE_DIR

STATE_FILE = APP_DIR / "state" / "processed_emails.json"
ASSET_DIR = BASE_DIR / "assets"

PET_SIZE = 96
PET_FLOAT_AMPLITUDE = 6
PET_RIGHT_MARGIN = 18
PET_BOTTOM_MARGIN = 18
PET_FLOAT_INTERVAL_MS = 120

PANEL_WIDTH = 860
PANEL_HEIGHT = 560

UI_STATE_IDLE = "idle"
UI_STATE_REVIEW = "review"
UI_STATE_P2 = "p2"
UI_STATE_P1 = "p1"
UI_STATE_ERROR = "error"
UI_STATE_BUSY = "busy"

PET_STATES = [
    UI_STATE_IDLE,
    UI_STATE_REVIEW,
    UI_STATE_P2,
    UI_STATE_P1,
    UI_STATE_ERROR,
]

# ============================================================
# Theme
# ============================================================
BG_APP = "#EEF3F8"
BG_CARD = "#FFFFFF"
BG_SUBTLE = "#F8FAFC"
BORDER = "#D7E2EE"
TEXT = "#0F172A"
TEXT_MUTED = "#64748B"
PRIMARY = "#3B82F6"
PRIMARY_HOVER = "#2563EB"
SUCCESS = "#10B981"
SUCCESS_HOVER = "#059669"
SECONDARY = "#E8F0FE"
SECONDARY_HOVER = "#DCE9FD"
NEUTRAL = "#F3F6FA"
NEUTRAL_HOVER = "#E8EEF5"
DANGER = "#EF4444"
P1_BG = "#FEECEC"
P2_BG = "#FFF4E8"
BTN_WIDE = 14
BTN_MEDIUM = 10


# ============================================================
# Windows work area helper
# ============================================================
class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


def get_work_area() -> tuple[int, int, int, int]:
    rect = RECT()
    SPI_GETWORKAREA = 48
    ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(rect), 0)
    return rect.left, rect.top, rect.right, rect.bottom


# ============================================================
# Data models
# ============================================================
@dataclass
class ScanSummary:
    records: list[dict[str, Any]]
    pending: list[dict[str, Any]]
    processed: list[dict[str, Any]]
    drafts_written: int
    status: str
    error: str = ""


# ============================================================
# Processed / status registry
# status:
#   pending -> still needs user handling
#   done    -> user marked as processed
# ============================================================
class ProcessedRegistry:
    def __init__(self, file_path: Path):
        self.file_path = file_path
        self.file_path.parent.mkdir(parents=True, exist_ok=True)
        self._items = self._load()

    def _load(self) -> dict[str, dict[str, Any]]:
        if not self.file_path.exists():
            return {}
        try:
            data = json.loads(self.file_path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
        except Exception:
            pass
        return {}

    def save(self) -> None:
        self.file_path.write_text(
            json.dumps(self._items, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def get(self, email_id: str) -> dict[str, Any] | None:
        return self._items.get(email_id)

    def get_status(self, email_id: str) -> str | None:
        row = self._items.get(email_id)
        if not row:
            return None
        return row.get("status")

    def is_done(self, email_id: str) -> bool:
        return self.get_status(email_id) == "done"

    def mark_pending(self, record: dict[str, Any]) -> None:
        email_id = str(record.get("id", ""))
        if not email_id:
            return

        old = self._items.get(email_id, {})
        status = old.get("status", "pending")

        self._items[email_id] = {
            "subject": record.get("subject", ""),
            "mail_type": record.get("mail_type", ""),
            "priority": record.get("priority", ""),
            "draft_written": record.get("draft_written", False),
            "action_required": record.get("action_required", False),
            "status": status,
            "updated_at": datetime.now(UTC).isoformat(),
        }

    def mark_done(self, email_id: str) -> None:
        if email_id not in self._items:
            return
        self._items[email_id]["status"] = "done"
        self._items[email_id]["updated_at"] = datetime.now(UTC).isoformat()

    def list_done(self, limit: int = 200) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        for email_id, row in self._items.items():
            if row.get("status") != "done":
                continue
            item = dict(row)
            item["id"] = email_id
            rows.append(item)
        rows.sort(key=lambda x: x.get("updated_at", ""), reverse=True)
        return rows[:limit]


# ============================================================
# Engine
# ============================================================
from rule_manager import RuleManager

class AgentEngine:
    
    def __init__(self):
        self.rule_manager = RuleManager()
        self.registry = ProcessedRegistry(STATE_FILE)
        init_tracker_db()

    def _merge_llm_result(self, email_record: dict[str, Any]) -> dict[str, Any]:
        llm_result = email_record.get("llm_result")
        if not llm_result or not llm_result.get("valid"):
            return email_record

        llm_conf = llm_result.get("confidence", "low")
        rule_mail_type = email_record.get("rule_mail_type", email_record.get("mail_type"))
        low_certainty_rule_types = {"GENERAL", "GENERAL_FORWARD", "HELP_REQUEST"}

        if llm_conf == "high":
            fields_to_override = ["mail_type", "action_type", "action_required", "priority", "reason"]
        elif rule_mail_type in low_certainty_rule_types and llm_conf in {"high", "medium"}:
            fields_to_override = ["mail_type", "action_type", "action_required", "reason"]
        else:
            email_record["llm_note"] = "llm_not_applied_due_to_confidence_policy"
            return email_record

        for key in fields_to_override:
            if key in llm_result:
                email_record[key] = llm_result[key]

        return email_record

    def _should_write_outlook_draft(self, email_record: dict[str, Any]) -> bool:
        if not email_record.get("draft_needed"):
            return False
        if WRITE_ONLY_INTERNAL_PROCESS and email_record.get("draft_purpose") != "internal_process":
            return False
        return True

    def run_cycle(self) -> ScanSummary:
        try:
            emails = read_recent_emails(limit=SCAN_WINDOW_LIMIT)
        except Exception as exc:
            return ScanSummary(
                records=[],
                pending=[],
                processed=[],
                drafts_written=0,
                status=UI_STATE_ERROR,
                error=str(exc),
            )

        records: list[dict[str, Any]] = []
        pending: list[dict[str, Any]] = []
        written_count = 0

        for email in emails:
            email_id = str(email.get("id", ""))
            if email_id and self.registry.is_done(email_id):
                continue

            email = preprocess_email(email)
            email = classify_email(email)
            email = self.rule_manager.apply(email)

            email["rule_mail_type"] = email.get("mail_type")
            email["rule_action_type"] = email.get("action_type")
            email["rule_priority"] = email.get("priority")

            if email.get("user_rule_applied"):
                use_llm = False
                gate_reason = "user_rule_override"
            else:
                use_llm, gate_reason = should_use_llm(email)

            email["gate_use_llm"] = use_llm
            email["gate_reason"] = gate_reason

            if use_llm:
                email = classify_with_llm(email)

            email = self._merge_llm_result(email)
            email["final_mail_type"] = email.get("mail_type")
            email["eapprove_tracked"] = False
            email["eapprove_track_error"] = ""
            try:
                track_result = upsert_eapprove_tracking(email)
                email["eapprove_tracked"] = bool(track_result.get("tracked"))
                if email["eapprove_tracked"]:
                    email["eapprove_track_doc_no"] = track_result.get("doc_no", "")
                    email["eapprove_track_status"] = track_result.get("status", "")
                    email["eapprove_track_flow_name"] = track_result.get("flow_name", "")
                    email["eapprove_track_comment"] = track_result.get("comment", "")
                    email["eapprove_track_date"] = track_result.get("date", "")
                else:
                    email["eapprove_track_error"] = str(track_result.get("reason", "not_tracked"))
            except Exception as exc:
                email["eapprove_track_error"] = str(exc)
            email = generate_draft(email)
            email["draft_written"] = False
            email["draft_write_error"] = ""

            existing = self.registry.get(email_id) if email_id else None
            already_drafted = bool(existing and existing.get("draft_written"))

            if (
                AUTO_WRITE_DRAFT
                and written_count < MAX_AUTO_DRAFTS_PER_CYCLE
                and not already_drafted
            ):
                if self._should_write_outlook_draft(email):
                    try:
                        email["draft_written"] = write_draft_from_record(email)
                        if email["draft_written"]:
                            written_count += 1
                    except Exception as exc:
                        email["draft_write_error"] = str(exc)
            else:
                email["draft_written"] = already_drafted

            self.registry.mark_pending(email)
            records.append(email)

            if email.get("action_required", False):
                pending.append(email)

        self.registry.save()

        if not pending:
            status = UI_STATE_IDLE
        elif any(item.get("priority") == "P1" for item in pending):
            status = UI_STATE_P1
        elif any(item.get("priority") == "P2" for item in pending):
            status = UI_STATE_P2
        else:
            status = UI_STATE_REVIEW

        return ScanSummary(
            records=records,
            pending=pending,
            processed=self.registry.list_done(),
            drafts_written=written_count,
            status=status,
        )


# ============================================================
# UI
# ============================================================
class DesktopPetUI:
    def __init__(self):
        self.engine = AgentEngine()

        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg="black")
        self.root.wm_attributes("-transparentcolor", "black")

        self.state_images = self._load_state_images()

        self.current_state = UI_STATE_IDLE
        self.pet_label = tk.Label(
            self.root,
            image=self.state_images[self.current_state],
            bg="black",
            bd=0,
            highlightthickness=0,
        )
        self.pet_label.pack(expand=True, fill="both")

        self.pet_label.bind("<Button-1>", self._toggle_panel)
        self.root.bind("<Button-3>", self._close_app)

        self.panel: tk.Toplevel | None = None
        self.pending_rows: list[dict[str, Any]] = []
        self.processed_rows: list[dict[str, Any]] = []
        self.selected_index: int | None = None
        self.last_summary: ScanSummary | None = None
        self.checked_ids: set[str] = set()
        self.scan_in_progress = False
        self.scan_result_queue: queue.Queue[ScanSummary] = queue.Queue()
        self.draft_result_queue: queue.Queue[tuple[bool, str]] = queue.Queue()
        self.pending_draft_lock = False
        self.tracking_panel: tk.Toplevel | None = None
        self.tracking_tree: ttk.Treeview | None = None

        self._floating_phase = 0.0

        self.pet_width = max(img.width() for img in self.state_images.values())
        self.pet_height = max(img.height() for img in self.state_images.values())

        _, _, right, bottom = get_work_area()
        self.base_x = right - self.pet_width - PET_RIGHT_MARGIN
        self.base_y = bottom - self.pet_height - PET_BOTTOM_MARGIN

        self._render_position()
        self._animate_float()
        self.root.after(200, self._poll_async_events)
        self._scan_and_schedule(initial=True)

    def _style_button(self, button: tk.Button, *, kind: str = "primary", width: int = BTN_WIDE) -> tk.Button:
        palette = {
            "primary": (PRIMARY, "white", PRIMARY_HOVER),
            "success": (SUCCESS, "white", SUCCESS_HOVER),
            "secondary": (SECONDARY, TEXT, SECONDARY_HOVER),
            "neutral": (NEUTRAL, TEXT_MUTED, NEUTRAL_HOVER),
        }
        bg, fg, active_bg = palette.get(kind, palette["primary"])
        button.configure(
            bg=bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground=fg,
            relief="flat",
            bd=0,
            width=width,
            pady=8,
            padx=10,
            cursor="hand2",
            highlightthickness=0,
        )
        return button

    def _build_button(self, parent, text: str, command, *, kind: str = "primary", width: int = BTN_WIDE, font=None) -> tk.Button:
        btn = tk.Button(parent, text=text, command=command, font=font or ("Microsoft JhengHei", 10, "bold"))
        return self._style_button(btn, kind=kind, width=width)

    def _setup_ttk_styles(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "Agent.Treeview",
            background=BG_CARD,
            fieldbackground=BG_CARD,
            foreground=TEXT,
            rowheight=30,
            borderwidth=0,
            relief="flat",
            font=("Microsoft JhengHei", 10),
        )
        style.configure(
            "Agent.Treeview.Heading",
            background=BG_SUBTLE,
            foreground=TEXT,
            relief="flat",
            borderwidth=0,
            font=("Microsoft JhengHei", 10, "bold"),
            padding=(8, 6),
        )
        style.map(
            "Agent.Treeview",
            background=[("selected", "#EAF2FF")],
            foreground=[("selected", TEXT)],
        )
        style.configure("Agent.Vertical.TScrollbar", troughcolor=BG_APP, borderwidth=0, arrowsize=12)
        style.configure("Agent.TPanedwindow", background=BG_APP, sashwidth=6)

    def _load_state_images(self) -> dict[str, tk.PhotoImage]:
        images: dict[str, tk.PhotoImage] = {}

        for state_name in PET_STATES:
            image_path = ASSET_DIR / f"{state_name}.png"
            if not image_path.exists():
                raise FileNotFoundError(f"Pet image not found: {image_path}")

            img = tk.PhotoImage(file=str(image_path))

            src_w = img.width()
            src_h = img.height()
            max_side = max(src_w, src_h)

            if max_side > PET_SIZE:
                scale = max(1, math.ceil(max_side / PET_SIZE))
                img = img.subsample(scale, scale)

            images[state_name] = img

        return images

    def _render_position(self) -> None:
        self.root.geometry(f"{self.pet_width}x{self.pet_height}+{self.base_x}+{self.base_y}")

    def _animate_float(self) -> None:
        self._floating_phase += 0.18
        y_offset = int(math.sin(self._floating_phase) * PET_FLOAT_AMPLITUDE)
        self.root.geometry(
            f"{self.pet_width}x{self.pet_height}+{self.base_x}+{self.base_y + y_offset}"
        )
        self.root.after(PET_FLOAT_INTERVAL_MS, self._animate_float)

    def _scan_and_schedule(self, initial: bool = False) -> None:
        self._start_background_scan()
        next_seconds = 2 if initial else random.randint(*POLL_RANGE_SECONDS)
        self.root.after(next_seconds * 1000, self._scan_and_schedule)

    @staticmethod
    def _run_with_com_init(func):
        pythoncom = None
        try:
            import pythoncom as _pythoncom  # type: ignore[import-not-found]
            pythoncom = _pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pythoncom = None

        try:
            return func()
        finally:
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _start_background_scan(self) -> None:
        if self.scan_in_progress:
            return

        self.scan_in_progress = True

        def worker():
            try:
                summary = self._run_with_com_init(self.engine.run_cycle)
            except Exception as exc:
                summary = ScanSummary(
                    records=[],
                    pending=[],
                    processed=[],
                    drafts_written=0,
                    status=UI_STATE_ERROR,
                    error=f"background_scan_failed: {exc}",
                )
            self.scan_result_queue.put(summary)

        threading.Thread(target=worker, daemon=True).start()

    def _poll_async_events(self) -> None:
        try:
            while True:
                summary = self.scan_result_queue.get_nowait()
                self.scan_in_progress = False
                self.last_summary = summary
                self.pending_rows = summary.pending[:]
                self.processed_rows = summary.processed[:]
                self.selected_index = None
                self._set_pet_state(summary.status)

                if self.panel is not None and self.panel.winfo_exists():
                    self._refresh_panel_content()
        except queue.Empty:
            pass

        try:
            while True:
                ok, msg = self.draft_result_queue.get_nowait()
                self.pending_draft_lock = False
                if hasattr(self, "detail_text"):
                    self.detail_text.config(state="normal")
                    self.detail_text.insert(tk.END, f"\n\n{msg}")
                    self.detail_text.config(state="disabled")
                if ok:
                    self._start_background_scan()
        except queue.Empty:
            pass

        self.root.after(200, self._poll_async_events)

    def _set_pet_state(self, state: str) -> None:
        self.current_state = state
        self.pet_label.configure(
            image=self.state_images.get(state, self.state_images[UI_STATE_ERROR])
        )

    def _toggle_panel(self, _event=None) -> None:
        if self.panel is not None and self.panel.winfo_exists():
            self.panel.destroy()
            self.panel = None
            return

        self.panel = tk.Toplevel(self.root)
        self.panel.title("Mail Agent - 智慧郵件管理")
        self.panel.geometry(f"{PANEL_WIDTH}x{PANEL_HEIGHT}+220+120")
        self.panel.minsize(900, 560)
        self.panel.attributes("-topmost", True)
        self.panel.configure(bg=BG_APP)
        self._setup_ttk_styles()

        font_main = ("Microsoft JhengHei", 10)
        font_bold = ("Microsoft JhengHei", 10, "bold")
        font_header = ("Microsoft JhengHei", 12, "bold")

        # Header
        header = tk.Frame(self.panel, bg=BG_CARD, height=72, highlightbackground=BORDER, highlightthickness=1)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        self.status_label = tk.Label(
            header,
            text="● 系統狀態：運行中",
            font=font_header,
            fg=PRIMARY,
            bg=BG_CARD,
            anchor="w",
        )
        self.status_label.pack(side="left", padx=14)

        btn_rules = self._build_button(
            header,
            "⚙ 規則設定",
            self._open_rule_editor,
            kind="secondary",
            width=BTN_WIDE,
            font=font_bold,
        )
        btn_rules.pack(side="right", padx=(10, 18), pady=14)

        btn_tracking = self._build_button(
            header,
            "📄 單據追蹤",
            self._open_tracking_panel,
            kind="secondary",
            width=BTN_WIDE,
            font=font_bold,
        )
        btn_tracking.pack(side="right", padx=10, pady=14)

        btn_refresh = self._build_button(
            header,
            "↻ 立即重新掃描",
            self._manual_refresh,
            kind="primary",
            width=BTN_WIDE,
            font=font_bold,
        )
        btn_refresh.pack(side="right", padx=10, pady=14)

        self.count_label = tk.Label(
            self.panel,
            text="正在準備數據...",
            font=font_main,
            bg=BG_APP,
            fg=TEXT_MUTED,
            anchor="w",
        )
        self.count_label.pack(fill="x", padx=20, pady=(10, 5))

        # Footer
        footer = tk.Frame(self.panel, bg=BG_APP, height=72)
        footer.pack(fill="x", side="bottom", padx=20, pady=(0, 14))
        footer.pack_propagate(False)

        # Main panel
        paned = ttk.PanedWindow(self.panel, orient="horizontal", style="Agent.TPanedwindow")
        paned.pack(fill="both", expand=True, padx=20, pady=10)

        # Left container
        left_container = tk.Frame(paned, bg=BG_CARD, bd=0)
        paned.add(left_container, weight=1)

        left_title_bar = tk.Frame(left_container, bg=BG_CARD)
        left_title_bar.pack(fill="x", padx=10, pady=10)

        tk.Label(
            left_title_bar,
            text="待處理清單",
            font=font_bold,
            bg=BG_CARD,
        ).pack(side="left")

        btn_group = tk.Frame(left_title_bar, bg=BG_CARD)
        btn_group.pack(side="right")

        self._build_button(
            btn_group,
            "全選",
            self._select_all_pending,
            kind="secondary",
            width=BTN_MEDIUM,
            font=font_main,
        ).pack(side="left", padx=4)

        self._build_button(
            btn_group,
            "取消",
            self._clear_all_pending_checks,
            kind="secondary",
            width=BTN_MEDIUM,
            font=font_main,
        ).pack(side="left")

        tree_frame = tk.Frame(left_container, bg=BG_CARD)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        columns = ("check", "priority", "type", "subject")
        self.pending_tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="headings",
            selectmode="browse",
            style="Agent.Treeview",
        )

        self.pending_tree.heading("check", text="狀態")
        self.pending_tree.heading("priority", text="優先級")
        self.pending_tree.heading("type", text="類型")
        self.pending_tree.heading("subject", text="郵件主旨")

        self.pending_tree.column("check", width=60, anchor="center")
        self.pending_tree.column("priority", width=70, anchor="center")
        self.pending_tree.column("type", width=120, anchor="w")
        self.pending_tree.column("subject", width=320, anchor="w")

        self.pending_tree.tag_configure("P1", background="#FEE2E2", foreground="#991B1B")
        self.pending_tree.tag_configure("P2", background="#FFEDD5", foreground="#9A3412")

        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.pending_tree.yview, style="Agent.Vertical.TScrollbar")
        self.pending_tree.configure(yscrollcommand=tree_scroll.set)
        self.pending_tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")

        self.pending_tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.pending_tree.bind("<Double-1>", self._toggle_tree_check)

        processed_title = tk.Frame(left_container, bg=BG_CARD)
        processed_title.pack(fill="x", padx=10, pady=(0, 6))
        tk.Label(
            processed_title,
            text="已處理",
            font=font_bold,
            bg=BG_CARD,
        ).pack(side="left")

        processed_frame = tk.Frame(left_container, bg=BG_CARD)
        processed_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.processed_tree = ttk.Treeview(
            processed_frame,
            columns=("priority", "type", "subject"),
            show="headings",
            height=4,
            selectmode="none",
            style="Agent.Treeview",
        )
        self.processed_tree.heading("priority", text="優先級")
        self.processed_tree.heading("type", text="類型")
        self.processed_tree.heading("subject", text="郵件主旨")
        self.processed_tree.column("priority", width=70, anchor="center")
        self.processed_tree.column("type", width=120, anchor="w")
        self.processed_tree.column("subject", width=320, anchor="w")
        self.processed_tree.pack(fill="x")

        # Right container
        right_container = tk.Frame(paned, bg=BG_CARD, bd=0)
        paned.add(right_container, weight=2)

        right_title_bar = tk.Frame(right_container, bg=BG_CARD)
        right_title_bar.pack(fill="x", padx=10, pady=10)

        tk.Label(
            right_title_bar,
            text="郵件詳細內容分析",
            font=font_bold,
            bg=BG_CARD,
        ).pack(side="left")

        detail_wrap = tk.Frame(right_container, bg=BG_CARD)
        detail_wrap.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        detail_scroll = ttk.Scrollbar(detail_wrap, orient="vertical", style="Agent.Vertical.TScrollbar")
        self.detail_text = tk.Text(
            detail_wrap,
            wrap="word",
            yscrollcommand=detail_scroll.set,
            font=("Consolas", 11),
            bg=BG_SUBTLE,
            fg=TEXT,
            padx=15,
            pady=15,
            relief="flat",
            borderwidth=0,
        )
        detail_scroll.config(command=self.detail_text.yview)

        self.detail_text.pack(side="left", fill="both", expand=True)
        detail_scroll.pack(side="right", fill="y")

        tracking_box = tk.LabelFrame(
            right_container,
            text="E-Approve Tracking",
            font=font_main,
            bg=BG_CARD,
            fg=TEXT,
            padx=12,
            pady=10,
            relief="groove",
            bd=1,
        )
        tracking_box.pack(fill="x", padx=10, pady=(0, 10))

        self.track_doc_no_var = tk.StringVar(value="-")
        self.track_status_var = tk.StringVar(value="-")
        self.track_flow_name_var = tk.StringVar(value="-")
        self.track_comment_var = tk.StringVar(value="-")
        self.track_date_var = tk.StringVar(value="-")
        self.track_error_var = tk.StringVar(value="")

        tk.Label(tracking_box, text="DOC NO:", font=font_main, bg=BG_CARD).grid(row=0, column=0, sticky="w")
        tk.Label(tracking_box, textvariable=self.track_doc_no_var, font=font_main, bg=BG_CARD, fg=TEXT).grid(row=0, column=1, sticky="w")
        tk.Label(tracking_box, text="STATUS:", font=font_main, bg=BG_CARD).grid(row=1, column=0, sticky="w")
        tk.Label(tracking_box, textvariable=self.track_status_var, font=font_main, bg=BG_CARD, fg=TEXT).grid(row=1, column=1, sticky="w")
        tk.Label(tracking_box, text="FLOW NAME:", font=font_main, bg=BG_CARD).grid(row=2, column=0, sticky="w")
        tk.Label(tracking_box, textvariable=self.track_flow_name_var, font=font_main, bg=BG_CARD, fg=TEXT).grid(row=2, column=1, sticky="w")
        tk.Label(tracking_box, text="COMMENT:", font=font_main, bg=BG_CARD).grid(row=3, column=0, sticky="w")
        tk.Label(tracking_box, textvariable=self.track_comment_var, font=font_main, bg=BG_CARD, fg=TEXT).grid(row=3, column=1, sticky="w")
        tk.Label(tracking_box, text="DATE:", font=font_main, bg=BG_CARD).grid(row=4, column=0, sticky="w")
        tk.Label(tracking_box, textvariable=self.track_date_var, font=font_main, bg=BG_CARD, fg=TEXT).grid(row=4, column=1, sticky="w")
        tk.Label(tracking_box, textvariable=self.track_error_var, font=font_main, bg=BG_CARD, fg=DANGER).grid(row=5, column=0, columnspan=2, sticky="w", pady=(4, 0))

        self._build_button(
            footer,
            "✓ 標記為已處理",
            self._mark_checked_done,
            kind="success",
            width=BTN_WIDE,
            font=font_bold,
        ).pack(side="left", padx=(0, 10), pady=8)

        tk.Button(
            footer,
            text="✉ 一鍵生成草稿",
            command=self._generate_selected_draft,
            font=font_main,
            bg=BG_CARD,
            fg="#1E293B",
            relief="solid",
            bd=1,
            padx=15,
            pady=7,
            cursor="hand2",
        ).pack(side="left")

        tk.Button(
            footer,
            text="關閉面板",
            command=self._toggle_panel,
            font=font_main,
            bg=BG_APP,
            fg=TEXT_MUTED,
            relief="flat",
            cursor="hand2",
        ).pack(side="right", padx=20)

        self._refresh_panel_content()

    def _refresh_panel_content(self) -> None:
        if self.panel is None or not self.panel.winfo_exists():
            return

        summary = self.last_summary
        if summary is None:
            return

        status_text = str(summary.status).upper()
        status_color = SUCCESS
        if "ERROR" in status_text:
            status_color = DANGER
        elif "SCANNING" in status_text or "BUSY" in status_text:
            status_color = PRIMARY

        self.status_label.configure(text=f"● 系統狀態：{summary.status}", fg=status_color)

        if summary.error:
            self.count_label.configure(
                text=f"⚠ 掃描異常：{summary.error}",
                fg=DANGER,
            )
        else:
            self.count_label.configure(
                text=(
                    f"總計待處理：{len(summary.pending)} 封  |  "
                    f"已處理：{len(summary.processed)} 封  |  "
                    f"本次掃描：{len(summary.records)} 封    "
                ),
                fg=TEXT_MUTED,
            )

        current_ids = {str(row.get("id", "")) for row in summary.pending if row.get("id")}
        self.checked_ids = {x for x in self.checked_ids if x in current_ids}

        for item in self.pending_tree.get_children():
            self.pending_tree.delete(item)

        for idx, row in enumerate(summary.pending[:100], start=1):
            email_id = str(row.get("id", ""))
            prio = str(row.get("priority", "?")).upper()
            checked_icon = "☑" if email_id in self.checked_ids else "☐"

            item_tags = ()
            if prio == "P1":
                item_tags = ("P1",)
            elif prio == "P2":
                item_tags = ("P2",)

            self.pending_tree.insert(
                "",
                "end",
                iid=email_id or f"row_{idx}",
                values=(
                    checked_icon,
                    prio,
                    row.get("mail_type", "General"),
                    row.get("subject", "無主旨")[:80],
                ),
                tags=item_tags,
            )

        for item in self.processed_tree.get_children():
            self.processed_tree.delete(item)

        for idx, row in enumerate(summary.processed[:100], start=1):
            self.processed_tree.insert(
                "",
                "end",
                iid=f"done_{idx}",
                values=(
                    str(row.get("priority", "?")).upper(),
                    row.get("mail_type", "General"),
                    row.get("subject", "無主旨")[:80],
                ),
            )

        self.detail_text.config(state="normal")
        self.detail_text.delete("1.0", tk.END)

        if summary.pending:
            tip_msg = "💡 操作提示：\n" + "─" * 30 + "\n"
            tip_msg += "1. 點擊左側列表：查看郵件詳細分析內容\n"
            tip_msg += "2. 雙擊第一欄：勾選 / 取消勾選\n"
            tip_msg += "3. 可使用全選 / 取消 / 批次標記已處理\n"
            self.detail_text.insert(tk.END, tip_msg)
        else:
            self.detail_text.insert(tk.END, "\n\n   ☕ 目前沒有待處理郵件，休息一下吧！")

        self.detail_text.config(state="disabled")
        self._update_tracking_box(None)

    def _toggle_tree_check(self, event=None) -> None:
        if not hasattr(self, "pending_tree") or event is None:
            return

        row_id = self.pending_tree.identify_row(event.y)
        col_id = self.pending_tree.identify_column(event.x)

        if not row_id:
            return

        if col_id != "#1":
            return

        if row_id in self.checked_ids:
            self.checked_ids.remove(row_id)
        else:
            self.checked_ids.add(row_id)

        current = list(self.pending_tree.item(row_id, "values"))
        if current:
            current[0] = "☑" if row_id in self.checked_ids else "☐"
            self.pending_tree.item(row_id, values=current)

    def _on_tree_select(self, _event=None) -> None:
        if self.panel is None or not hasattr(self, "pending_tree"):
            return

        sel = self.pending_tree.selection()
        if not sel:
            return

        row_id = sel[0]

        found_index = None
        for idx, row in enumerate(self.pending_rows):
            if str(row.get("id", "")) == row_id:
                found_index = idx
                break

        if found_index is None:
            return

        self.selected_index = found_index
        row = self.pending_rows[self.selected_index]

        detail = (
            f"Subject       : {row.get('subject', '')}\n"
            f"Sender        : {row.get('sender_email', '')}\n"
            f"Type          : {row.get('mail_type', '')}\n"
            f"Priority      : {row.get('priority', '')}\n"
            f"Action        : {row.get('action_type', '')}\n"
            f"Draft Needed  : {row.get('draft_needed', False)}\n"
            f"Draft Written : {row.get('draft_written', False)}\n"
            f"Draft Purpose : {row.get('draft_purpose', '')}\n"
            f"Missing Fields: {row.get('missing_fields', [])}\n"
            f"Reason        : {row.get('reason', [])}\n"
            "--------------------------------------------------\n"
            f"{row.get('clean_body', '')[:4000]}"
        )

        self.detail_text.config(state="normal")
        self.detail_text.delete("1.0", tk.END)
        self.detail_text.insert(tk.END, detail)
        self.detail_text.config(state="disabled")
        self._update_tracking_box(row)

    def _update_tracking_box(self, row: dict[str, Any] | None) -> None:
        if row is None:
            self.track_doc_no_var.set("-")
            self.track_status_var.set("-")
            self.track_flow_name_var.set("-")
            self.track_comment_var.set("-")
            self.track_date_var.set("-")
            self.track_error_var.set("")
            return

        if row.get("eapprove_tracked"):
            self.track_doc_no_var.set(str(row.get("eapprove_track_doc_no", "") or "-"))
            self.track_status_var.set(str(row.get("eapprove_track_status", "") or "-"))
            self.track_flow_name_var.set(str(row.get("eapprove_track_flow_name", "") or "-"))
            self.track_comment_var.set(str(row.get("eapprove_track_comment", "") or "-"))
            self.track_date_var.set(str(row.get("eapprove_track_date", "") or "-"))
            self.track_error_var.set("")
        else:
            self.track_doc_no_var.set("-")
            self.track_status_var.set("-")
            self.track_flow_name_var.set("-")
            self.track_comment_var.set("-")
            self.track_date_var.set("-")
            err = str(row.get("eapprove_track_error", "") or "")
            self.track_error_var.set(f"TrackErr: {err}" if err else "")

    def _select_all_pending(self) -> None:
        for row in self.pending_rows:
            email_id = str(row.get("id", ""))
            if email_id:
                self.checked_ids.add(email_id)
        self._refresh_panel_content()

    def _clear_all_pending_checks(self) -> None:
        self.checked_ids.clear()
        self._refresh_panel_content()

    def _mark_checked_done(self) -> None:
        if not self.checked_ids:
            return

        moved_ids = set(self.checked_ids)

        # 1️⃣ 寫入 registry（資料層）
        for email_id in moved_ids:
            self.engine.registry.mark_done(email_id)

        self.engine.registry.save()

        # 2️⃣ 本地資料更新（關鍵🔥）
        moved_rows = [
            row for row in self.pending_rows
            if str(row.get("id", "")) in moved_ids
        ]

        self.pending_rows = [
            row for row in self.pending_rows
            if str(row.get("id", "")) not in moved_ids
        ]

        for row in moved_rows:
            row["status"] = "done"

        self.processed_rows = moved_rows + self.processed_rows

        self.checked_ids.clear()

        # 3️⃣ 同步到 summary（你UI真正吃這個🔥）
        if self.last_summary:
            self.last_summary.pending = self.pending_rows[:]
            self.last_summary.processed = self.processed_rows[:]

        # 4️⃣ 立即刷新 UI（🔥重點）
        self._refresh_panel_content()

        # 5️⃣ 背景同步（非阻塞）
        self._start_background_scan()

    def _manual_refresh(self) -> None:
        self._start_background_scan()

    def _mark_selected_done(self) -> None:
        if self.selected_index is None:
            return
        if self.selected_index >= len(self.pending_rows):
            return

        row = self.pending_rows[self.selected_index]
        email_id = str(row.get("id", ""))
        if not email_id:
            return

        self.engine.registry.mark_done(email_id)
        self.engine.registry.save()

        if email_id in self.checked_ids:
            self.checked_ids.remove(email_id)

        self._start_background_scan()

    def _generate_selected_draft(self) -> None:
        if self.selected_index is None:
            return
        if self.selected_index >= len(self.pending_rows):
            return

        row = self.pending_rows[self.selected_index]
        self.detail_text.config(state="normal")

        if not row.get("draft_needed"):
            self.detail_text.insert(tk.END, "\n\n此信件不需要 draft。")
            self.detail_text.config(state="disabled")
            return

        if self.pending_draft_lock:
            self.detail_text.insert(tk.END, "\n\n已有草稿任務進行中，請稍候。")
            self.detail_text.config(state="disabled")
            return

        self.pending_draft_lock = True
        self.detail_text.insert(tk.END, "\n\n草稿生成中，請稍候...")
        self.detail_text.config(state="disabled")

        def worker(target_row: dict[str, Any]):
            try:
                success = bool(
                    self._run_with_com_init(lambda: write_draft_from_record(target_row))
                )
                if success:
                    self.draft_result_queue.put((True, "已一鍵生成 Outlook 草稿。"))
                else:
                    self.draft_result_queue.put((False, "Outlook 草稿建立失敗（回傳 False）。"))
            except Exception as exc:
                self.draft_result_queue.put((False, f"開草稿失敗：{exc}"))

        threading.Thread(target=worker, args=(dict(row),), daemon=True).start()

    def _close_app(self, _event=None) -> None:
        if self.panel is not None and self.panel.winfo_exists():
            self.panel.destroy()
        self.root.destroy()
        
    def _open_rule_editor(self):
        win = tk.Toplevel(self.panel)
        
        win.title("規則設定")
        win.geometry("500x400")
        
        win.attributes("-topmost", True)
        win.lift()
        win.focus_force()

        manager = self.engine.rule_manager

        # listbox 顯示規則
        listbox = tk.Listbox(win)
        listbox.pack(fill="both", expand=True, padx=10, pady=10)

        def refresh():
            listbox.delete(0, tk.END)
            for r in manager.rules:
                listbox.insert(tk.END, f"{r.get('name')} → {r.get('priority')}")

        refresh()

        # ===== 新增規則區 =====
        frame = tk.Frame(win)
        frame.pack(fill="x", padx=10, pady=5)

        tk.Label(frame, text="名稱").grid(row=0, column=0)
        name_entry = tk.Entry(frame)
        name_entry.grid(row=0, column=1)

        tk.Label(frame, text="寄件者").grid(row=1, column=0)
        sender_entry = tk.Entry(frame)
        sender_entry.grid(row=1, column=1)

        tk.Label(frame, text="關鍵字").grid(row=2, column=0)
        keyword_entry = tk.Entry(frame)
        keyword_entry.grid(row=2, column=1)

        tk.Label(frame, text="優先級").grid(row=3, column=0)
        priority_var = tk.StringVar(value="P2")
        ttk.Combobox(frame, textvariable=priority_var, values=["P1", "P2", "REVIEW"]).grid(row=3, column=1)

        def add_rule():
            rule = {
                "name": name_entry.get(),
                "sender_contains": [sender_entry.get()] if sender_entry.get() else [],
                "subject_contains": [keyword_entry.get()] if keyword_entry.get() else [],
                "priority": priority_var.get(),
            }
            manager.add_rule(rule)
            refresh()
            self._start_background_scan()

        def delete_rule():
            sel = listbox.curselection()
            if sel:
                manager.delete_rule(sel[0])
                refresh()
                self._start_background_scan()
                
        tk.Button(win, text="新增規則", command=add_rule).pack(pady=5)
        tk.Button(win, text="刪除選取規則", command=delete_rule).pack(pady=5)
    
    def _open_tracking_panel(self) -> None:
        if self.tracking_panel is not None and self.tracking_panel.winfo_exists():
            self.tracking_panel.lift()
            self.tracking_panel.focus_force()
            self._refresh_tracking_panel()
            return

        panel = tk.Toplevel(self.panel if self.panel is not None else self.root)
        panel.title("單據追蹤")
        panel.geometry("900x500+260+140")
        panel.minsize(760, 360)
        panel.attributes("-topmost", True)
        panel.configure(bg=BG_CARD)
        self.tracking_panel = panel
        self._setup_ttk_styles()

        top_bar = tk.Frame(panel, bg=BG_CARD)
        top_bar.pack(fill="x", padx=12, pady=(12, 6))
        tk.Label(
            top_bar,
            text="單據追蹤總覽（最新狀態）",
            font=("Microsoft JhengHei", 11, "bold"),
            bg=BG_CARD,
        ).pack(side="left")
        self._build_button(
            top_bar,
            "↻ 重新整理",
            self._refresh_tracking_panel,
            kind="secondary",
            width=BTN_WIDE,
            font=("Microsoft JhengHei", 10, "bold"),
        ).pack(side="right", padx=12, pady=10)

        tree_frame = tk.Frame(panel, bg=BG_CARD)
        tree_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        columns = ("doc_no", "status", "flow_name", "event_date", "updated_at")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse", style="Agent.Treeview")
        tree.heading("doc_no", text="DOC NO")
        tree.heading("status", text="STATUS")
        tree.heading("flow_name", text="FLOW")
        tree.heading("event_date", text="DATE")
        tree.heading("updated_at", text="UPDATED AT")
        tree.column("doc_no", width=140, anchor="w")
        tree.column("status", width=120, anchor="center")
        tree.column("flow_name", width=320, anchor="w")
        tree.column("event_date", width=140, anchor="center")
        tree.column("updated_at", width=180, anchor="center")

        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview, style="Agent.Vertical.TScrollbar")
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.tracking_tree = tree
        self._refresh_tracking_panel()

    def _refresh_tracking_panel(self) -> None:
        if self.tracking_tree is None or not self.tracking_tree.winfo_exists():
            return

        rows = search_eapprove_docs("")
        for item in self.tracking_tree.get_children():
            self.tracking_tree.delete(item)

        for idx, row in enumerate(rows, start=1):
            self.tracking_tree.insert(
                "",
                "end",
                iid=f"doc_{idx}",
                values=(
                    str(row.get("doc_no", "") or "-"),
                    str(row.get("status", "") or "-"),
                    str(row.get("flow_name", "") or "-"),
                    str(row.get("event_date", "") or "-"),
                    str(row.get("updated_at", "") or "-"),
                ),
            )

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    DesktopPetUI().run()