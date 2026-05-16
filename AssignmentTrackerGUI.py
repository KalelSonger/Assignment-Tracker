import importlib.util
import json
import os
import queue
import socket
import subprocess
import sys
import threading
import time
import traceback
import urllib.parse
import urllib.request
import ctypes
from contextlib import redirect_stderr, redirect_stdout

import tkinter as tk
from tkinter import messagebox, ttk

from keys import CONFIG_SOURCE


SINGLE_INSTANCE_HOST = "127.0.0.1"
SINGLE_INSTANCE_PORT = 48523
CANVAS_SESSION_FILE = "canvas_session.local.json"
SHEET_ENDPOINTS_FILE = "sheet_endpoints.local.json"
SHEET_URL_PLACEHOLDER = "add sheet url here"


class SingleInstanceGuard:
    def __init__(self, host: str = SINGLE_INSTANCE_HOST, port: int = SINGLE_INSTANCE_PORT):
        self.host = host
        self.port = port
        self.sock = None

    def acquire(self) -> bool:
        try:
            self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.sock.bind((self.host, self.port))
            self.sock.listen(1)
            return True
        except OSError:
            self.release()
            return False

    def release(self):
        if self.sock is not None:
            try:
                self.sock.close()
            except OSError:
                pass
            self.sock = None


class QueueWriter:
    def __init__(self, out_queue: queue.Queue):
        self.out_queue = out_queue
        self.buffer = ""

    def write(self, text: str):
        self.buffer += text
        while "\n" in self.buffer:
            line, self.buffer = self.buffer.split("\n", 1)
            if line.strip():
                self.out_queue.put(line)

    def flush(self):
        if self.buffer.strip():
            self.out_queue.put(self.buffer)
        self.buffer = ""


class AssignmentTrackerGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Assignment Tracker")
        self.geometry("1100x700")
        self._set_windows_app_id()
        self._set_window_icon()

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.playwright_manager = None
        self.browser = None
        self.context = None
        self.page = None
        self.sheet_patterns = None
        self.allowed_tabs = []
        self.storage_state = None
        self.sheet_registry = {"selected_api_url": "", "sheets": []}
        self.sheet_name_to_url: dict[str, str] = {}
        self.selected_sheet_name_var = tk.StringVar(value="")
        self.sheet_url_input_var = tk.StringVar(value="")
        self.login_hint_var = tk.StringVar(value="Preparing dependencies...")
        self.reopen_login_button = None
        self.sheet_dropdown = None
        self.sheet_url_entry = None
        self.sheet_url_has_placeholder = False
        self.awaiting_initial_sheet_url = False
        self.top_controls_frame = None
        self.top_controls_notice_var = tk.StringVar(value="")
        self.sync_running = False
        self.backend = None

        self.state_dir = self._state_dir()
        self.canvas_session_path = os.path.join(self.state_dir, CANVAS_SESSION_FILE)
        self.sheet_endpoints_path = os.path.join(self.state_dir, SHEET_ENDPOINTS_FILE)
        self._ensure_local_state_files()

        self._build_ui()
        self.after(100, self._drain_logs)

        threading.Thread(target=self._bootstrap_and_start_login, daemon=True).start()

    def _set_window_icon(self):
        icon_candidates = []

        if hasattr(sys, "_MEIPASS"):
            icon_candidates.append(os.path.join(sys._MEIPASS, "app.ico"))

        if getattr(sys, "frozen", False):
            exe_dir = os.path.dirname(sys.executable)
            icon_candidates.append(os.path.join(exe_dir, "app.ico"))

        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_candidates.extend(
            [
                os.path.join(script_dir, "app.ico"),
                os.path.join(script_dir, "assets", "app.ico"),
                os.path.join(os.getcwd(), "app.ico"),
                os.path.join(os.getcwd(), "assets", "app.ico"),
            ]
        )

        for icon_path in icon_candidates:
            if os.path.isfile(icon_path):
                try:
                    self.iconbitmap(default=icon_path)
                    break
                except Exception:
                    pass

    def _set_windows_app_id(self):
        if os.name != "nt":
            return
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("AssignmentTracker.GUI")
        except Exception:
            pass

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.main_frame = ttk.Frame(self)
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)

        header_frame = ttk.Frame(self.main_frame)
        header_frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
        header_frame.grid_columnconfigure(0, weight=1)
        header_frame.grid_columnconfigure(1, weight=0)

        self.status_var = tk.StringVar(value="Starting...")
        status_label = ttk.Label(header_frame, textvariable=self.status_var, font=("Segoe UI", 12, "bold"))
        status_label.grid(row=0, column=0, sticky="w")

        self.top_controls_frame = ttk.Frame(header_frame)
        self.top_controls_frame.grid(row=0, column=1, sticky="e")

        self.content_frame = ttk.Frame(self.main_frame)
        self.content_frame.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        self.content_frame.grid_columnconfigure(0, weight=0)
        self.content_frame.grid_columnconfigure(1, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)

        self.left_panel = ttk.Frame(self.content_frame)
        self.left_panel.grid(row=0, column=0, sticky="nsw", padx=(0, 12))

        self.right_panel = ttk.Frame(self.content_frame)
        self.right_panel.grid(row=0, column=1, sticky="nsew")
        self.right_panel.grid_columnconfigure(0, weight=1)
        self.right_panel.grid_rowconfigure(0, weight=1)

        self.log_text = tk.Text(self.right_panel, wrap="word", state="disabled", font=("Consolas", 10))
        self.log_text.grid(row=0, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(self.right_panel, orient="vertical", command=self.log_text.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scroll.set)

        self._show_login_panel()

    def _render_top_sheet_controls(self, show_prompt: bool = False):
        if self.top_controls_frame is None:
            return

        for child in self.top_controls_frame.winfo_children():
            child.destroy()

        notice_text = ""
        if show_prompt:
            notice_text = "No sheet saved yet. Add a sheet URL to continue."

        self.top_controls_notice_var.set(notice_text)
        if notice_text:
            ttk.Label(
                self.top_controls_frame,
                textvariable=self.top_controls_notice_var,
                foreground="#555555",
            ).grid(row=0, column=0, columnspan=3, sticky="e", pady=(0, 2))
            controls_row = 1
        else:
            controls_row = 0

        dropdown_state = "readonly" if self.sheet_registry.get("sheets") else "disabled"
        self.sheet_dropdown = ttk.Combobox(
            self.top_controls_frame,
            textvariable=self.selected_sheet_name_var,
            state=dropdown_state,
            width=30,
        )
        self.sheet_dropdown.grid(row=controls_row, column=0, padx=(0, 8), sticky="e")
        self.sheet_dropdown.bind("<<ComboboxSelected>>", self._on_sheet_selected)

        self.sheet_url_entry = ttk.Entry(self.top_controls_frame, textvariable=self.sheet_url_input_var, width=38)
        self.sheet_url_entry.grid(row=controls_row, column=1, padx=(0, 8), sticky="e")
        self.sheet_url_entry.bind("<FocusIn>", self._on_sheet_url_focus_in)
        self.sheet_url_entry.bind("<FocusOut>", self._on_sheet_url_focus_out)
        self._apply_sheet_url_placeholder()

        ttk.Button(
            self.top_controls_frame,
            text="Add",
            command=self._handle_add_sheet_from_login if self.awaiting_initial_sheet_url else self._add_sheet_endpoint,
            width=10,
        ).grid(row=controls_row, column=2, sticky="e")

        self._refresh_sheet_dropdown()

    def _state_dir(self) -> str:
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))

    def _ensure_local_state_files(self):
        os.makedirs(self.state_dir, exist_ok=True)
        if not os.path.isfile(self.sheet_endpoints_path):
            initial = {"selected_api_url": "", "sheets": []}
            with open(self.sheet_endpoints_path, "w", encoding="utf-8") as file:
                json.dump(initial, file, indent=2)

    def _show_login_panel(self):
        for child in self.left_panel.winfo_children():
            child.destroy()

        self._render_top_sheet_controls(show_prompt=not self.sheet_registry.get("sheets"))

        ttk.Label(
            self.left_panel,
            text="Canvas Sign-In",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", pady=(0, 8))

        ttk.Label(
            self.left_panel,
            text=(
                "A browser window will open for UMSYSTEM login.\n"
                "Complete sign-in there. This app will detect login\n"
                "automatically and unlock sync buttons."
            ),
            justify="left",
        ).pack(anchor="w")

        ttk.Separator(self.left_panel, orient="horizontal").pack(fill="x", pady=10)
        ttk.Label(self.left_panel, textvariable=self.login_hint_var, foreground="#444").pack(anchor="w")

        self.reopen_login_button = ttk.Button(
            self.left_panel,
            text="Reopen browser",
            command=self._retry_login_browser,
            state="disabled",
        )
        self.reopen_login_button.pack(anchor="w", pady=(8, 0))

    def _show_sync_panel(self):
        for child in self.left_panel.winfo_children():
            child.destroy()

        self._render_top_sheet_controls(show_prompt=False)

        class_sync_labels = [f"Sync: {class_tab}" for class_tab in self.allowed_tabs]
        class_clear_labels = [f"Clear: {class_tab}" for class_tab in self.allowed_tabs]
        base_labels = [
            "Sync all assignments",
            "Sync future assignments",
            "Dry-sync (no writes)",
            "Clear all class tabs",
            "Close",
        ]
        button_width = self._button_width_for_labels(base_labels + class_sync_labels + class_clear_labels)

        ttk.Separator(self.left_panel, orient="horizontal").pack(fill="x", pady=10)

        ttk.Label(self.left_panel, text="Sync Actions", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 8))

        ttk.Button(
            self.left_panel,
            text="Sync all assignments",
            command=lambda: self._start_sync(include_past=True, dry_run=False, replace_existing=False),
            width=button_width,
        ).pack(anchor="w", pady=4)

        ttk.Button(
            self.left_panel,
            text="Sync future assignments",
            command=lambda: self._start_sync(include_past=False, dry_run=False, replace_existing=False),
            width=button_width,
        ).pack(anchor="w", pady=4)

        ttk.Button(
            self.left_panel,
            text="Dry-sync (no writes)",
            command=lambda: self._start_sync(include_past=True, dry_run=True, replace_existing=False),
            width=button_width,
        ).pack(anchor="w", pady=4)

        ttk.Label(self.left_panel, text="Sync individual class tab:").pack(anchor="w", pady=(8, 2))
        for class_tab in self.allowed_tabs:
            ttk.Button(
                self.left_panel,
                text=f"Sync: {class_tab}",
                command=lambda tab_name=class_tab: self._start_sync_single_tab(tab_name),
                width=button_width,
            ).pack(anchor="w", pady=2)

        ttk.Separator(self.left_panel, orient="horizontal").pack(fill="x", pady=10)
        ttk.Label(self.left_panel, text="Clear Actions", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 6))

        ttk.Button(
            self.left_panel,
            text="Clear all class tabs",
            command=self._start_clear_all_tabs,
            width=button_width,
        ).pack(anchor="w", pady=4)

        ttk.Label(self.left_panel, text="Clear individual class tab:").pack(anchor="w", pady=(8, 2))
        for class_tab in self.allowed_tabs:
            ttk.Button(
                self.left_panel,
                text=f"Clear: {class_tab}",
                command=lambda tab_name=class_tab: self._start_clear_single_tab(tab_name),
                width=button_width,
            ).pack(anchor="w", pady=2)

        ttk.Separator(self.left_panel, orient="horizontal").pack(fill="x", pady=10)
        ttk.Button(self.left_panel, text="Close", command=self._close_app, width=button_width).pack(anchor="w", pady=4)

    def _button_width_for_labels(self, labels: list[str], min_width: int = 30, padding: int = 2) -> int:
        if not labels:
            return min_width
        longest = max(len(label) for label in labels)
        return max(min_width, longest + padding)

    def _log(self, message: str):
        self.log_queue.put(message)

    def _drain_logs(self):
        try:
            while True:
                line = self.log_queue.get_nowait()
                self.log_text.configure(state="normal")
                self.log_text.insert("end", line + "\n")
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._drain_logs)

    def _load_sheet_registry(self):
        try:
            with open(self.sheet_endpoints_path, "r", encoding="utf-8") as file:
                raw = json.load(file)
        except Exception:
            raw = {"selected_api_url": "", "sheets": []}

        selected_api_url = str(raw.get("selected_api_url") or "").strip() if isinstance(raw, dict) else ""
        sheets_raw = raw.get("sheets", []) if isinstance(raw, dict) else []

        sheets: list[dict] = []
        seen_urls: set[str] = set()
        for entry in sheets_raw:
            if not isinstance(entry, dict):
                continue
            api_url = str(entry.get("api_url") or "").strip()
            if not api_url:
                continue
            normalized = self._normalize_api_url(api_url)
            if normalized in seen_urls:
                continue
            seen_urls.add(normalized)
            display_name = str(entry.get("display_name") or "").strip() or self._fallback_sheet_name(api_url)
            sheets.append({"api_url": api_url, "display_name": display_name})

        if not selected_api_url and sheets:
            selected_api_url = sheets[0]["api_url"]

        self.sheet_registry = {"selected_api_url": selected_api_url, "sheets": sheets}
        self._save_sheet_registry()

    def _save_sheet_registry(self):
        with open(self.sheet_endpoints_path, "w", encoding="utf-8") as file:
            json.dump(self.sheet_registry, file, indent=2)

    def _fallback_sheet_name(self, api_url: str) -> str:
        if not api_url:
            return "Unnamed sheet"
        parsed = urllib.parse.urlparse(api_url)
        host = parsed.netloc or "sheet"
        return f"Sheet @ {host}"

    def _safe_infer_sheet_name(self, api_url: str) -> str:
        payload = urllib.parse.urlencode({"action": "tabs"}).encode("utf-8")
        request = urllib.request.Request(api_url, data=payload, method="POST")
        request.add_header("Content-Type", "application/x-www-form-urlencoded")

        try:
            with urllib.request.urlopen(request, timeout=15) as response:
                raw = response.read().decode("utf-8")
            parsed = json.loads(raw)
        except Exception:
            return self._fallback_sheet_name(api_url)

        if isinstance(parsed, dict):
            for key in ("spreadsheetName", "spreadsheet_name", "sheetName", "sheet_name", "title", "name"):
                value = parsed.get(key)
                if isinstance(value, str) and value.strip():
                    return value.strip()

        return self._fallback_sheet_name(api_url)

    def _refresh_sheet_dropdown(self):
        names = []
        self.sheet_name_to_url = {}
        dedupe: dict[str, int] = {}

        for item in self.sheet_registry.get("sheets", []):
            base_name = str(item.get("display_name") or "").strip() or self._fallback_sheet_name(item.get("api_url", ""))
            count = dedupe.get(base_name, 0) + 1
            dedupe[base_name] = count
            display_name = base_name if count == 1 else f"{base_name} ({count})"
            api_url = str(item.get("api_url") or "").strip()
            if not api_url:
                continue

            names.append(display_name)
            self.sheet_name_to_url[display_name] = api_url

        if self.sheet_dropdown is not None:
            self.sheet_dropdown["values"] = names
            longest_name = max((len(name) for name in names), default=24)
            dynamic_width = max(24, min(72, longest_name + 2))
            self.sheet_dropdown.configure(width=dynamic_width)

        selected_url = str(self.sheet_registry.get("selected_api_url") or "").strip()
        selected_name = ""
        for name, api_url in self.sheet_name_to_url.items():
            if api_url == selected_url:
                selected_name = name
                break

        if not selected_name and names:
            selected_name = names[0]
            self.sheet_registry["selected_api_url"] = self.sheet_name_to_url[selected_name]
            self._save_sheet_registry()

        self.selected_sheet_name_var.set(selected_name)

    def _selected_sheet_api_url(self) -> str:
        selected_name = self.selected_sheet_name_var.get().strip()
        if selected_name and selected_name in self.sheet_name_to_url:
            return self.sheet_name_to_url[selected_name]
        return str(self.sheet_registry.get("selected_api_url") or "").strip()

    def _on_sheet_selected(self, _event=None):
        api_url = self._selected_sheet_api_url()
        if not api_url:
            return
        self.sheet_registry["selected_api_url"] = api_url
        self._save_sheet_registry()
        self._reload_selected_sheet_tabs()

    def _normalize_api_url(self, api_url: str) -> str:
        cleaned = str(api_url or "").strip()
        if not cleaned:
            return ""

        parsed = urllib.parse.urlparse(cleaned)
        scheme = (parsed.scheme or "https").lower()
        netloc = parsed.netloc.lower()
        path = parsed.path.rstrip("/")
        return urllib.parse.urlunparse((scheme, netloc, path, "", "", ""))

    def _add_sheet_endpoint(self, reload_tabs: bool = True):
        raw_url = self.sheet_url_input_var.get().strip()
        if self.sheet_url_has_placeholder and raw_url == SHEET_URL_PLACEHOLDER:
            raw_url = ""
        if not raw_url:
            messagebox.showwarning("Missing URL", "Paste a Google Apps Script API URL first.")
            return

        if not raw_url.startswith("http://") and not raw_url.startswith("https://"):
            messagebox.showwarning("Invalid URL", "Sheet API URL must start with http:// or https://")
            return

        normalized_incoming = self._normalize_api_url(raw_url)
        existing = next(
            (
                item
                for item in self.sheet_registry.get("sheets", [])
                if self._normalize_api_url(item.get("api_url", "")) == normalized_incoming
            ),
            None,
        )
        if existing is None:
            name = self._safe_infer_sheet_name(raw_url)
            self.sheet_registry.setdefault("sheets", []).append({"api_url": raw_url, "display_name": name})
            self._log(f"Added sheet endpoint: {name}")
        else:
            self._log("Sheet endpoint already exists; selecting it.")
            raw_url = str(existing.get("api_url") or raw_url)

        self.sheet_registry["selected_api_url"] = raw_url
        self._save_sheet_registry()
        self.sheet_url_input_var.set("")
        self.sheet_url_has_placeholder = False
        self._refresh_sheet_dropdown()
        if reload_tabs:
            self._reload_selected_sheet_tabs()
        self._apply_sheet_url_placeholder()

    def _handle_add_sheet_from_login(self):
        self._add_sheet_endpoint(reload_tabs=False)
        if not self.sheet_registry.get("sheets"):
            return

        if self.awaiting_initial_sheet_url:
            self.awaiting_initial_sheet_url = False
            self._set_status("Sheet saved. Continuing startup...")
            self._set_login_hint("Initializing with selected sheet...")
            threading.Thread(target=self._bootstrap_and_start_login, daemon=True).start()

    def _apply_sheet_url_placeholder(self):
        if self.sheet_url_entry is None:
            return
        if self.sheet_url_input_var.get().strip():
            return
        self.sheet_url_has_placeholder = True
        self.sheet_url_input_var.set(SHEET_URL_PLACEHOLDER)
        self.sheet_url_entry.configure(foreground="#777777")

    def _on_sheet_url_focus_in(self, _event=None):
        if not self.sheet_url_has_placeholder:
            return
        self.sheet_url_has_placeholder = False
        self.sheet_url_input_var.set("")
        self.sheet_url_entry.configure(foreground="#000000")

    def _on_sheet_url_focus_out(self, _event=None):
        if self.sheet_url_input_var.get().strip():
            self.sheet_url_entry.configure(foreground="#000000")
            return
        self._apply_sheet_url_placeholder()

    def _set_reopen_login_enabled(self, enabled: bool):
        def update_button():
            if self.reopen_login_button is not None:
                self.reopen_login_button.configure(state="normal" if enabled else "disabled")

        self.after(0, update_button)

    def _dispose_login_browser(self):
        try:
            if self.page is not None:
                self.page.close()
        except Exception:
            pass
        self.page = None

        try:
            if self.context is not None:
                self.context.close()
        except Exception:
            pass
        self.context = None

        try:
            if self.browser is not None:
                self.browser.close()
        except Exception:
            pass
        self.browser = None

        try:
            if self.playwright_manager is not None:
                self.playwright_manager.stop()
        except Exception:
            pass
        self.playwright_manager = None

    def _save_canvas_session(self):
        if self.storage_state is None:
            return
        with open(self.canvas_session_path, "w", encoding="utf-8") as file:
            json.dump(self.storage_state, file, indent=2)

    def _clear_canvas_session(self):
        self.storage_state = None
        if os.path.isfile(self.canvas_session_path):
            try:
                os.remove(self.canvas_session_path)
            except OSError:
                pass

    def _load_canvas_session_from_disk(self):
        if not os.path.isfile(self.canvas_session_path):
            return None
        try:
            with open(self.canvas_session_path, "r", encoding="utf-8") as file:
                data = json.load(file)
            return data if isinstance(data, dict) else None
        except Exception:
            return None

    def _is_storage_state_authenticated(self, storage_state: dict) -> bool:
        from playwright.sync_api import sync_playwright

        with sync_playwright() as p:
            api_context = p.request.new_context(storage_state=storage_state)
            try:
                return self.backend._is_canvas_authenticated(api_context)
            finally:
                api_context.dispose()

    def _reload_selected_sheet_tabs(self):
        if self.backend is None:
            return
        if self.sync_running:
            self._log("Cannot switch sheet while sync is running.")
            return

        api_url = self._selected_sheet_api_url()
        if not api_url:
            self._log("No sheet endpoint selected.")
            return

        try:
            self.backend.set_sheet_api_url(api_url)
            self.allowed_tabs = self.backend.fetch_allowed_sheet_classes()
            self.sheet_patterns = self.backend._build_sheet_class_patterns(self.allowed_tabs)
            self._set_status(f"Loaded {len(self.allowed_tabs)} tab(s) for selected sheet")
            self._log(f"Selected sheet endpoint: {api_url}")
            self._log(f"Loaded {len(self.allowed_tabs)} class tabs from selected sheet.")
            self.after(0, self._show_sync_panel)
        except Exception as error:
            self._log(f"Sheet reload error: {error}")
            self._log(traceback.format_exc())
            self._set_status("Sheet load failed")

    def _retry_login_browser(self):
        self._set_reopen_login_enabled(False)
        threading.Thread(target=self._reopen_login_worker, daemon=True).start()

    def _reopen_login_worker(self):
        try:
            self._set_status("Waiting for Canvas sign-in...")
            self._set_login_hint("Opening browser for sign-in...")
            self._open_browser_for_login()
            self._wait_for_login_status()
        except Exception as error:
            self._set_status("Login retry failed")
            self._set_login_hint("See console log for details.")
            self._log(f"Login retry error: {error}")
            self._log(traceback.format_exc())
            self._set_reopen_login_enabled(True)

    def _bootstrap_and_start_login(self):
        self._set_status("Checking dependencies...")
        self._log("Checking dependencies...")

        try:
            self._ensure_dependencies()
            self._log("Dependencies ready.")

            import PullFromCanvas as backend_module
            self.backend = backend_module

            self._load_sheet_registry()

            selected_api_url = self._selected_sheet_api_url()
            if not selected_api_url.strip():
                self.awaiting_initial_sheet_url = True
                self._set_status("Sheet URL required")
                self._set_login_hint("Add a sheet URL at the top to continue.")
                self._log("No saved sheet endpoints found. Add a sheet URL in the top field.")
                self.after(0, self._show_login_panel)
                return

            self.awaiting_initial_sheet_url = False

            self.backend.set_sheet_api_url(selected_api_url)

            self._log(f"Config source: {CONFIG_SOURCE}")
            self._log(f"Active sheet API URL: {self.backend.get_sheet_api_url()}")

            self._log("Loading selected sheet class tabs...")
            allowed_tabs = self.backend.fetch_allowed_sheet_classes()
            self.allowed_tabs = allowed_tabs
            self.sheet_patterns = self.backend._build_sheet_class_patterns(allowed_tabs)
            self._log(f"Loaded {len(allowed_tabs)} class tabs from selected sheet.")

            self._set_status("Checking saved Canvas session...")
            saved_state = self._load_canvas_session_from_disk()
            if saved_state and self._is_storage_state_authenticated(saved_state):
                self.storage_state = saved_state
                self._set_status("Signed in. Ready to sync.")
                self._set_login_hint("Saved Canvas session restored.")
                self._log("Using saved Canvas session. No login needed.")
                self.after(0, self._show_sync_panel)
                return

            if saved_state:
                self._log("Saved Canvas session expired. Re-login required.")
                self._clear_canvas_session()

            self._set_status("Waiting for Canvas sign-in...")
            self._set_login_hint("Opening browser for sign-in...")
            self._open_browser_for_login()
            self._wait_for_login_status()
        except Exception as error:
            self._set_status("Startup failed")
            self._set_login_hint("See console log for details.")
            self._log(f"Startup error: {error}")
            self._log(traceback.format_exc())

    def _ensure_dependencies(self):
        if getattr(sys, "frozen", False):
            if importlib.util.find_spec("playwright") is None:
                raise RuntimeError(
                    "Playwright is not available in this EXE build. Rebuild after installing playwright in the build environment."
                )

            self._log("Running as EXE: skipping pip/install bootstrap to avoid self-relaunch loops.")
            return

        if importlib.util.find_spec("playwright") is None:
            self._log("Installing playwright package...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "playwright"])

        self._log("Ensuring Chromium is installed for Playwright...")
        subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"])

    def _open_browser_for_login(self):
        from playwright.sync_api import sync_playwright

        self._dispose_login_browser()

        self.playwright_manager = sync_playwright().start()

        launch_plan = [
            ("chromium", {}),
            ("msedge", {"channel": "msedge"}),
            ("chrome", {"channel": "chrome"}),
        ]
        if getattr(sys, "frozen", False):
            launch_plan = [
                ("msedge", {"channel": "msedge"}),
                ("chrome", {"channel": "chrome"}),
                ("chromium", {}),
            ]

        launch_errors: list[str] = []
        launched_with = ""

        for browser_name, launch_kwargs in launch_plan:
            try:
                self.browser = self.playwright_manager.chromium.launch(headless=False, **launch_kwargs)
                launched_with = browser_name
                break
            except Exception as error:
                launch_errors.append(f"{browser_name}: {error}")

        if self.browser is None:
            error_text = "\n".join(launch_errors)
            if getattr(sys, "frozen", False):
                raise RuntimeError(
                    "Could not launch a browser for login. Tried Edge, Chrome, then bundled Chromium. "
                    "Ensure Microsoft Edge or Google Chrome is installed, then try again.\n\n"
                    f"Launch details:\n{error_text}"
                )
            raise RuntimeError(
                "Could not launch a browser for login. If Chromium is missing, run: "
                "python -m playwright install chromium\n\n"
                f"Launch details:\n{error_text}"
            )

        self.context = self.browser.new_context()
        self.page = self.context.new_page()
        self.page.goto(self.backend.LOGIN_URL, wait_until="domcontentloaded")
        self._log(f"{launched_with.capitalize()} opened. Complete Canvas/Microsoft sign-in in that window.")
        self._set_reopen_login_enabled(False)

    def _wait_for_login_status(self, timeout_seconds: int = 300, poll_interval_seconds: float = 1.5):
        started = time.monotonic()
        self._set_login_hint("Waiting for sign-in completion...")

        while (time.monotonic() - started) < timeout_seconds:
            if self.backend._is_canvas_authenticated(self.context.request):
                self.storage_state = self.context.storage_state()
                self._save_canvas_session()
                self._set_status("Signed in. Ready to sync.")
                self._set_login_hint("Sign-in detected.")
                self._log("Canvas sign-in detected.")
                self._dispose_login_browser()
                self.after(0, self._show_sync_panel)
                return
            self.page.wait_for_timeout(int(poll_interval_seconds * 1000))

        self._dispose_login_browser()
        self._set_status("Canvas login timed out")
        self._set_login_hint("Browser closed. Click 'Reopen browser' to try sign-in again.")
        self._set_reopen_login_enabled(True)
        self._log("Canvas login timed out. Browser closed; use 'Reopen browser' to retry.")

    def _start_sync(self, include_past: bool, dry_run: bool, replace_existing: bool):
        if self.sync_running:
            self._log("An operation is already running. Please wait.")
            return

        self.sync_running = True
        self._set_status("Sync running...")
        threading.Thread(
            target=self._run_sync_worker,
            args=(include_past, dry_run, replace_existing, None),
            daemon=True,
        ).start()

    def _start_sync_single_tab(self, class_tab: str):
        if self.sync_running:
            self._log("An operation is already running. Please wait.")
            return

        self.sync_running = True
        self._set_status(f"Syncing {class_tab}...")
        threading.Thread(
            target=self._run_sync_worker,
            args=(True, False, False, [class_tab]),
            daemon=True,
        ).start()

    def _start_clear_all_tabs(self):
        if self.sync_running:
            self._log("An operation is already running. Please wait.")
            return

        if not messagebox.askyesno("Confirm clear all", "Clear all class tabs on the sheet?"):
            return

        self.sync_running = True
        self._set_status("Clearing all class tabs...")
        threading.Thread(target=self._run_clear_worker, args=(None,), daemon=True).start()

    def _start_clear_single_tab(self, class_tab: str):
        if self.sync_running:
            self._log("An operation is already running. Please wait.")
            return

        if not messagebox.askyesno("Confirm clear tab", f"Clear tab '{class_tab}'?"):
            return

        self.sync_running = True
        self._set_status(f"Clearing {class_tab}...")
        threading.Thread(target=self._run_clear_worker, args=(class_tab,), daemon=True).start()

    def _run_clear_worker(self, class_tab: str | None):
        writer = QueueWriter(self.log_queue)
        try:
            with redirect_stdout(writer), redirect_stderr(writer):
                if self.backend is None:
                    raise RuntimeError("Backend is not loaded.")

                if class_tab:
                    response = self.backend.clear_single_class_tab(class_tab)
                    print(f"Cleared tab '{class_tab}'. Rows cleared: {response.get('clearedRows', 0)}")
                    self._set_status(f"Cleared {class_tab}")
                else:
                    response = self.backend.clear_all_class_tabs()
                    print(f"Cleared all class tabs. Total rows cleared: {response.get('clearedRows', 0)}")
                    for entry in response.get("clearedTabs", []):
                        print(f"- {entry.get('sheetName')}: {entry.get('clearedRows', 0)} rows cleared")
                    self._set_status("Cleared all class tabs")
        except Exception as error:
            self._log(f"Clear action error: {error}")
            self._log(traceback.format_exc())
            self._set_status("Clear action failed")
        finally:
            writer.flush()
            self.sync_running = False

    def _run_sync_worker(
        self,
        include_past: bool,
        dry_run: bool,
        replace_existing: bool,
        selected_tabs: list[str] | None,
    ):
        writer = QueueWriter(self.log_queue)
        try:
            from playwright.sync_api import sync_playwright

            with redirect_stdout(writer), redirect_stderr(writer):
                if self.storage_state is None:
                    raise RuntimeError("No Canvas login session available. Please sign in again.")

                storage_state = self.storage_state

                with sync_playwright() as p:
                    api_context = p.request.new_context(storage_state=storage_state)

                    if not self.backend._is_canvas_authenticated(api_context):
                        self._clear_canvas_session()
                        raise RuntimeError(
                            "Canvas session expired. Click 'Reopen browser' from the sign-in panel to login again."
                        )

                    class RequestContextShim:
                        def __init__(self, request):
                            self.request = request

                    shim = RequestContextShim(api_context)

                    patterns_to_use = self.sheet_patterns
                    if selected_tabs:
                        selected_set = set(selected_tabs)
                        patterns_to_use = [
                            pattern
                            for pattern in (self.sheet_patterns or [])
                            if pattern.get("tab_name") in selected_set
                        ]
                        if not patterns_to_use:
                            raise RuntimeError(
                                f"No sheet pattern found for selected tab(s): {', '.join(selected_tabs)}"
                            )

                    if selected_tabs:
                        print(f"Sync limited to tab(s): {', '.join(selected_tabs)}")

                    assignments_by_class = self.backend.fetch_assignments_from_canvas_context(
                        shim,
                        patterns_to_use,
                        include_past_assignments=include_past,
                    )

                    file_count = self.backend.write_outputs_by_class(assignments_by_class, self.backend.OUTPUT_DIR)
                    total_assignments = sum(len(records) for records in assignments_by_class.values())
                    print(f"Saved {total_assignments} assignments into {file_count} file(s) in '{self.backend.OUTPUT_DIR}'.")

                    sync_response = self.backend.sync_assignments_to_sheet(
                        assignments_by_class,
                        dry_run=dry_run,
                        replace_existing=replace_existing,
                    )

                    print(f"Sheet sync response saved to {self.backend.SHEET_SYNC_RESPONSE_FILE}")
                    print(f"Sheet sync status: {sync_response.get('status', 'unknown')}")
                    print(f"Rows written: {sync_response.get('rowsWritten', 0)}")
                    if sync_response.get("dryRun"):
                        print("Dry run mode: no spreadsheet changes were made.")
                    for message in sync_response.get("debugMessages", []):
                        print(message)
                    for class_name, stats in sync_response.get("classStats", {}).items():
                        print(
                            f"[{class_name}] incoming={stats.get('incomingCount', 0)} "
                            f"existing={stats.get('existingNamedCount', 0)} matched={stats.get('matchedCount', 0)} "
                            f"added={stats.get('addedCount', 0)} updated={stats.get('updatedCount', 0)}"
                        )

                    api_context.dispose()

            self._set_status("Sync completed")
        except Exception as error:
            self._log(f"Sync error: {error}")
            self._log(traceback.format_exc())
            self._set_status("Sync failed")
            if "session expired" in str(error).lower() or "please sign in" in str(error).lower():
                self.after(0, self._show_login_panel)
                self._set_reopen_login_enabled(True)
        finally:
            writer.flush()
            self.sync_running = False

    def _set_status(self, value: str):
        self.after(0, lambda: self.status_var.set(value))

    def _set_login_hint(self, value: str):
        self.after(0, lambda: self.login_hint_var.set(value))

    def _close_app(self):
        self._dispose_login_browser()

        self.destroy()


def main():
    instance_guard = SingleInstanceGuard()
    if not instance_guard.acquire():
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Assignment Tracker", "Assignment Tracker is already running.")
        root.destroy()
        return

    app = AssignmentTrackerGUI()
    app.protocol("WM_DELETE_WINDOW", app._close_app)
    app.mainloop()
    instance_guard.release()


if __name__ == "__main__":
    main()
