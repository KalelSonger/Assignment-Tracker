import importlib.util
import os
import queue
import subprocess
import sys
import threading
import traceback
from contextlib import redirect_stderr, redirect_stdout

import tkinter as tk
from tkinter import ttk

from keys import SHEET_API_URL


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
        self._set_window_icon()

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.playwright_manager = None
        self.browser = None
        self.context = None
        self.page = None
        self.sheet_patterns = None
        self.sync_running = False
        self.backend = None

        self._build_ui()
        self.after(100, self._drain_logs)

        threading.Thread(target=self._bootstrap_and_start_login, daemon=True).start()

    def _set_window_icon(self):
        icon_candidates = [
            os.path.join(os.getcwd(), "app.ico"),
            os.path.join(os.getcwd(), "assets", "app.ico"),
        ]

        for icon_path in icon_candidates:
            if os.path.isfile(icon_path):
                try:
                    self.iconbitmap(icon_path)
                    break
                except Exception:
                    pass

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.main_frame = ttk.Frame(self)
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)

        self.status_var = tk.StringVar(value="Starting...")
        status_label = ttk.Label(self.main_frame, textvariable=self.status_var, font=("Segoe UI", 12, "bold"))
        status_label.grid(row=0, column=0, sticky="w", padx=12, pady=(10, 6))

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

    def _show_login_panel(self):
        for child in self.left_panel.winfo_children():
            child.destroy()

        ttk.Label(
            self.left_panel,
            text="Canvas Sign-In",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", pady=(0, 8))

        ttk.Label(
            self.left_panel,
            text=(
                "A Chromium window will open for UMSYSTEM login.\n"
                "Complete sign-in there. This app will detect login\n"
                "automatically and unlock sync buttons."
            ),
            justify="left",
        ).pack(anchor="w")

        ttk.Separator(self.left_panel, orient="horizontal").pack(fill="x", pady=10)

        self.login_hint_var = tk.StringVar(value="Preparing dependencies...")
        ttk.Label(self.left_panel, textvariable=self.login_hint_var, foreground="#444").pack(anchor="w")

    def _show_sync_panel(self):
        for child in self.left_panel.winfo_children():
            child.destroy()

        ttk.Label(self.left_panel, text="Sync Actions", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 8))

        ttk.Button(
            self.left_panel,
            text="Sync all assignments",
            command=lambda: self._start_sync(include_past=True, dry_run=False, replace_existing=True),
            width=30,
        ).pack(anchor="w", pady=4)

        ttk.Button(
            self.left_panel,
            text="Sync future assignments",
            command=lambda: self._start_sync(include_past=False, dry_run=False, replace_existing=True),
            width=30,
        ).pack(anchor="w", pady=4)

        ttk.Button(
            self.left_panel,
            text="Dry-sync (no writes)",
            command=lambda: self._start_sync(include_past=True, dry_run=True, replace_existing=True),
            width=30,
        ).pack(anchor="w", pady=4)

        ttk.Separator(self.left_panel, orient="horizontal").pack(fill="x", pady=10)
        ttk.Button(self.left_panel, text="Close", command=self._close_app, width=30).pack(anchor="w", pady=4)

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

    def _bootstrap_and_start_login(self):
        self._set_status("Checking dependencies...")
        self._log("Checking dependencies...")

        try:
            self._ensure_dependencies()
            self._log("Dependencies ready.")

            import PullFromCanvas as backend_module
            self.backend = backend_module

            if not SHEET_API_URL.strip():
                raise RuntimeError("keys.py has empty SHEET_API_URL. Fill it before running GUI.")

            self._log("Loading sheet class tabs...")
            allowed_tabs = self.backend.fetch_allowed_sheet_classes()
            self.sheet_patterns = self.backend._build_sheet_class_patterns(allowed_tabs)
            self._log(f"Loaded {len(allowed_tabs)} class tabs from sheet.")

            self._set_status("Waiting for Canvas sign-in...")
            self._set_login_hint("Opening Chromium for sign-in...")
            self._open_browser_for_login()
            self._poll_login_status()
        except Exception as error:
            self._set_status("Startup failed")
            self._set_login_hint("See console log for details.")
            self._log(f"Startup error: {error}")
            self._log(traceback.format_exc())

    def _ensure_dependencies(self):
        if importlib.util.find_spec("playwright") is None:
            self._log("Installing playwright package...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "playwright"])

        self._log("Ensuring Chromium is installed for Playwright...")
        subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"])

    def _open_browser_for_login(self):
        from playwright.sync_api import sync_playwright

        self.playwright_manager = sync_playwright().start()
        self.browser = self.playwright_manager.chromium.launch(headless=False)
        self.context = self.browser.new_context()
        self.page = self.context.new_page()
        self.page.goto(self.backend.LOGIN_URL, wait_until="domcontentloaded")
        self._log("Chromium opened. Complete Canvas/Microsoft sign-in in that window.")

    def _poll_login_status(self):
        try:
            if self.backend._is_canvas_authenticated(self.context.request):
                self._set_status("Signed in. Ready to sync.")
                self._set_login_hint("Sign-in detected.")
                self._log("Canvas sign-in detected.")
                self._show_sync_panel()
                return

            self._set_login_hint("Waiting for sign-in completion...")
            self.after(1500, self._poll_login_status)
        except Exception as error:
            self._set_status("Login check failed")
            self._log(f"Login check error: {error}")
            self._log(traceback.format_exc())

    def _start_sync(self, include_past: bool, dry_run: bool, replace_existing: bool):
        if self.sync_running:
            self._log("A sync is already running. Please wait.")
            return

        self.sync_running = True
        self._set_status("Sync running...")
        threading.Thread(
            target=self._run_sync_worker,
            args=(include_past, dry_run, replace_existing),
            daemon=True,
        ).start()

    def _run_sync_worker(self, include_past: bool, dry_run: bool, replace_existing: bool):
        writer = QueueWriter(self.log_queue)
        try:
            from playwright.sync_api import sync_playwright

            with redirect_stdout(writer), redirect_stderr(writer):
                storage_state = self.context.storage_state()

                with sync_playwright() as p:
                    api_context = p.request.new_context(storage_state=storage_state)

                    class RequestContextShim:
                        def __init__(self, request):
                            self.request = request

                    shim = RequestContextShim(api_context)
                    assignments_by_class = self.backend.fetch_assignments_from_canvas_context(
                        shim,
                        self.sheet_patterns,
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
        finally:
            writer.flush()
            self.sync_running = False

    def _set_status(self, value: str):
        self.after(0, lambda: self.status_var.set(value))

    def _set_login_hint(self, value: str):
        self.after(0, lambda: self.login_hint_var.set(value))

    def _close_app(self):
        try:
            if self.browser is not None:
                self.browser.close()
        except Exception:
            pass

        try:
            if self.playwright_manager is not None:
                self.playwright_manager.stop()
        except Exception:
            pass

        self.destroy()


def main():
    app = AssignmentTrackerGUI()
    app.protocol("WM_DELETE_WINDOW", app._close_app)
    app.mainloop()


if __name__ == "__main__":
    main()
