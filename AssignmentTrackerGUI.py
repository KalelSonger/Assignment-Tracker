import importlib.util
import base64
import json
import os
import queue
import re
import socket
import subprocess
import sys
import threading
import time
import traceback
import urllib.parse
import webbrowser
import ctypes
from contextlib import redirect_stderr, redirect_stdout

try:
    import winreg
except Exception:
    winreg = None

import tkinter as tk
from tkinter import messagebox, ttk

from keys import CONFIG_SOURCE

try:
    import resvg_py
except Exception:
    resvg_py = None


SINGLE_INSTANCE_HOST = "127.0.0.1"
SINGLE_INSTANCE_PORT = 48523
CANVAS_SESSION_FILE = "canvas_session.local.json"
SHEET_ENDPOINTS_FILE = "sheet_endpoints.local.json"
APP_SETTINGS_FILE = "app_settings.local.json"
SHEET_URL_PLACEHOLDER = "add sheet url here"
DEFAULT_APP_SETTINGS = {
    "auto_sync_on_startup": False,
    "run_on_windows_startup": False,
    "theme": "system",
}


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
        self._skip_next_auto_sync = False
        self.top_controls_frame = None
        self.top_controls_notice_var = tk.StringVar(value="")
        self.sync_running = False
        self.backend = None
        self.app_settings = dict(DEFAULT_APP_SETTINGS)
        self.settings_auto_sync_var = tk.BooleanVar(value=False)
        self.settings_startup_app_var = tk.BooleanVar(value=False)
        self.settings_theme_var = tk.StringVar(value="light")
        self.theme_palette = {
            "bg": "#f2f2f2",
            "button_bg": "#e8e8e8",
            "accent": "#1f6fff",
            "normal_input_fg": "#000000",
            "placeholder_fg": "#777777",
            "muted_fg": "#444444",
        }
        self.settings_button = None
        self.settings_button_icon = None
        self.settings_tooltip = None
        self.generate_sheet_tooltip = None
        self.remove_sheet_tooltip = None
        self.add_sheet_button = None
        self.remove_sheet_button = None
        self.settings_window = None
        self.settings_theme_combo = None
        self.settings_theme_label_var = tk.StringVar(value="Light theme")

        self.state_dir = self._state_dir()
        self.canvas_session_path = os.path.join(self.state_dir, CANVAS_SESSION_FILE)
        self.sheet_endpoints_path = os.path.join(self.state_dir, SHEET_ENDPOINTS_FILE)
        self.app_settings_path = os.path.join(self.state_dir, APP_SETTINGS_FILE)
        self._ensure_local_state_files()
        self._load_app_settings()

        self._build_ui()
        self._apply_theme()
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

    def _open_settings_window(self):
        if self.settings_window is not None and self.settings_window.winfo_exists():
            self.settings_window.deiconify()
            self.settings_window.lift()
            self.settings_window.focus_force()
            return

        self.settings_window = tk.Toplevel(self)
        self.settings_window.title("Settings")
        self.settings_window.geometry("420x320")
        self.settings_window.resizable(False, False)
        self.settings_window.transient(self)
        self.settings_window.protocol("WM_DELETE_WINDOW", self._close_settings_window)

        container = ttk.Frame(self.settings_window, padding=14)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)

        ttk.Label(container, text="Startup", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 8)
        )

        ttk.Checkbutton(
            container,
            text="Auto sync on startup",
            variable=self.settings_auto_sync_var,
            command=self._on_toggle_auto_sync,
            style="Settings.TCheckbutton",
        ).grid(row=1, column=0, columnspan=2, sticky="w")

        ttk.Checkbutton(
            container,
            text="Run as Windows startup app",
            variable=self.settings_startup_app_var,
            command=self._on_toggle_windows_startup,
            style="Settings.TCheckbutton",
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 0))

        ttk.Separator(container, orient="horizontal").grid(row=3, column=0, columnspan=2, sticky="ew", pady=12)

        ttk.Label(container, text="Appearance", font=("Segoe UI", 10, "bold")).grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(0, 8)
        )

        ttk.Label(container, text="Program theme:").grid(row=5, column=0, sticky="w", padx=(0, 8))
        self.settings_theme_label_var.set(self._theme_to_label(self.settings_theme_var.get()))
        self.settings_theme_combo = ttk.Combobox(
            container,
            textvariable=self.settings_theme_label_var,
            values=("System default", "Light theme", "Dark theme", "Coral theme", "Spotify theme"),
            state="readonly",
            width=20,
        )
        self.settings_theme_combo.grid(row=5, column=1, sticky="w")
        self.settings_theme_combo.bind("<<ComboboxSelected>>", self._on_theme_combo_selected)

        ttk.Separator(container, orient="horizontal").grid(row=6, column=0, columnspan=2, sticky="ew", pady=12)

        ttk.Label(container, text="Data", font=("Segoe UI", 10, "bold")).grid(
            row=7, column=0, columnspan=2, sticky="w", pady=(0, 8)
        )

        ttk.Button(
            container,
            text="Remove saved sheet URLs",
            command=self._settings_clear_sheet_urls,
            width=24,
        ).grid(row=8, column=0, columnspan=2, sticky="w")

        ttk.Button(container, text="Close", command=self._close_settings_window, width=12).grid(
            row=9, column=1, sticky="e", pady=(18, 0)
        )

        self._apply_theme()
        self.settings_window.lift()
        self.settings_window.focus_force()

    def _close_settings_window(self):
        if self.settings_window is not None and self.settings_window.winfo_exists():
            self.settings_window.destroy()
        self.settings_window = None
        self.settings_theme_combo = None

    def _settings_clear_sheet_urls(self):
        confirmed = messagebox.askyesno(
            "Remove saved sheet URLs",
            "Are you sure you want to remove all saved sheet URLs?\n"
            "You will need to add a sheet URL again before syncing.",
            parent=self.settings_window,
        )
        if not confirmed:
            return
        self._close_settings_window()
        self.sheet_registry = {"selected_api_url": "", "sheets": []}
        self._save_sheet_registry()
        self.allowed_tabs = []
        self.sheet_patterns = []
        self.awaiting_initial_sheet_url = True
        self._log("All saved sheet URLs removed. Add a sheet URL to continue.")
        self.after(0, lambda: self._show_sync_panel())

    def _theme_to_label(self, theme_name: str) -> str:
        normalized = str(theme_name or "").strip().lower()
        if normalized == "system":
            return "System default"
        if normalized == "spotify":
            return "Spotify theme"
        if normalized == "coral":
            return "Coral theme"
        if normalized == "dark":
            return "Dark theme"
        return "Light theme"

    def _label_to_theme(self, label: str) -> str:
        normalized = str(label or "").strip().lower()
        if normalized.startswith("system"):
            return "system"
        if normalized.startswith("spotify"):
            return "spotify"
        if normalized.startswith("coral"):
            return "coral"
        if normalized.startswith("dark"):
            return "dark"
        return "light"

    def _detect_system_theme(self) -> str:
        if os.name != "nt":
            return "light"

        try:
            import winreg

            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
            ) as key:
                apps_use_light, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
                return "light" if int(apps_use_light) == 1 else "dark"
        except Exception:
            return "light"

    def _on_theme_combo_selected(self, _event=None):
        selected_theme = self._label_to_theme(self.settings_theme_label_var.get())
        self.settings_theme_var.set(selected_theme)
        self._on_theme_selected()

    def _settings_svg_path(self) -> str | None:
        candidates = []

        if hasattr(sys, "_MEIPASS"):
            candidates.append(os.path.join(sys._MEIPASS, "settings.svg"))

        script_dir = os.path.dirname(os.path.abspath(__file__))
        candidates.extend(
            [
                os.path.join(script_dir, "settings.svg"),
                os.path.join(os.getcwd(), "settings.svg"),
            ]
        )

        for candidate in candidates:
            if os.path.isfile(candidate):
                return candidate

        return None

    def _render_settings_svg_icon(self, accent_hex: str, icon_size: int = 18):
        global resvg_py

        svg_path = self._settings_svg_path()
        if not svg_path:
            return None

        try:
            with open(svg_path, "r", encoding="utf-8") as file:
                svg_text = file.read()
        except Exception:
            return None

        # Recolor common black values to the active accent color for theme-sync icon tinting.
        recolored_svg = re.sub(
            r"(#000000|#000\b|black\b|rgb\(0\s*,\s*0\s*,\s*0\))",
            accent_hex,
            svg_text,
            flags=re.IGNORECASE,
        )

        try:
            if resvg_py is None:
                resvg_py = importlib.import_module("resvg_py")

            png_bytes = resvg_py.svg_to_bytes(
                svg_string=recolored_svg,
                width=icon_size,
                height=icon_size,
            )
            encoded = base64.b64encode(png_bytes).decode("ascii")
            return tk.PhotoImage(data=encoded)
        except Exception:
            return None

    def _update_settings_button_visual(self, hover: bool = False):
        if self.settings_button is None:
            return

        palette = self.theme_palette
        accent = palette.get("accent", "#1f6fff")
        button_bg = palette.get("button_bg", "#e8e8e8")
        icon_size = 16
        if self.add_sheet_button is not None and self.add_sheet_button.winfo_exists():
            base_height = max(self.add_sheet_button.winfo_height(), self.add_sheet_button.winfo_reqheight())
            if base_height:
                icon_size = max(14, min(22, int(base_height) - 4))

        icon = self._render_settings_svg_icon("#ffffff" if hover else accent, icon_size=icon_size)
        self.settings_button_icon = icon

        style = ttk.Style(self)
        style.configure(
            "SettingsIcon.TButton",
            background=button_bg,
            foreground=accent,
            bordercolor=accent,
            lightcolor=accent,
            darkcolor=accent,
            padding=(2, 2),
        )
        style.map(
            "SettingsIcon.TButton",
            background=[("active", accent), ("pressed", accent)],
            foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
        )

        if icon is not None:
            self.settings_button.configure(image=icon, text="", compound="center")
        else:
            self.settings_button.configure(image="", text="⚙")

    def _on_settings_button_enter(self, _event=None):
        self._update_settings_button_visual(hover=True)
        if _event is not None:
            self._show_settings_tooltip(_event.x_root, _event.y_root)

    def _on_generate_button_enter(self, _event=None):
        if _event is not None:
            self._show_generate_sheet_tooltip(_event.x_root, _event.y_root)

    def _on_generate_button_leave(self, _event=None):
        self._hide_generate_sheet_tooltip()

    def _on_generate_button_motion(self, _event=None):
        if _event is not None:
            self._show_generate_sheet_tooltip(_event.x_root, _event.y_root)

    def _on_settings_button_leave(self, _event=None):
        self._update_settings_button_visual(hover=False)
        self._hide_settings_tooltip()

    def _on_settings_button_motion(self, _event=None):
        if _event is not None:
            self._show_settings_tooltip(_event.x_root, _event.y_root)

    def _show_settings_tooltip(self, root_x: int, root_y: int):
        if self.settings_tooltip is None or not self.settings_tooltip.winfo_exists():
            self.settings_tooltip = tk.Toplevel(self)
            self.settings_tooltip.withdraw()
            self.settings_tooltip.overrideredirect(True)
            self.settings_tooltip.attributes("-topmost", True)
            tk.Label(
                self.settings_tooltip,
                text="Settings",
                bg="#111111",
                fg="#f5f5f5",
                padx=6,
                pady=3,
                font=("Segoe UI", 9),
            ).pack()

        self.settings_tooltip.geometry(f"+{root_x + 12}+{root_y + 18}")
        self.settings_tooltip.deiconify()

    def _hide_settings_tooltip(self):
        if self.settings_tooltip is not None and self.settings_tooltip.winfo_exists():
            self.settings_tooltip.withdraw()

    def _show_generate_sheet_tooltip(self, root_x: int, root_y: int):
        if self.generate_sheet_tooltip is None or not self.generate_sheet_tooltip.winfo_exists():
            self.generate_sheet_tooltip = tk.Toplevel(self)
            self.generate_sheet_tooltip.withdraw()
            self.generate_sheet_tooltip.overrideredirect(True)
            self.generate_sheet_tooltip.attributes("-topmost", True)
            tk.Label(
                self.generate_sheet_tooltip,
                text="generate blank spreadsheet from template.",
                bg="#111111",
                fg="#f5f5f5",
                padx=6,
                pady=3,
                font=("Segoe UI", 9),
            ).pack()

        self.generate_sheet_tooltip.geometry(f"+{root_x + 12}+{root_y + 18}")
        self.generate_sheet_tooltip.deiconify()

    def _hide_generate_sheet_tooltip(self):
        if self.generate_sheet_tooltip is not None and self.generate_sheet_tooltip.winfo_exists():
            self.generate_sheet_tooltip.withdraw()

    def _on_remove_sheet_button_enter(self, _event=None):
        if _event is not None:
            self._show_remove_sheet_tooltip(_event.x_root, _event.y_root)

    def _on_remove_sheet_button_leave(self, _event=None):
        self._hide_remove_sheet_tooltip()

    def _on_remove_sheet_button_motion(self, _event=None):
        if _event is not None:
            self._show_remove_sheet_tooltip(_event.x_root, _event.y_root)

    def _show_remove_sheet_tooltip(self, root_x: int, root_y: int):
        if self.remove_sheet_tooltip is None or not self.remove_sheet_tooltip.winfo_exists():
            self.remove_sheet_tooltip = tk.Toplevel(self)
            self.remove_sheet_tooltip.withdraw()
            self.remove_sheet_tooltip.overrideredirect(True)
            self.remove_sheet_tooltip.attributes("-topmost", True)
            tk.Label(
                self.remove_sheet_tooltip,
                text="Remove spreadsheet",
                bg="#111111",
                fg="#f5f5f5",
                padx=6,
                pady=3,
                font=("Segoe UI", 9),
            ).pack()

        self.remove_sheet_tooltip.geometry(f"+{root_x + 12}+{root_y + 18}")
        self.remove_sheet_tooltip.deiconify()

    def _hide_remove_sheet_tooltip(self):
        if self.remove_sheet_tooltip is not None and self.remove_sheet_tooltip.winfo_exists():
            self.remove_sheet_tooltip.withdraw()

    def _hex_to_colorref(self, color_hex: str) -> int:
        raw = str(color_hex or "").strip().lstrip("#")
        if len(raw) != 6:
            return 0
        red = int(raw[0:2], 16)
        green = int(raw[2:4], 16)
        blue = int(raw[4:6], 16)
        return (blue << 16) | (green << 8) | red

    def _apply_windows_title_bar_theme(self, window: tk.Tk | tk.Toplevel, palette: dict[str, str], theme_name: str):
        if os.name != "nt":
            return

        try:
            window.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
            if not hwnd:
                return

            use_dark_mode = ctypes.c_int(1 if theme_name == "dark" else 0)
            caption_color = ctypes.c_int(self._hex_to_colorref(palette["bg"]))
            text_color_hex = palette["accent"]
            text_color = ctypes.c_int(self._hex_to_colorref(text_color_hex))

            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            DWMWA_CAPTION_COLOR = 35
            DWMWA_TEXT_COLOR = 36

            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd,
                DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(use_dark_mode),
                ctypes.sizeof(use_dark_mode),
            )
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd,
                DWMWA_CAPTION_COLOR,
                ctypes.byref(caption_color),
                ctypes.sizeof(caption_color),
            )
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd,
                DWMWA_TEXT_COLOR,
                ctypes.byref(text_color),
                ctypes.sizeof(text_color),
            )
        except Exception:
            pass

    def _light_palette(self) -> dict[str, str]:
        return {
            "bg": "#f2f2f2",
            "surface": "#ffffff",
            "fg": "#1a1a1a",
            "muted_fg": "#444444",
            "normal_input_fg": "#000000",
            "placeholder_fg": "#777777",
            "accent": "#1f6fff",
            "text_bg": "#ffffff",
            "text_fg": "#111111",
            "button_fg": "#1a1a1a",
            "button_bg": "#e8e8e8",
            "check_hover_bg": "#e2e2e2",
            "check_hover_fg": "#1a1a1a",
        }

    def _dark_palette(self) -> dict[str, str]:
        return {
            "bg": "#101010",
            "surface": "#1a1a1a",
            "fg": "#f5f5f5",
            "muted_fg": "#dddddd",
            "normal_input_fg": "#ffffff",
            "placeholder_fg": "#9b9b9b",
            "accent": "#ff9c1a",
            "text_bg": "#111111",
            "text_fg": "#f3f3f3",
            "button_fg": "#ffffff",
            "button_bg": "#2a2a2a",
            "check_hover_bg": "#2f2f2f",
            "check_hover_fg": "#ffffff",
        }

    def _spotify_palette(self) -> dict[str, str]:
        palette = self._dark_palette().copy()
        palette["bg"] = "#121212"
        palette["text_bg"] = "#121212"
        palette["accent"] = "#1ED760"
        return palette

    def _coral_palette(self) -> dict[str, str]:
        return {
            "bg": "#142E4C",
            "surface": "#142E4C",
            "fg": "#FFFFFF",
            "muted_fg": "#D5E2F0",
            "normal_input_fg": "#FFFFFF",
            "placeholder_fg": "#B8C7D8",
            "accent": "#FF9B82",
            "text_bg": "#142E4C",
            "text_fg": "#FFFFFF",
            "button_fg": "#FFFFFF",
            "button_bg": "#142E4C",
            "check_hover_bg": "#1B3A5E",
            "check_hover_fg": "#FFFFFF",
        }

    def _apply_theme(self):
        requested = str(self.settings_theme_var.get() or "system").lower()
        if requested not in ("system", "light", "dark", "coral", "spotify"):
            requested = "system"

        active_theme = self._detect_system_theme() if requested == "system" else requested
        if active_theme == "dark":
            palette = self._dark_palette()
        elif active_theme == "spotify":
            palette = self._spotify_palette()
        elif active_theme == "coral":
            palette = self._coral_palette()
        else:
            palette = self._light_palette()

        self.theme_palette = palette

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", background=palette["bg"], foreground=palette["fg"])
        style.configure("TFrame", background=palette["bg"])
        style.configure("TLabel", background=palette["bg"], foreground=palette["fg"])
        style.configure("TSeparator", background=palette["accent"])
        style.configure(
            "TButton",
            background=palette["button_bg"],
            foreground=palette["button_fg"],
            bordercolor=palette["accent"],
            lightcolor=palette["accent"],
            darkcolor=palette["accent"],
        )
        style.map(
            "TButton",
            background=[("active", palette["accent"]), ("pressed", palette["accent"])],
            foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
        )
        style.configure(
            "TEntry",
            fieldbackground=palette["surface"],
            foreground=palette["normal_input_fg"],
        )
        style.configure(
            "TCombobox",
            fieldbackground=palette["surface"],
            background=palette["surface"],
            foreground=palette["normal_input_fg"],
            arrowcolor=palette["fg"],
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", palette["surface"])],
            foreground=[("readonly", palette["normal_input_fg"])],
            selectbackground=[("readonly", palette["accent"])],
            selectforeground=[("readonly", "#ffffff")],
        )
        style.configure(
            "Settings.TCheckbutton",
            background=palette["bg"],
            foreground=palette["fg"],
        )
        style.map(
            "Settings.TCheckbutton",
            background=[("active", palette.get("check_hover_bg", palette["bg"]))],
            foreground=[("active", palette.get("check_hover_fg", palette["fg"]))],
        )

        self.configure(bg=palette["bg"])

        if hasattr(self, "log_text") and self.log_text is not None:
            self.log_text.configure(
                background=palette["text_bg"],
                foreground=palette["text_fg"],
                insertbackground=palette["text_fg"],
                selectbackground=palette["accent"],
                selectforeground="#ffffff",
            )

        if hasattr(self, "sheet_url_entry") and self.sheet_url_entry is not None:
            if self.sheet_url_has_placeholder:
                self.sheet_url_entry.configure(foreground=palette["placeholder_fg"])
            else:
                self.sheet_url_entry.configure(foreground=palette["normal_input_fg"])

        self._update_settings_button_visual(hover=False)

        if self.settings_window is not None and self.settings_window.winfo_exists():
            self.settings_window.configure(bg=palette["bg"])

        self._apply_windows_title_bar_theme(self, palette, active_theme)
        if self.settings_window is not None and self.settings_window.winfo_exists():
            self._apply_windows_title_bar_theme(self.settings_window, palette, active_theme)

    def _load_app_settings(self):
        try:
            with open(self.app_settings_path, "r", encoding="utf-8") as file:
                parsed = json.load(file)
        except Exception:
            parsed = {}

        merged = dict(DEFAULT_APP_SETTINGS)
        if isinstance(parsed, dict):
            merged.update(parsed)

        configured_theme = str(merged.get("theme") or "system").lower()
        if configured_theme not in ("light", "dark", "system", "coral", "spotify"):
            configured_theme = "system"

        merged["theme"] = configured_theme
        merged["auto_sync_on_startup"] = bool(merged.get("auto_sync_on_startup"))
        merged["run_on_windows_startup"] = bool(merged.get("run_on_windows_startup"))
        self.app_settings = merged
        self.settings_auto_sync_var.set(self.app_settings["auto_sync_on_startup"])
        self.settings_startup_app_var.set(self.app_settings["run_on_windows_startup"])
        self._apply_windows_startup_preference(silent=True)
        actual_startup_enabled = self._is_windows_startup_enabled()
        if actual_startup_enabled != self.app_settings["run_on_windows_startup"]:
            self.app_settings["run_on_windows_startup"] = actual_startup_enabled
            self.settings_startup_app_var.set(actual_startup_enabled)
        self.settings_theme_var.set(configured_theme)
        self.settings_theme_label_var.set(self._theme_to_label(configured_theme))
        self._save_app_settings()

    def _save_app_settings(self):
        with open(self.app_settings_path, "w", encoding="utf-8") as file:
            json.dump(self.app_settings, file, indent=2)

    def _on_toggle_auto_sync(self):
        enabled = bool(self.settings_auto_sync_var.get())
        self.app_settings["auto_sync_on_startup"] = enabled
        self._save_app_settings()
        self._log(f"Auto sync on startup {'enabled' if enabled else 'disabled'}.")

    def _windows_startup_registry_path(self) -> str:
        return r"Software\Microsoft\Windows\CurrentVersion\Run"

    def _windows_startup_value_name(self) -> str:
        return "AssignmentTrackerGUI"

    def _windows_startup_command(self) -> str:
        if getattr(sys, "frozen", False):
            return f'"{sys.executable}"'

        interpreter = sys.executable
        if interpreter.lower().endswith("python.exe"):
            pythonw = os.path.join(os.path.dirname(interpreter), "pythonw.exe")
            if os.path.isfile(pythonw):
                interpreter = pythonw

        script_path = os.path.abspath(__file__)
        return f'"{interpreter}" "{script_path}"'

    def _is_windows_startup_enabled(self) -> bool:
        if os.name != "nt" or winreg is None:
            return False

        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self._windows_startup_registry_path(), 0, winreg.KEY_READ) as key:
                stored_command, _ = winreg.QueryValueEx(key, self._windows_startup_value_name())
                command_text = str(stored_command or "").strip()
                if not command_text:
                    return False

                executable_path = ""
                if command_text.startswith('"') and '"' in command_text[1:]:
                    executable_path = command_text.split('"', 2)[1].strip()
                else:
                    executable_path = command_text.split(" ", 1)[0].strip()

                if executable_path and not os.path.isfile(executable_path):
                    return False

                return True
        except OSError:
            return False

    def _set_windows_startup_enabled(self, enabled: bool) -> None:
        if os.name != "nt" or winreg is None:
            raise RuntimeError("Windows startup apps are only supported on Windows.")

        value_name = self._windows_startup_value_name()
        if enabled:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, self._windows_startup_registry_path()) as key:
                winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, self._windows_startup_command())
        else:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, self._windows_startup_registry_path()) as key:
                try:
                    winreg.DeleteValue(key, value_name)
                except FileNotFoundError:
                    pass

    def _apply_windows_startup_preference(self, silent: bool = False) -> None:
        enabled = bool(self.app_settings.get("run_on_windows_startup", False))
        if os.name != "nt" or winreg is None:
            return

        try:
            self._set_windows_startup_enabled(enabled)
        except Exception as error:
            if not silent:
                self._log(f"Windows startup setting error: {error}")

    def _on_toggle_windows_startup(self):
        enabled = bool(self.settings_startup_app_var.get())
        try:
            self._set_windows_startup_enabled(enabled)
            self.app_settings["run_on_windows_startup"] = enabled
            self._save_app_settings()
            self._log(f"Windows startup app {'enabled' if enabled else 'disabled'}.")
        except Exception as error:
            self.settings_startup_app_var.set(not enabled)
            messagebox.showerror("Startup setting failed", str(error))

    def _on_theme_selected(self):
        selected_theme = str(self.settings_theme_var.get() or "system").lower()
        if selected_theme not in ("system", "light", "dark", "coral", "spotify"):
            selected_theme = "system"
        self.app_settings["theme"] = selected_theme
        self.settings_theme_var.set(selected_theme)
        self.settings_theme_label_var.set(self._theme_to_label(selected_theme))
        self._save_app_settings()
        self._apply_theme()
        self._log(f"Theme changed to {selected_theme}.")

    def _maybe_start_auto_sync(self):
        if self._skip_next_auto_sync:
            self._skip_next_auto_sync = False
            self._log("Auto sync skipped for first-time sheet setup.")
            return
        if not self.app_settings.get("auto_sync_on_startup", False):
            return
        if self.sync_running:
            return
        if not self._selected_sheet_api_url().strip():
            self._log("Auto sync skipped: no sheet URL selected.")
            return
        self._log("Auto sync on startup is enabled. Running 'Sync all assignments'.")
        self._start_sync(include_past=True, dry_run=False, replace_existing=False)

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
                foreground=self.theme_palette.get("muted_fg", "#555555"),
            ).grid(row=0, column=0, columnspan=6, sticky="e", pady=(0, 2))
            controls_row = 1
        else:
            controls_row = 0

        # "Generate formatted blank spreadsheet" button
        self.generate_sheet_button = ttk.Button(
            self.top_controls_frame,
            text="Generate",
            command=self._generate_formatted_sheet,
            width=12,
        )
        self.generate_sheet_button.grid(row=controls_row, column=0, padx=(0, 8), sticky="e")
        self.generate_sheet_button.bind("<Enter>", self._on_generate_button_enter)
        self.generate_sheet_button.bind("<Leave>", self._on_generate_button_leave)
        self.generate_sheet_button.bind("<Motion>", self._on_generate_button_motion)

        dropdown_state = "readonly" if self.sheet_registry.get("sheets") else "disabled"
        self.sheet_dropdown = ttk.Combobox(
            self.top_controls_frame,
            textvariable=self.selected_sheet_name_var,
            state=dropdown_state,
            width=30,
        )
        self.sheet_dropdown.grid(row=controls_row, column=1, padx=(0, 8), sticky="e")
        self.sheet_dropdown.bind("<<ComboboxSelected>>", self._on_sheet_selected)

        self.remove_sheet_button = ttk.Button(
            self.top_controls_frame,
            text="x",
            command=self._remove_selected_sheet,
            width=3,
        )
        self.remove_sheet_button.grid(row=controls_row, column=2, padx=(0, 8), sticky="e")
        self.remove_sheet_button.bind("<Enter>", self._on_remove_sheet_button_enter)
        self.remove_sheet_button.bind("<Leave>", self._on_remove_sheet_button_leave)
        self.remove_sheet_button.bind("<Motion>", self._on_remove_sheet_button_motion)

        self.sheet_url_entry = ttk.Entry(self.top_controls_frame, textvariable=self.sheet_url_input_var, width=38)
        self.sheet_url_entry.grid(row=controls_row, column=3, padx=(0, 8), sticky="e")
        self.sheet_url_entry.bind("<FocusIn>", self._on_sheet_url_focus_in)
        self.sheet_url_entry.bind("<FocusOut>", self._on_sheet_url_focus_out)
        self._apply_sheet_url_placeholder()

        self.add_sheet_button = ttk.Button(
            self.top_controls_frame,
            text="Add",
            command=self._handle_add_sheet_from_login if self.awaiting_initial_sheet_url else self._add_sheet_endpoint,
            width=10,
        )
        self.add_sheet_button.grid(row=controls_row, column=4, sticky="e")

        self.settings_button = ttk.Button(
            self.top_controls_frame,
            text="",
            command=self._open_settings_window,
            style="SettingsIcon.TButton",
            width=2,
        )
        self.settings_button.grid(row=controls_row, column=5, sticky="e", padx=(8, 0))
        self.settings_button.bind("<Enter>", self._on_settings_button_enter)
        self.settings_button.bind("<Leave>", self._on_settings_button_leave)
        self.settings_button.bind("<Motion>", self._on_settings_button_motion)

        self._refresh_sheet_dropdown()
        self._update_settings_button_visual(hover=False)
        self.after(0, lambda: self._update_settings_button_visual(hover=False))

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
        if not os.path.isfile(self.app_settings_path):
            with open(self.app_settings_path, "w", encoding="utf-8") as file:
                json.dump(DEFAULT_APP_SETTINGS, file, indent=2)

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
        ttk.Label(
            self.left_panel,
            textvariable=self.login_hint_var,
            foreground=self.theme_palette.get("muted_fg", "#444444"),
        ).pack(anchor="w")

        self.reopen_login_button = ttk.Button(
            self.left_panel,
            text="Reopen browser",
            command=self._retry_login_browser,
            state="normal" if self.backend is not None else "disabled",
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

        scroll_container = ttk.Frame(self.left_panel)
        scroll_container.pack(fill="both", expand=True)

        scroll_canvas = tk.Canvas(
            scroll_container,
            highlightthickness=0,
            bd=0,
            relief="flat",
            background=self.theme_palette.get("bg", "#f2f2f2"),
        )
        scroll_bar = ttk.Scrollbar(scroll_container, orient="vertical", command=scroll_canvas.yview)
        scroll_canvas.configure(yscrollcommand=scroll_bar.set)

        scroll_canvas.pack(side="left", fill="both", expand=True)
        scroll_bar.pack(side="right", fill="y")

        scroll_body = ttk.Frame(scroll_canvas)
        body_window = scroll_canvas.create_window((0, 0), window=scroll_body, anchor="nw")

        def _on_body_configure(_event=None):
            scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))

        def _on_canvas_configure(event):
            scroll_canvas.itemconfigure(body_window, width=event.width)

        def _on_mousewheel(event):
            delta = int(getattr(event, "delta", 0))
            if delta == 0:
                return "break"
            scroll_canvas.yview_scroll(int(-1 * (delta / 120)), "units")
            return "break"

        def _bind_mousewheel(_event=None):
            scroll_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_mousewheel(_event=None):
            scroll_canvas.unbind_all("<MouseWheel>")

        scroll_body.bind("<Configure>", _on_body_configure)
        scroll_canvas.bind("<Configure>", _on_canvas_configure)
        scroll_canvas.bind("<Enter>", _bind_mousewheel)
        scroll_canvas.bind("<Leave>", _unbind_mousewheel)
        scroll_body.bind("<Enter>", _bind_mousewheel)
        scroll_body.bind("<Leave>", _unbind_mousewheel)

        ttk.Separator(scroll_body, orient="horizontal").pack(fill="x", pady=10)

        ttk.Label(scroll_body, text="Sync Actions", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 8))

        ttk.Button(
            scroll_body,
            text="Sync all assignments",
            command=lambda: self._start_sync(include_past=True, dry_run=False, replace_existing=False),
            width=button_width,
        ).pack(anchor="w", pady=4)

        ttk.Button(
            scroll_body,
            text="Sync future assignments",
            command=lambda: self._start_sync(include_past=False, dry_run=False, replace_existing=False),
            width=button_width,
        ).pack(anchor="w", pady=4)

        ttk.Button(
            scroll_body,
            text="Dry-sync (no writes)",
            command=lambda: self._start_sync(include_past=True, dry_run=True, replace_existing=False),
            width=button_width,
        ).pack(anchor="w", pady=4)

        ttk.Label(scroll_body, text="Sync individual class tab:").pack(anchor="w", pady=(8, 2))
        for class_tab in self.allowed_tabs:
            ttk.Button(
                scroll_body,
                text=f"Sync: {class_tab}",
                command=lambda tab_name=class_tab: self._start_sync_single_tab(tab_name),
                width=button_width,
            ).pack(anchor="w", pady=2)

        ttk.Separator(scroll_body, orient="horizontal").pack(fill="x", pady=10)
        ttk.Label(scroll_body, text="Clear Actions", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 6))

        ttk.Button(
            scroll_body,
            text="Clear all class tabs",
            command=self._start_clear_all_tabs,
            width=button_width,
        ).pack(anchor="w", pady=4)

        ttk.Label(scroll_body, text="Clear individual class tab:").pack(anchor="w", pady=(8, 2))
        for class_tab in self.allowed_tabs:
            ttk.Button(
                scroll_body,
                text=f"Clear: {class_tab}",
                command=lambda tab_name=class_tab: self._start_clear_single_tab(tab_name),
                width=button_width,
            ).pack(anchor="w", pady=2)

        ttk.Separator(scroll_body, orient="horizontal").pack(fill="x", pady=10)
        ttk.Button(scroll_body, text="Close", command=self._close_app, width=button_width).pack(anchor="w", pady=4)

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
        try:
            if self.backend is not None and hasattr(self.backend, "infer_sheet_display_name"):
                inferred = self.backend.infer_sheet_display_name(api_url)
                if isinstance(inferred, str) and inferred.strip():
                    return inferred.strip()
        except Exception:
            pass

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
            self.sheet_dropdown.configure(state="readonly" if names else "disabled")

        if self.remove_sheet_button is not None:
            self.remove_sheet_button.configure(state="normal" if names else "disabled")

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

        sheet_id_match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", cleaned)
        if sheet_id_match:
            return f"sheet://{sheet_id_match.group(1)}"

        parsed = urllib.parse.urlparse(cleaned)
        scheme = (parsed.scheme or "https").lower()
        netloc = parsed.netloc.lower()
        path = parsed.path.rstrip("/")
        return urllib.parse.urlunparse((scheme, netloc, path, "", "", ""))

    def _add_sheet_endpoint(self, reload_tabs: bool = True):
        raw_url = self.sheet_url_input_var.get().strip()
        if self.sheet_url_has_placeholder and raw_url == SHEET_URL_PLACEHOLDER:
            raw_url = ""
        self._register_sheet_endpoint(raw_url, reload_tabs=reload_tabs)
        self.sheet_url_input_var.set("")
        self.sheet_url_has_placeholder = False
        self._apply_sheet_url_placeholder()

    def _register_sheet_endpoint(self, raw_url: str, reload_tabs: bool = True, preferred_name: str | None = None):
        if not raw_url:
            messagebox.showwarning("Missing URL", "Paste a Google Sheet URL first.")
            return

        if not raw_url.startswith("http://") and not raw_url.startswith("https://"):
            messagebox.showwarning("Invalid URL", "Google Sheet URL must start with http:// or https://")
            return

        if "/spreadsheets/d/" not in raw_url:
            messagebox.showwarning(
                "Invalid URL",
                "Paste a full Google Sheet URL, like https://docs.google.com/spreadsheets/d/<ID>/edit",
            )
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
            if not self._ensure_google_sheet_access_or_prompt_reauth(raw_url):
                return
            name = str(preferred_name or "").strip() or self._safe_infer_sheet_name(raw_url)
            self.sheet_registry.setdefault("sheets", []).append({"api_url": raw_url, "display_name": name})
            self._log(f"Added sheet endpoint: {name}")
        else:
            self._log("Sheet endpoint already exists; selecting it.")
            raw_url = str(existing.get("api_url") or raw_url)

        self.sheet_registry["selected_api_url"] = raw_url
        self._save_sheet_registry()
        self._refresh_sheet_dropdown()
        if reload_tabs:
            self._reload_selected_sheet_tabs()

    def _ensure_google_sheet_access_or_prompt_reauth(self, api_url: str) -> bool:
        if self.backend is None or not hasattr(self.backend, "validate_google_sheet_access"):
            return True

        try:
            self.backend.validate_google_sheet_access(api_url)
            return True
        except Exception as error:
            raw_message = str(error)
            message = raw_message.casefold()
            reauth_keywords = (
                "permission",
                "forbidden",
                "insufficient",
                "requested entity was not found",
                "not found",
                "caller does not have permission",
            )
            should_prompt_reauth = any(keyword in message for keyword in reauth_keywords)

            if should_prompt_reauth and hasattr(self.backend, "reset_google_login"):
                wants_reauth = messagebox.askyesno(
                    "Google access required",
                    "That sheet may belong to a different Google account or is not shared with this account.\n\n"
                    "Sign in to Google again now and retry adding this sheet?",
                )
                if wants_reauth:
                    try:
                        self.backend.reset_google_login()
                        self.backend.validate_google_sheet_access(api_url)
                        self._log("Google sign-in refreshed for sheet access.")
                        return True
                    except Exception as retry_error:
                        messagebox.showerror(
                            "Google sign-in required",
                            f"Could not access that sheet after re-sign in:\n\n{retry_error}",
                        )
                        return False

            messagebox.showerror(
                "Sheet access failed",
                f"Could not access that Google Sheet with the current Google sign-in:\n\n{raw_message}",
            )
            return False

    def _remove_selected_sheet(self):
        selected_url = self._selected_sheet_api_url()
        if not selected_url:
            return

        selected_name = self.selected_sheet_name_var.get().strip() or "selected sheet"
        if not messagebox.askyesno(
            "Remove saved sheet",
            f"Remove '{selected_name}' from the saved sheet list?",
        ):
            return

        normalized_selected = self._normalize_api_url(selected_url)
        remaining = [
            item
            for item in self.sheet_registry.get("sheets", [])
            if self._normalize_api_url(item.get("api_url", "")) != normalized_selected
        ]

        self.sheet_registry["sheets"] = remaining
        self.sheet_registry["selected_api_url"] = remaining[0]["api_url"] if remaining else ""
        self._save_sheet_registry()
        self._refresh_sheet_dropdown()

        if remaining:
            self._log(f"Removed sheet endpoint: {selected_name}")
            self._reload_selected_sheet_tabs()
        else:
            self.allowed_tabs = []
            self.sheet_patterns = []
            self.awaiting_initial_sheet_url = True
            if self.storage_state is not None:
                self._set_status("Signed in. Sheet URL required.")
                self._set_login_hint("Add a sheet URL at the top to continue.")
            else:
                self._set_status("Sheet URL required")
                self._set_login_hint("Add a sheet URL at the top to continue.")
            self._log("No sheet endpoints remain. Add a sheet URL to continue.")
            self.after(0, self._show_sync_panel)

    def _handle_add_sheet_from_login(self):
        self._add_sheet_endpoint(reload_tabs=False)
        if not self.sheet_registry.get("sheets"):
            return

        if self.awaiting_initial_sheet_url:
            self.awaiting_initial_sheet_url = False
            self._skip_next_auto_sync = True
            self._set_status("Sheet saved. Continuing startup...")
            self._set_login_hint("Initializing with selected sheet...")
            threading.Thread(target=self._bootstrap_and_start_login, daemon=True).start()

    def _generate_formatted_sheet(self):
        """Generate a formatted sheet from template with Canvas courses."""
        if self.backend is None or self.storage_state is None:
            messagebox.showwarning(
                "Not Ready",
                "Please complete Canvas sign-in first.",
            )
            return

        def run_generation():
            try:
                self._set_status("Generating formatted sheet...")
                self._log("Starting sheet generation...")
                self._log("This may take 30-60 seconds. Please wait...")

                from playwright.sync_api import sync_playwright

                with sync_playwright() as p:
                    api_context = p.request.new_context(storage_state=self.storage_state)
                    try:
                        auth_status = self._canvas_auth_status(api_context)
                        if auth_status != "authenticated":
                            if auth_status == "unauthenticated":
                                self._clear_canvas_session()
                                raise RuntimeError(
                                    "Canvas session expired. Please sign in again before generating a sheet."
                                )
                            raise RuntimeError(
                                "Could not verify Canvas session (network unavailable). Check your connection and retry."
                            )

                        class RequestContextShim:
                            def __init__(self, request):
                                self.request = request

                        shim = RequestContextShim(api_context)
                        new_sheet_url = self.backend.generate_formatted_sheet_from_template(shim)
                    finally:
                        api_context.dispose()

                generated_name = self.backend.infer_sheet_display_name(new_sheet_url)
                self.after(
                    0,
                    lambda: self._register_sheet_endpoint(
                        new_sheet_url,
                        reload_tabs=True,
                        preferred_name=generated_name,
                    ),
                )

                opened = webbrowser.open_new_tab(new_sheet_url)
                self._log(f"Sheet generated successfully!")
                self._log(f"URL: {new_sheet_url}")
                if opened:
                    self._log("Opened generated sheet in your default browser.")
                else:
                    self._log("Could not auto-open browser tab. Open the URL above manually.")
                
                self._set_status("Sheet generated. Ready to sync.")
                
            except Exception as error:
                error_msg = str(error)
                self._log(f"Sheet generation failed: {error_msg}")
                self._set_status("Sheet generation failed.")
                self.after(
                    0,
                    lambda: messagebox.showerror(
                        "Generation Failed",
                        f"Could not generate sheet:\n\n{error_msg}"
                    )
                )

        threading.Thread(target=run_generation, daemon=True).start()

    def _apply_sheet_url_placeholder(self):
        if self.sheet_url_entry is None:
            return
        if self.sheet_url_input_var.get().strip():
            return
        self.sheet_url_has_placeholder = True
        self.sheet_url_input_var.set(SHEET_URL_PLACEHOLDER)
        self.sheet_url_entry.configure(foreground=self.theme_palette.get("placeholder_fg", "#777777"))

    def _on_sheet_url_focus_in(self, _event=None):
        if not self.sheet_url_has_placeholder:
            return
        self.sheet_url_has_placeholder = False
        self.sheet_url_input_var.set("")
        self.sheet_url_entry.configure(foreground=self.theme_palette.get("normal_input_fg", "#000000"))

    def _on_sheet_url_focus_out(self, _event=None):
        if self.sheet_url_input_var.get().strip():
            self.sheet_url_entry.configure(foreground=self.theme_palette.get("normal_input_fg", "#000000"))
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

    def _canvas_auth_status(self, api_context) -> str:
        if self.backend is None:
            return "unreachable"

        if hasattr(self.backend, "get_canvas_auth_status"):
            try:
                status = str(self.backend.get_canvas_auth_status(api_context) or "").strip().lower()
            except Exception:
                status = ""
            if status in ("authenticated", "unauthenticated", "unreachable"):
                return status

        try:
            return "authenticated" if self.backend._is_canvas_authenticated(api_context) else "unauthenticated"
        except Exception:
            return "unreachable"

    def _storage_state_auth_status(self, storage_state: dict) -> str:
        from playwright.sync_api import sync_playwright

        with sync_playwright() as p:
            api_context = p.request.new_context(storage_state=storage_state)
            try:
                return self._canvas_auth_status(api_context)
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

    def _apply_selected_sheet_endpoint(self) -> bool:
        if self.backend is None:
            self._log("Backend is not loaded.")
            return False

        api_url = self._selected_sheet_api_url()
        if not api_url:
            self._log("No sheet endpoint selected.")
            return False

        try:
            self.backend.set_sheet_api_url(api_url)
            if self.sheet_registry.get("selected_api_url") != api_url:
                self.sheet_registry["selected_api_url"] = api_url
                self._save_sheet_registry()
            return True
        except Exception as error:
            self._log(f"Could not set selected sheet endpoint: {error}")
            self._log(traceback.format_exc())
            return False

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
            self.awaiting_initial_sheet_url = not selected_api_url.strip()

            if self.awaiting_initial_sheet_url:
                self.allowed_tabs = []
                self.sheet_patterns = []
                self._log("No saved sheet endpoints found. Canvas sign-in will still open.")
            else:
                try:
                    self.backend.set_sheet_api_url(selected_api_url)
                except Exception:
                    self.sheet_registry["selected_api_url"] = ""
                    self._save_sheet_registry()
                    self.awaiting_initial_sheet_url = True
                    self.allowed_tabs = []
                    self.sheet_patterns = []
                    self._log("Saved sheet URL is invalid. Canvas sign-in will still open; add a new sheet URL after login.")
                else:
                    self._log(f"Config source: {CONFIG_SOURCE}")
                    self._log(f"Active sheet API URL: {self.backend.get_sheet_api_url()}")
                    self._log("Loading selected sheet class tabs...")
                    allowed_tabs = self.backend.fetch_allowed_sheet_classes()
                    self.allowed_tabs = allowed_tabs
                    self.sheet_patterns = self.backend._build_sheet_class_patterns(allowed_tabs)
                    self._log(f"Loaded {len(allowed_tabs)} class tabs from selected sheet.")

            self._set_status("Checking saved Canvas session...")
            saved_state = self._load_canvas_session_from_disk()
            if saved_state:
                saved_status = self._storage_state_auth_status(saved_state)
                if saved_status in ("authenticated", "unreachable"):
                    self.storage_state = saved_state
                    if self.awaiting_initial_sheet_url:
                        self._set_status("Signed in. Sheet URL required.")
                        hint = "Saved Canvas session restored. Add a sheet URL at the top to continue."
                        if saved_status == "unreachable":
                            hint = "Loaded saved Canvas session (network check unavailable). Add a sheet URL to continue."
                        self._set_login_hint(hint)
                        self._log("Using saved Canvas session. Add a sheet URL to continue.")
                    else:
                        self._set_status("Signed in. Ready to sync.")
                        hint = "Saved Canvas session restored."
                        if saved_status == "unreachable":
                            hint = "Loaded saved Canvas session (network check unavailable)."
                        self._set_login_hint(hint)
                        self._log("Using saved Canvas session. No login needed.")

                    if saved_status == "unreachable":
                        self._log("Canvas auth could not be verified (network unavailable). Keeping saved session.")

                    self.after(0, self._show_sync_panel)
                    self.after(250, self._maybe_start_auto_sync)
                    return

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

        if importlib.util.find_spec("resvg_py") is None:
            self._log("Installing resvg-py package for themed SVG icon rendering...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "resvg-py"])
            except Exception:
                self._log("Could not install resvg-py; settings button will use fallback gear text.")

        if importlib.util.find_spec("googleapiclient") is None or importlib.util.find_spec("google_auth_oauthlib") is None:
            self._log("Installing Google Sheets dependencies...")
            subprocess.check_call(
                [
                    sys.executable,
                    "-m",
                    "pip",
                    "install",
                    "google-api-python-client",
                    "google-auth-oauthlib",
                ]
            )

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
                self.after(250, self._maybe_start_auto_sync)
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

        if not self._apply_selected_sheet_endpoint():
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

        if not self._apply_selected_sheet_endpoint():
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

        if not self._apply_selected_sheet_endpoint():
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

        if not self._apply_selected_sheet_endpoint():
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

                print(f"Using endpoint: {self.backend.get_sheet_api_url()}")

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

                    auth_status = self._canvas_auth_status(api_context)
                    if auth_status != "authenticated":
                        if auth_status == "unauthenticated":
                            self._clear_canvas_session()
                            raise RuntimeError(
                                "Canvas session expired. Click 'Reopen browser' from the sign-in panel to login again."
                            )
                        raise RuntimeError(
                            "Could not verify Canvas session (network unavailable). Check your connection and retry."
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
        self._hide_generate_sheet_tooltip()
        self._hide_settings_tooltip()
        self._hide_remove_sheet_tooltip()
        if self.generate_sheet_tooltip is not None and self.generate_sheet_tooltip.winfo_exists():
            self.generate_sheet_tooltip.destroy()
            self.generate_sheet_tooltip = None
        if self.settings_tooltip is not None and self.settings_tooltip.winfo_exists():
            self.settings_tooltip.destroy()
            self.settings_tooltip = None
        if self.remove_sheet_tooltip is not None and self.remove_sheet_tooltip.winfo_exists():
            self.remove_sheet_tooltip.destroy()
            self.remove_sheet_tooltip = None

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
