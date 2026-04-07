import ctypes
import csv
import io
import json
import os
import shutil
import subprocess
import sys
import threading
import tkinter as tk
import unicodedata
import urllib.error
import urllib.request
from ctypes import wintypes
from datetime import date
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from tkinter import filedialog, messagebox


APP_NAME = "Buscador Cliente HeadCargo"
APP_VERSION = "1.1.0"
APP_DIR = Path(__file__).resolve().parent
APPDATA_DIR = Path(os.environ.get("APPDATA", str(APP_DIR))) / "BuscadorClienteHeadCargo"
CONFIG_FILE = APPDATA_DIR / "launcher_clientes_onedrive_config.json"
UPDATES_DIR = APPDATA_DIR / "updates"
LOGO_FILE = "logo_buscador.png"
DEFAULT_ROOT_PATH = r"D:\OneDrive - headsoft.com.br\HeadSoft Home - Suporte\Pastinha Clientes\Acessos Clientes"
DESKTOP_EXE_NAME = "Desktop.exe"
SERVER_DCN_NAME = "server.dcn"
SETUP_EXE_NAME = "Setup Buscador Cliente HeadCargo.exe"
GITHUB_REPO = "kauanlauer/Buscador-Cliente"
GITHUB_REPO_URL = f"https://github.com/{GITHUB_REPO}"
GITHUB_MANIFEST_URL = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/github_update_manifest.json"
GITHUB_LATEST_INSTALLER_URL = f"{GITHUB_REPO_URL}/releases/latest/download/Setup%20Buscador%20Cliente%20HeadCargo.exe"
HOTKEY_ID = 0xA110
MOD_ALT = 0x0001
VK_SPACE = 0x20
WM_HOTKEY = 0x0312
WM_QUIT = 0x0012
MAX_VISIBLE_RESULTS = 8
MUTEX_NAME = "HeadCargoClientSearcherSingleton"

WM_USER = 0x0400
TRAY_MESSAGE = WM_USER + 1
WM_DESTROY = 0x0002
WM_COMMAND = 0x0111
WM_CLOSE = 0x0010
WM_LBUTTONUP = 0x0202
WM_RBUTTONUP = 0x0205
WM_NULL = 0x0000
NIM_ADD = 0x00000000
NIM_MODIFY = 0x00000001
NIM_DELETE = 0x00000002
NIF_MESSAGE = 0x00000001
NIF_ICON = 0x00000002
NIF_TIP = 0x00000004
IDI_APPLICATION = 32512
TPM_LEFTALIGN = 0x0000
TPM_BOTTOMALIGN = 0x0020
MF_STRING = 0x0000
MF_SEPARATOR = 0x0800

TRAY_CMD_SHOW_PANEL = 1001
TRAY_CMD_SHOW_SEARCH = 1002
TRAY_CMD_SETTINGS = 1003
TRAY_CMD_REINDEX = 1004
TRAY_CMD_EXIT = 1005

WM_SETTEXT = 0x000C
BM_GETCHECK = 0x00F0
BM_SETCHECK = 0x00F1
BM_CLICK = 0x00F5
BST_CHECKED = 1
BST_UNCHECKED = 0
GWL_STYLE = -16
ES_PASSWORD = 0x0020
CB_GETCOUNT = 0x0146
CB_GETLBTEXTLEN = 0x0149
CB_GETLBTEXT = 0x0148
CB_SETCURSEL = 0x014E
CB_FINDSTRINGEXACT = 0x0158
CB_SELECTSTRING = 0x014D

BG_APP = "#f5f7fb"
BG_PANEL = "#ffffff"
BG_INPUT = "#eef2ff"
BG_RESULTS = "#f8fafc"
TEXT_MAIN = "#0f172a"
TEXT_MUTED = "#64748b"
TEXT_SOFT = "#94a3b8"
ACCENT = "#2563eb"
ACCENT_HOVER = "#1d4ed8"
BORDER = "#dbe4f0"
SUCCESS = "#0f766e"

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
shell32 = ctypes.windll.shell32
WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.DefWindowProcW.restype = ctypes.c_ssize_t


class WNDCLASSW(ctypes.Structure):
    _fields_ = [
        ("style", wintypes.UINT),
        ("lpfnWndProc", WNDPROC),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", wintypes.HINSTANCE),
        ("hIcon", wintypes.HICON),
        ("hCursor", wintypes.HCURSOR),
        ("hbrBackground", wintypes.HBRUSH),
        ("lpszMenuName", wintypes.LPCWSTR),
        ("lpszClassName", wintypes.LPCWSTR),
    ]


class POINT(ctypes.Structure):
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]


class NOTIFYICONDATAW(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("hWnd", wintypes.HWND),
        ("uID", wintypes.UINT),
        ("uFlags", wintypes.UINT),
        ("uCallbackMessage", wintypes.UINT),
        ("hIcon", wintypes.HICON),
        ("szTip", wintypes.WCHAR * 128),
        ("dwState", wintypes.DWORD),
        ("dwStateMask", wintypes.DWORD),
        ("szInfo", wintypes.WCHAR * 256),
        ("uTimeoutOrVersion", wintypes.UINT),
        ("szInfoTitle", wintypes.WCHAR * 64),
        ("dwInfoFlags", wintypes.DWORD),
        ("guidItem", ctypes.c_byte * 16),
        ("hBalloonIcon", wintypes.HICON),
    ]


def normalize_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value.lower())
    chars = []
    for char in normalized:
        if unicodedata.combining(char):
            continue
        chars.append(char if char.isalnum() else " ")
    return " ".join("".join(chars).split())


def get_window_text(hwnd: int) -> str:
    length = user32.GetWindowTextLengthW(hwnd)
    buffer = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buffer, len(buffer))
    return buffer.value


def get_class_name(hwnd: int) -> str:
    buffer = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buffer, len(buffer))
    return buffer.value


@dataclass(slots=True)
class ClientEntry:
    display_name: str
    normalized_name: str
    exe_path: Path
    folder_path: Path


@dataclass(slots=True)
class LoginPreferences:
    username: str
    mark_save_password: bool


def version_tuple(version_text: str) -> tuple[int, ...]:
    parts = []
    for piece in str(version_text).strip().split("."):
        try:
            parts.append(int(piece))
        except ValueError:
            parts.append(0)
    return tuple(parts or [0])


def today_iso() -> str:
    return date.today().isoformat()


def download_text(url: str) -> str:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": f"{APP_NAME}/{APP_VERSION}",
            "Accept": "application/json,text/plain,*/*",
        },
    )
    with urllib.request.urlopen(request, timeout=20) as response:
        charset = response.headers.get_content_charset() or "utf-8"
        return response.read().decode(charset, errors="replace")


def download_file(url: str, destination: Path) -> None:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": f"{APP_NAME}/{APP_VERSION}",
            "Accept": "application/octet-stream,*/*",
        },
    )
    destination.parent.mkdir(parents=True, exist_ok=True)
    with urllib.request.urlopen(request, timeout=60) as response, destination.open("wb") as handle:
        shutil.copyfileobj(response, handle)


def resource_path(filename: str) -> Path:
    base_dir = Path(getattr(sys, "_MEIPASS", APP_DIR))
    return base_dir / filename


def read_text_with_fallback(path: Path) -> tuple[str, str]:
    for encoding in ("utf-8", "latin-1", "cp1252"):
        try:
            return path.read_text(encoding=encoding), encoding
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace"), "utf-8"


def parse_user_list(value: str) -> list[str]:
    try:
        return next(csv.reader([value], skipinitialspace=False))
    except Exception:
        return [item.strip() for item in value.split(",") if item.strip()]


def serialize_user_list(values: list[str]) -> str:
    stream = io.StringIO()
    writer = csv.writer(stream, lineterminator="")
    writer.writerow(values)
    return stream.getvalue()


class ConfigStore:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.data = self.load()

    def load(self) -> dict:
        default = {
            "root_path": DEFAULT_ROOT_PATH,
            "preferred_username": "",
            "mark_save_password": False,
            "last_update_check": "",
        }
        if not self.path.exists():
            self.path.write_text(json.dumps(default, indent=2), encoding="utf-8")
            return default
        try:
            loaded = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            loaded = {}
        if not isinstance(loaded, dict):
            loaded = {}
        for key, value in default.items():
            loaded.setdefault(key, value)
        return loaded

    def save(self) -> None:
        self.path.write_text(json.dumps(self.data, indent=2), encoding="utf-8")

    @property
    def root_path(self) -> Path:
        return Path(self.data["root_path"])

    @root_path.setter
    def root_path(self, value: Path) -> None:
        self.data["root_path"] = str(value)
        self.save()

    @property
    def preferred_username(self) -> str:
        return str(self.data.get("preferred_username", "")).strip()

    @preferred_username.setter
    def preferred_username(self, value: str) -> None:
        self.data["preferred_username"] = value.strip()
        self.save()

    @property
    def mark_save_password(self) -> bool:
        return bool(self.data.get("mark_save_password", False))

    @mark_save_password.setter
    def mark_save_password(self, value: bool) -> None:
        self.data["mark_save_password"] = bool(value)
        self.save()

    def login_preferences(self) -> LoginPreferences:
        return LoginPreferences(self.preferred_username, self.mark_save_password)

    @property
    def last_update_check(self) -> str:
        return str(self.data.get("last_update_check", "")).strip()

    @last_update_check.setter
    def last_update_check(self, value: str) -> None:
        self.data["last_update_check"] = value.strip()
        self.save()


class ClientIndexer:
    def scan(self, root_path: Path) -> list[ClientEntry]:
        if not root_path.exists():
            raise FileNotFoundError(f"Pasta nao encontrada: {root_path}")
        if not root_path.is_dir():
            raise NotADirectoryError(f"O caminho nao e uma pasta: {root_path}")
        entries = []
        for child in sorted(root_path.iterdir(), key=lambda item: item.name.lower()):
            if not child.is_dir():
                continue
            exe_path = self._find_desktop_exe(child)
            if exe_path:
                entries.append(ClientEntry(child.name, normalize_text(child.name), exe_path, child))
        return entries

    def _find_desktop_exe(self, client_folder: Path) -> Path | None:
        direct = client_folder / DESKTOP_EXE_NAME
        if direct.exists():
            return direct
        target = DESKTOP_EXE_NAME.lower()
        for current_root, _, filenames in os.walk(client_folder):
            for filename in filenames:
                if filename.lower() == target:
                    return Path(current_root) / filename
        return None


class HotkeyListener(threading.Thread):
    def __init__(self, callback) -> None:
        super().__init__(daemon=True)
        self.callback = callback
        self.ready = threading.Event()
        self.thread_id = 0
        self.error_message = ""

    def run(self) -> None:
        self.thread_id = kernel32.GetCurrentThreadId()
        if not user32.RegisterHotKey(None, HOTKEY_ID, MOD_ALT, VK_SPACE):
            self.error_message = "Nao foi possivel registrar Alt+Espaco."
            self.ready.set()
            return
        self.ready.set()
        message = wintypes.MSG()
        while True:
            result = user32.GetMessageW(ctypes.byref(message), None, 0, 0)
            if result in (0, -1):
                break
            if message.message == WM_HOTKEY and message.wParam == HOTKEY_ID:
                self.callback()
        user32.UnregisterHotKey(None, HOTKEY_ID)

    def stop(self) -> None:
        if self.thread_id:
            user32.PostThreadMessageW(self.thread_id, WM_QUIT, 0, 0)


class HeadCargoLoginAutomator:
    def __init__(self, app) -> None:
        self.app = app

    def apply_async(self, process_id: int, preferences: LoginPreferences) -> None:
        if not preferences.username:
            return
        threading.Thread(target=self._apply_worker, args=(process_id, preferences), daemon=True).start()

    def _apply_worker(self, process_id: int, preferences: LoginPreferences) -> None:
        hwnd = self._wait_for_window(process_id)
        if not hwnd:
            self.app.after(0, lambda: self.app.set_status("Cliente abriu, mas a tela de login nao foi localizada."))
            return
        user32.ShowWindow(hwnd, 5)
        user32.SetForegroundWindow(hwnd)
        user32.BringWindowToTop(hwnd)
        children = self._enumerate_children(hwnd)
        username_control = self._find_username_control(children)
        password_hwnd = self._find_password_control(children)
        checkbox_hwnd = self._find_save_checkbox(children)
        if username_control:
            self._apply_username(username_control, preferences.username)
        if checkbox_hwnd:
            self._set_checkbox_state(checkbox_hwnd, preferences.mark_save_password)
        if password_hwnd:
            user32.SetFocus(password_hwnd)
        self.app.after(0, lambda: self.app.set_status(f"Usuario '{preferences.username}' preparado no login."))

    def _wait_for_window(self, process_id: int) -> int | None:
        for _ in range(80):
            hwnd = self._find_window(process_id)
            if hwnd:
                return hwnd
            kernel32.Sleep(150)
        return None

    def _find_window(self, process_id: int) -> int | None:
        matches = []

        @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        def enum_proc(hwnd, _):
            pid = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value == process_id and user32.IsWindowVisible(hwnd):
                if "headcargo" in normalize_text(get_window_text(hwnd)):
                    matches.append(hwnd)
                    return False
            return True

        user32.EnumWindows(enum_proc, 0)
        return matches[0] if matches else None

    def _enumerate_children(self, hwnd: int) -> list[dict]:
        items = []

        @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        def enum_proc(child_hwnd, _):
            items.append(
                {
                    "hwnd": child_hwnd,
                    "class_name": get_class_name(child_hwnd).lower(),
                    "text": normalize_text(get_window_text(child_hwnd)),
                    "style": user32.GetWindowLongW(child_hwnd, GWL_STYLE),
                }
            )
            return True

        user32.EnumChildWindows(hwnd, enum_proc, 0)
        return items

    def _find_username_control(self, children: list[dict]) -> dict | None:
        for child in children:
            if "combo" in child["class_name"]:
                return child
        for child in children:
            if "edit" in child["class_name"] and not (child["style"] & ES_PASSWORD):
                return child
        return None

    def _find_password_control(self, children: list[dict]) -> int | None:
        edits = []
        for child in children:
            if "edit" not in child["class_name"]:
                continue
            if child["style"] & ES_PASSWORD:
                return child["hwnd"]
            edits.append(child["hwnd"])
        return edits[1] if len(edits) > 1 else None

    def _find_save_checkbox(self, children: list[dict]) -> int | None:
        for child in children:
            if "salvar senha" in child["text"] or ("check" in child["class_name"] and "senha" in child["text"]):
                return child["hwnd"]
        return None

    def _set_checkbox_state(self, hwnd: int, should_check: bool) -> None:
        expected = BST_CHECKED if should_check else BST_UNCHECKED
        current = user32.SendMessageW(hwnd, BM_GETCHECK, 0, 0)
        if current == expected:
            return
        user32.SendMessageW(hwnd, BM_SETCHECK, expected, 0)
        if user32.SendMessageW(hwnd, BM_GETCHECK, 0, 0) != expected:
            user32.SendMessageW(hwnd, BM_CLICK, 0, 0)

    def _apply_username(self, control: dict, username: str) -> None:
        hwnd = control["hwnd"]
        class_name = control["class_name"]
        if "combo" in class_name:
            index = user32.SendMessageW(hwnd, CB_FINDSTRINGEXACT, -1, username)
            if index != -1:
                user32.SendMessageW(hwnd, CB_SETCURSEL, index, 0)
            else:
                user32.SendMessageW(hwnd, CB_SELECTSTRING, -1, username)
                user32.SendMessageW(hwnd, WM_SETTEXT, 0, username)
            combo_edit = user32.FindWindowExW(hwnd, 0, "Edit", None)
            if combo_edit:
                user32.SendMessageW(combo_edit, WM_SETTEXT, 0, username)
            return
        user32.SendMessageW(hwnd, WM_SETTEXT, 0, username)


class TrayIcon(threading.Thread):
    def __init__(self, app) -> None:
        super().__init__(daemon=True)
        self.app = app
        self.ready = threading.Event()
        self.thread_id = 0
        self.hwnd = None
        self.hicon = None
        self.class_name = "HeadCargoClientSearchTrayClass"
        self.wnd_proc = WNDPROC(self._window_proc)

    def run(self) -> None:
        self.thread_id = kernel32.GetCurrentThreadId()
        instance = kernel32.GetModuleHandleW(None)
        self.hicon = user32.LoadIconW(None, IDI_APPLICATION)
        wc = WNDCLASSW()
        wc.lpfnWndProc = self.wnd_proc
        wc.hInstance = instance
        wc.lpszClassName = self.class_name
        wc.hIcon = self.hicon
        user32.RegisterClassW(ctypes.byref(wc))
        self.hwnd = user32.CreateWindowExW(0, self.class_name, APP_NAME, 0, 0, 0, 0, 0, 0, 0, instance, None)
        self._add_icon()
        self.ready.set()
        message = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(message), None, 0, 0) > 0:
            user32.TranslateMessage(ctypes.byref(message))
            user32.DispatchMessageW(ctypes.byref(message))
        self._remove_icon()

    def stop(self) -> None:
        if self.hwnd:
            user32.PostMessageW(self.hwnd, WM_CLOSE, 0, 0)

    def _add_icon(self) -> None:
        data = NOTIFYICONDATAW()
        data.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        data.hWnd = self.hwnd
        data.uID = 1
        data.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP
        data.uCallbackMessage = TRAY_MESSAGE
        data.hIcon = self.hicon
        data.szTip = APP_NAME
        shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(data))

    def _remove_icon(self) -> None:
        data = NOTIFYICONDATAW()
        data.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        data.hWnd = self.hwnd
        data.uID = 1
        shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(data))

    def _show_menu(self) -> None:
        menu = user32.CreatePopupMenu()
        user32.AppendMenuW(menu, MF_STRING, TRAY_CMD_SHOW_PANEL, "Abrir painel")
        user32.AppendMenuW(menu, MF_STRING, TRAY_CMD_SHOW_SEARCH, "Abrir busca")
        user32.AppendMenuW(menu, MF_STRING, TRAY_CMD_SETTINGS, "Configuracoes")
        user32.AppendMenuW(menu, MF_STRING, TRAY_CMD_REINDEX, "Reindexar")
        user32.AppendMenuW(menu, MF_SEPARATOR, 0, None)
        user32.AppendMenuW(menu, MF_STRING, TRAY_CMD_EXIT, "Sair")
        point = POINT()
        user32.GetCursorPos(ctypes.byref(point))
        user32.SetForegroundWindow(self.hwnd)
        selected = user32.TrackPopupMenu(menu, TPM_LEFTALIGN | TPM_BOTTOMALIGN, point.x, point.y, 0, self.hwnd, None)
        if selected:
            user32.PostMessageW(self.hwnd, WM_COMMAND, selected, 0)
        user32.PostMessageW(self.hwnd, WM_NULL, 0, 0)
        user32.DestroyMenu(menu)

    def _window_proc(self, hwnd, msg, wparam, lparam):
        if msg == TRAY_MESSAGE:
            if lparam == WM_LBUTTONUP:
                self.app.after(0, self.app.show_palette)
                return 0
            if lparam == WM_RBUTTONUP:
                self._show_menu()
                return 0
        if msg == WM_COMMAND:
            command_id = wparam & 0xFFFF
            if command_id == TRAY_CMD_SHOW_PANEL:
                self.app.after(0, self.app.show_panel)
            elif command_id == TRAY_CMD_SHOW_SEARCH:
                self.app.after(0, self.app.show_palette)
            elif command_id == TRAY_CMD_SETTINGS:
                self.app.after(0, self.app.open_settings)
            elif command_id == TRAY_CMD_REINDEX:
                self.app.after(0, self.app.refresh_index)
            elif command_id == TRAY_CMD_EXIT:
                self.app.after(0, self.app.shutdown)
            return 0
        if msg == WM_CLOSE:
            user32.DestroyWindow(hwnd)
            return 0
        if msg == WM_DESTROY:
            user32.PostQuitMessage(0)
            return 0
        return user32.DefWindowProcW(hwnd, msg, wparam, lparam)


class SettingsWindow(tk.Toplevel):
    def __init__(self, app) -> None:
        super().__init__(app)
        self.app = app
        self.title(f"{APP_NAME} - Configuracoes")
        self.geometry("700x560")
        self.resizable(False, False)
        self.configure(bg=BG_APP)
        self.transient(app)
        self.grab_set()
        if getattr(app, "app_icon", None):
            self.iconphoto(True, app.app_icon)
        self.path_var = tk.StringVar(value=str(app.config_store.root_path))
        self.username_var = tk.StringVar(value=app.config_store.preferred_username)
        self.save_password_var = tk.BooleanVar(value=app.config_store.mark_save_password)
        self._build_ui()
        self._center()

    def _build_ui(self) -> None:
        container = tk.Frame(self, bg=BG_APP, padx=18, pady=18)
        container.pack(fill="both", expand=True)
        card = tk.Frame(container, bg=BG_PANEL, padx=18, pady=18, highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="both", expand=True)
        title_row = tk.Frame(card, bg=BG_PANEL)
        title_row.pack(fill="x")
        if getattr(self.app, "header_logo_small", None):
            tk.Label(title_row, image=self.app.header_logo_small, bg=BG_PANEL).pack(side="left", padx=(0, 12))
        text_col = tk.Frame(title_row, bg=BG_PANEL)
        text_col.pack(side="left", fill="x", expand=True)
        tk.Label(text_col, text="Configuracoes do buscador", font=("Segoe UI Semibold", 15), fg=TEXT_MAIN, bg=BG_PANEL).pack(anchor="w")
        tk.Label(text_col, text="Defina a pasta dos clientes e o usuario padrao da tela do HeadCargo.", font=("Segoe UI", 9), fg=TEXT_MUTED, bg=BG_PANEL).pack(anchor="w", pady=(6, 14))
        tk.Label(card, text="Pasta dos clientes", font=("Segoe UI Semibold", 10), fg=TEXT_MAIN, bg=BG_PANEL).pack(anchor="w")
        row = tk.Frame(card, bg=BG_PANEL)
        row.pack(fill="x", pady=(6, 12))
        tk.Entry(row, textvariable=self.path_var, font=("Consolas", 10), relief="flat", bd=0, fg=TEXT_MAIN, bg=BG_RESULTS, insertbackground=TEXT_MAIN).pack(side="left", fill="x", expand=True, ipady=8)
        tk.Button(row, text="Selecionar", command=self._select_path, font=("Segoe UI", 9), relief="flat", bd=0, padx=14, pady=8, fg=TEXT_MAIN, bg="#e2e8f0", activebackground="#cbd5e1", activeforeground=TEXT_MAIN).pack(side="left", padx=(10, 0))
        tk.Label(card, text="Usuario padrao do HeadCargo", font=("Segoe UI Semibold", 10), fg=TEXT_MAIN, bg=BG_PANEL).pack(anchor="w")
        tk.Entry(card, textvariable=self.username_var, font=("Segoe UI", 10), relief="flat", bd=0, fg=TEXT_MAIN, bg=BG_RESULTS, insertbackground=TEXT_MAIN).pack(fill="x", pady=(6, 12), ipady=8)
        tk.Checkbutton(card, text="Marcar 'Salvar senha' automaticamente", variable=self.save_password_var, font=("Segoe UI", 10), fg=TEXT_MAIN, bg=BG_PANEL, activebackground=BG_PANEL, activeforeground=TEXT_MAIN, selectcolor=BG_PANEL).pack(anchor="w")
        tk.Label(card, text="Atualizacoes automáticas", font=("Segoe UI Semibold", 10), fg=TEXT_MAIN, bg=BG_PANEL).pack(anchor="w", pady=(16, 0))
        update_card = tk.Frame(card, bg=BG_RESULTS, padx=12, pady=12, highlightbackground=BORDER, highlightthickness=1)
        update_card.pack(fill="x", pady=(6, 0))
        tk.Label(update_card, text="Origem: GitHub / kauanlauer/Buscador-Cliente", font=("Consolas", 9), fg=TEXT_MAIN, bg=BG_RESULTS, justify="left", anchor="w").pack(fill="x")
        tk.Label(update_card, text="O programa verifica novas versoes todo dia e voce tambem pode usar o botao 'Verificar atualizacao' no painel principal.", font=("Segoe UI", 9), fg=TEXT_MUTED, bg=BG_RESULTS, justify="left", wraplength=590).pack(anchor="w", pady=(8, 0))
        buttons = tk.Frame(card, bg=BG_PANEL)
        buttons.pack(fill="x", side="bottom", pady=(24, 0))
        tk.Button(buttons, text="Salvar", command=self._save, font=("Segoe UI Semibold", 10), relief="flat", bd=0, padx=18, pady=10, fg="#ffffff", bg=ACCENT, activebackground=ACCENT_HOVER, activeforeground="#ffffff").pack(side="left")
        tk.Button(buttons, text="Cancelar", command=self.destroy, font=("Segoe UI", 10), relief="flat", bd=0, padx=18, pady=10, fg=TEXT_MAIN, bg="#e2e8f0", activebackground="#cbd5e1", activeforeground=TEXT_MAIN).pack(side="left", padx=(10, 0))

    def _center(self) -> None:
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = max((self.winfo_screenwidth() - width) // 2, 0)
        y = max((self.winfo_screenheight() - height) // 3, 0)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _select_path(self) -> None:
        selected = filedialog.askdirectory(title="Selecione a pasta raiz dos clientes", initialdir=self.path_var.get() or str(self.app.config_store.root_path))
        if selected:
            self.path_var.set(selected)

    def _save(self) -> None:
        selected_path = Path(self.path_var.get().strip())
        if not selected_path.exists():
            messagebox.showerror("Pasta invalida", "Selecione uma pasta valida.")
            return
        self.app.config_store.root_path = selected_path
        self.app.config_store.preferred_username = self.username_var.get().strip()
        self.app.config_store.mark_save_password = self.save_password_var.get()
        self.app.refresh_config_labels()
        self.app.refresh_index()
        self.app.set_status("Configuracoes salvas.")
        self.destroy()


class SearchPalette(tk.Toplevel):
    def __init__(self, master) -> None:
        super().__init__(master)
        self.app = master
        self.filtered_entries = []
        self.query_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Digite o nome do cliente.")
        self.result_count_var = tk.StringVar(value="")
        self.withdraw()
        self.overrideredirect(True)
        self.transient(master)
        self.configure(bg="#d6deea")
        self._build_ui()
        self._bind_events()

    def _build_ui(self) -> None:
        shell = tk.Frame(self, bg="#d6deea", padx=1, pady=1)
        shell.pack(fill="both", expand=True)
        body = tk.Frame(shell, bg=BG_PANEL)
        body.pack(fill="both", expand=True)
        header = tk.Frame(body, bg=BG_PANEL, padx=22, pady=18)
        header.pack(fill="x")
        tk.Label(header, text="Pesquisar cliente", font=("Segoe UI Semibold", 15), fg=TEXT_MAIN, bg=BG_PANEL).pack(anchor="w")
        tk.Label(header, text="Alt+Espaco abre a busca. Enter executa o Desktop.exe.", font=("Segoe UI", 9), fg=TEXT_MUTED, bg=BG_PANEL).pack(anchor="w", pady=(4, 0))
        input_wrap = tk.Frame(body, bg=BG_PANEL, padx=22)
        input_wrap.pack(fill="x")
        search_box = tk.Frame(input_wrap, bg=BG_INPUT, padx=14, pady=10)
        search_box.pack(fill="x")
        tk.Label(search_box, text=">", font=("Consolas", 14), fg=ACCENT, bg=BG_INPUT).pack(side="left", padx=(0, 10))
        self.search_entry = tk.Entry(search_box, textvariable=self.query_var, font=("Segoe UI", 16), bd=0, relief="flat", bg=BG_INPUT, fg=TEXT_MAIN, insertbackground=TEXT_MAIN)
        self.search_entry.pack(fill="x", expand=True)
        results_wrap = tk.Frame(body, bg=BG_PANEL, padx=22)
        results_wrap.pack(fill="both", expand=True, pady=(14, 18))
        self.placeholder_label = tk.Label(results_wrap, text="Digite para filtrar. Nenhum cliente aparece com a busca vazia.", font=("Segoe UI", 10), fg=TEXT_MUTED, bg=BG_RESULTS, anchor="w", justify="left", padx=16, pady=16)
        self.placeholder_label.pack(fill="x")
        self.results_panel = tk.Frame(results_wrap, bg=BG_RESULTS)
        top_line = tk.Frame(self.results_panel, bg=BG_RESULTS)
        top_line.pack(fill="x", padx=14, pady=(12, 6))
        tk.Label(top_line, textvariable=self.result_count_var, font=("Segoe UI Semibold", 9), fg=TEXT_MUTED, bg=BG_RESULTS).pack(side="left")
        self.result_list = tk.Listbox(self.results_panel, activestyle="none", font=("Segoe UI", 11), bd=0, relief="flat", bg=BG_RESULTS, fg=TEXT_MAIN, selectbackground="#dbeafe", selectforeground=TEXT_MAIN, highlightthickness=0, exportselection=False)
        self.result_list.pack(fill="both", expand=True, padx=8, pady=(0, 4))
        footer = tk.Frame(body, bg=BG_PANEL, padx=22)
        footer.pack(fill="x", pady=(0, 16))
        tk.Label(footer, textvariable=self.status_var, font=("Segoe UI", 9), fg=TEXT_SOFT, bg=BG_PANEL, anchor="w", justify="left").pack(fill="x")

    def _bind_events(self) -> None:
        self.query_var.trace_add("write", lambda *_: self.update_results())
        self.search_entry.bind("<Down>", self._focus_next_result)
        self.search_entry.bind("<Up>", self._focus_previous_result)
        self.search_entry.bind("<Return>", lambda _: self.launch_selected())
        self.search_entry.bind("<Escape>", lambda _: self.hide_palette())
        self.result_list.bind("<Double-Button-1>", lambda _: self.launch_selected())
        self.result_list.bind("<Return>", lambda _: self.launch_selected())
        self.result_list.bind("<<ListboxSelect>>", lambda _: self._update_status_for_selection())
        self.result_list.bind("<Escape>", lambda _: self.hide_palette())
        self.bind("<Escape>", lambda _: self.hide_palette())
        self.bind("<FocusOut>", lambda _: self.after(120, self._hide_if_unfocused))

    def _hide_if_unfocused(self) -> None:
        focused = self.focus_displayof()
        if focused not in {self, self.search_entry, self.result_list}:
            self.hide_palette()

    def show_palette(self) -> None:
        self.deiconify()
        width, height = 720, 330
        x = max((self.winfo_screenwidth() - width) // 2, 0)
        y = max((self.winfo_screenheight() - height) // 5, 0)
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.lift()
        self.attributes("-topmost", True)
        self.query_var.set("")
        self.search_entry.focus_force()
        self.update_results()

    def hide_palette(self) -> None:
        self.withdraw()

    def toggle_palette(self) -> None:
        self.show_palette() if self.state() == "withdrawn" else self.hide_palette()

    def update_results(self) -> None:
        query = normalize_text(self.query_var.get())
        if not query:
            self.filtered_entries = []
            self.result_list.delete(0, tk.END)
            self.results_panel.pack_forget()
            self.placeholder_label.configure(text="Digite para filtrar. Nenhum cliente aparece com a busca vazia.")
            self.placeholder_label.pack(fill="x")
            self.result_count_var.set("")
            self.status_var.set("Indexando clientes..." if self.app.is_scanning else f"{len(self.app.entries)} clientes prontos para busca.")
            return
        ranked = self.app.rank_entries(query)
        self.filtered_entries = ranked[:MAX_VISIBLE_RESULTS]
        self.result_list.delete(0, tk.END)
        if not self.filtered_entries:
            self.results_panel.pack_forget()
            self.placeholder_label.configure(text="Nenhum cliente encontrado para esse texto.")
            self.placeholder_label.pack(fill="x")
            self.result_count_var.set("")
            self.status_var.set("Tente outro nome ou reindexe a pasta.")
            return
        self.placeholder_label.pack_forget()
        self.results_panel.pack(fill="both", expand=True)
        for entry in self.filtered_entries:
            self.result_list.insert(tk.END, entry.display_name)
        self.result_list.selection_clear(0, tk.END)
        self.result_list.selection_set(0)
        self.result_list.activate(0)
        self.result_count_var.set(f"{len(self.filtered_entries)} resultado(s)")
        self._update_status_for_selection()

    def _focus_next_result(self, _=None):
        if self.filtered_entries:
            idx = min((self.result_list.curselection() or (0,))[0] + 1, len(self.filtered_entries) - 1)
            self._select(idx)
        return "break"

    def _focus_previous_result(self, _=None):
        if self.filtered_entries:
            idx = max((self.result_list.curselection() or (0,))[0] - 1, 0)
            self._select(idx)
        return "break"

    def _select(self, index: int) -> None:
        self.result_list.selection_clear(0, tk.END)
        self.result_list.selection_set(index)
        self.result_list.activate(index)
        self.result_list.see(index)
        self._update_status_for_selection()

    def _selected_entry(self) -> ClientEntry | None:
        selection = self.result_list.curselection()
        return self.filtered_entries[selection[0]] if selection else None

    def _update_status_for_selection(self) -> None:
        entry = self._selected_entry()
        if entry:
            self.status_var.set(str(entry.exe_path))

    def launch_selected(self) -> None:
        entry = self._selected_entry()
        if entry:
            self.app.launch_entry(entry)


class LauncherApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("720x920")
        self.minsize(720, 920)
        self.configure(bg=BG_APP)
        self.protocol("WM_DELETE_WINDOW", self.minimize_to_taskbar)
        self.config_store = ConfigStore(CONFIG_FILE)
        self.indexer = ClientIndexer()
        self.login_automator = HeadCargoLoginAutomator(self)
        self.hotkey_listener = None
        self.tray_icon = None
        self.entries = []
        self.is_scanning = False
        self.is_checking_updates = False
        self.settings_window = None
        self.status_var = tk.StringVar(value="Preparando buscador...")
        self.folder_var = tk.StringVar()
        self.user_var = tk.StringVar()
        self.save_password_text = tk.StringVar()
        self.update_source_var = tk.StringVar()
        self.hotkey_var = tk.StringVar(value="Hotkey: Alt+Espaco")
        self.count_var = tk.StringVar(value=f"v{APP_VERSION} | Clientes indexados: 0")
        self.app_icon = None
        self.header_logo = None
        self.header_logo_small = None
        self._load_brand_assets()
        self._build_menu()
        self._build_ui()
        self.refresh_config_labels()
        self.palette = SearchPalette(self)
        self.after(100, self._start_hotkey_listener)
        self.after(160, self._start_tray_icon)
        self.after(250, self.refresh_index)
        self.after(3500, self.check_for_daily_updates)

    def _build_menu(self) -> None:
        menu_bar = tk.Menu(self)
        menu_busca = tk.Menu(menu_bar, tearoff=False)
        menu_busca.add_command(label="Abrir busca", command=self.show_palette)
        menu_busca.add_command(label="Abrir painel", command=self.show_panel)
        menu_busca.add_command(label="Reindexar", command=self.refresh_index)
        menu_busca.add_command(label="Verificar atualizacao", command=self.check_for_updates)
        menu_busca.add_separator()
        menu_busca.add_command(label="Sair", command=self.shutdown)
        menu_bar.add_cascade(label="Buscador", menu=menu_busca)
        menu_config = tk.Menu(menu_bar, tearoff=False)
        menu_config.add_command(label="Configuracoes", command=self.open_settings)
        menu_config.add_command(label="Abrir pasta atual", command=self.open_root_folder)
        menu_bar.add_cascade(label="Configuracoes", menu=menu_config)
        self.configure(menu=menu_bar)

    def _load_brand_assets(self) -> None:
        logo_path = resource_path(LOGO_FILE)
        if not logo_path.exists():
            return
        try:
            logo = tk.PhotoImage(file=str(logo_path))
            self.app_icon = logo
            self.iconphoto(True, logo)
            width = max(logo.width(), 1)
            factor = max(1, width // 44)
            self.header_logo_small = logo.subsample(factor, factor)
            width_large = max(logo.width(), 1)
            factor_large = max(1, width_large // 78)
            self.header_logo = logo.subsample(factor_large, factor_large)
        except Exception:
            self.app_icon = None
            self.header_logo = None
            self.header_logo_small = None

    def _build_ui(self) -> None:
        container = tk.Frame(self, bg=BG_APP, padx=18, pady=18)
        container.pack(fill="both", expand=True)
        header = tk.Frame(container, bg=BG_PANEL, padx=20, pady=18, highlightbackground=BORDER, highlightthickness=1)
        header.pack(fill="x")
        branding = tk.Frame(header, bg=BG_PANEL)
        branding.pack(fill="x")
        if self.header_logo:
            tk.Label(branding, image=self.header_logo, bg=BG_PANEL).pack(side="left", padx=(0, 12))
        title_block = tk.Frame(branding, bg=BG_PANEL)
        title_block.pack(side="left", fill="x", expand=True)
        tk.Label(title_block, text=APP_NAME, font=("Segoe UI Semibold", 17), fg=TEXT_MAIN, bg=BG_PANEL).pack(anchor="w")
        tk.Label(title_block, text="Busca rapida do Desktop.exe e prepara a tela de login do HeadCargo.", font=("Segoe UI", 10), fg=TEXT_MUTED, bg=BG_PANEL).pack(anchor="w", pady=(4, 0))
        chips = tk.Frame(header, bg=BG_PANEL)
        chips.pack(fill="x", pady=(14, 0))
        tk.Label(chips, textvariable=self.hotkey_var, font=("Segoe UI Semibold", 9), fg=ACCENT, bg="#e0edff", padx=10, pady=5).pack(side="left")
        tk.Label(chips, textvariable=self.count_var, font=("Segoe UI Semibold", 9), fg=SUCCESS, bg="#dcfce7", padx=10, pady=5).pack(side="left", padx=(10, 0))
        config_card = tk.Frame(container, bg=BG_PANEL, padx=20, pady=16, highlightbackground=BORDER, highlightthickness=1)
        config_card.pack(fill="x", pady=(14, 0))
        tk.Label(config_card, text="Pasta atual dos clientes", font=("Segoe UI Semibold", 10), fg=TEXT_MAIN, bg=BG_PANEL).pack(anchor="w")
        tk.Label(config_card, textvariable=self.folder_var, font=("Consolas", 10), fg=TEXT_MUTED, bg=BG_PANEL, justify="left", anchor="w", wraplength=480).pack(fill="x", pady=(8, 12))
        tk.Label(config_card, text="Usuario padrao do HeadCargo", font=("Segoe UI Semibold", 10), fg=TEXT_MAIN, bg=BG_PANEL).pack(anchor="w")
        tk.Label(config_card, textvariable=self.user_var, font=("Segoe UI", 10), fg=TEXT_MUTED, bg=BG_PANEL, anchor="w").pack(fill="x", pady=(6, 6))
        tk.Label(config_card, textvariable=self.save_password_text, font=("Segoe UI", 10), fg=TEXT_MUTED, bg=BG_PANEL, anchor="w").pack(fill="x")
        tk.Label(config_card, text="Origem de atualizacao", font=("Segoe UI Semibold", 10), fg=TEXT_MAIN, bg=BG_PANEL).pack(anchor="w", pady=(12, 0))
        tk.Label(config_card, textvariable=self.update_source_var, font=("Consolas", 9), fg=TEXT_MUTED, bg=BG_PANEL, justify="left", anchor="w", wraplength=480).pack(fill="x", pady=(6, 0))
        actions = tk.Frame(container, bg=BG_APP)
        actions.pack(fill="x", pady=(14, 0))
        tk.Button(actions, text="Abrir busca", command=self.show_palette, font=("Segoe UI Semibold", 10), relief="flat", bd=0, padx=16, pady=10, fg="#ffffff", bg=ACCENT, activebackground=ACCENT_HOVER, activeforeground="#ffffff").pack(side="left")
        tk.Button(actions, text="Configuracoes", command=self.open_settings, font=("Segoe UI", 10), relief="flat", bd=0, padx=16, pady=10, fg=TEXT_MAIN, bg="#e2e8f0", activebackground="#cbd5e1", activeforeground=TEXT_MAIN).pack(side="left", padx=(10, 0))
        tk.Button(actions, text="Reindexar", command=self.refresh_index, font=("Segoe UI", 10), relief="flat", bd=0, padx=16, pady=10, fg=TEXT_MAIN, bg="#e2e8f0", activebackground="#cbd5e1", activeforeground=TEXT_MAIN).pack(side="left", padx=(10, 0))
        tk.Button(actions, text="Verificar atualizacao", command=self.check_for_updates, font=("Segoe UI", 10), relief="flat", bd=0, padx=16, pady=10, fg=TEXT_MAIN, bg="#e2e8f0", activebackground="#cbd5e1", activeforeground=TEXT_MAIN).pack(side="left", padx=(10, 0))
        hint_card = tk.Frame(container, bg=BG_PANEL, padx=20, pady=16, highlightbackground=BORDER, highlightthickness=1)
        hint_card.pack(fill="both", expand=True, pady=(14, 0))
        tk.Label(hint_card, text="Como usar", font=("Segoe UI Semibold", 10), fg=TEXT_MAIN, bg=BG_PANEL).pack(anchor="w")
        tk.Label(hint_card, text="1. Deixe o buscador aberto.\n2. Use Alt+Espaco.\n3. Digite o cliente e pressione Enter.\n4. O programa abre o cliente, preenche o usuario e deixa o foco na senha.", font=("Segoe UI", 10), fg=TEXT_MUTED, bg=BG_PANEL, justify="left", anchor="w").pack(fill="x", pady=(8, 8))
        tk.Label(hint_card, textvariable=self.status_var, font=("Segoe UI", 9), fg=TEXT_SOFT, bg=BG_PANEL, justify="left", anchor="w").pack(fill="x")

    def refresh_config_labels(self) -> None:
        self.folder_var.set(str(self.config_store.root_path))
        self.user_var.set(self.config_store.preferred_username or "Nao definido")
        self.save_password_text.set("Salvar senha: marcar automaticamente" if self.config_store.mark_save_password else "Salvar senha: nao marcar automaticamente")
        self.update_source_var.set("GitHub: kauanlauer/Buscador-Cliente")

    def set_status(self, message: str) -> None:
        self.status_var.set(message)

    def show_panel(self) -> None:
        self.deiconify()
        self.state("normal")
        self.lift()
        self.focus_force()

    def minimize_to_taskbar(self) -> None:
        self.iconify()

    def show_palette(self) -> None:
        self.palette.show_palette()

    def open_settings(self) -> None:
        if self.settings_window and self.settings_window.winfo_exists():
            self.settings_window.lift()
            self.settings_window.focus_force()
            return
        self.settings_window = SettingsWindow(self)

    def check_for_daily_updates(self) -> None:
        if not getattr(sys, "frozen", False):
            return
        if self.config_store.last_update_check == today_iso():
            return
        self._start_update_check(silent_no_update=True, silent_errors=True)

    def check_for_updates(self) -> None:
        self._start_update_check(silent_no_update=False, silent_errors=False)

    def _start_update_check(self, silent_no_update: bool, silent_errors: bool) -> None:
        if self.is_checking_updates:
            if not silent_no_update:
                self.set_status("Ja existe uma verificacao de atualizacao em andamento.")
            return
        self.is_checking_updates = True
        self.set_status("Verificando atualizacoes no GitHub...")
        threading.Thread(
            target=self._update_check_worker,
            args=(silent_no_update, silent_errors),
            daemon=True,
        ).start()

    def _update_check_worker(self, silent_no_update: bool, silent_errors: bool) -> None:
        try:
            manifest = json.loads(download_text(GITHUB_MANIFEST_URL))
            latest_version = str(manifest.get("version", "")).strip()
            installer_url = str(manifest.get("installer_url", "")).strip() or GITHUB_LATEST_INSTALLER_URL
            notes = str(manifest.get("notes", "")).strip()
            if not latest_version or not installer_url:
                raise ValueError("Manifesto do GitHub sem 'version' ou 'installer_url'.")
            self.after(
                0,
                lambda: self._handle_update_manifest(
                    latest_version,
                    installer_url,
                    notes,
                    silent_no_update,
                    silent_errors,
                ),
            )
        except Exception as exc:
            self.after(0, lambda: self._handle_update_error(str(exc), silent_errors))

    def _handle_update_manifest(
        self,
        latest_version: str,
        installer_url: str,
        notes: str,
        silent_no_update: bool,
        silent_errors: bool,
    ) -> None:
        self.is_checking_updates = False
        self.config_store.last_update_check = today_iso()
        if version_tuple(latest_version) <= version_tuple(APP_VERSION):
            self.set_status(f"Sem atualizacao. Versao atual: {APP_VERSION}.")
            if not silent_no_update:
                messagebox.showinfo("Atualizacao", f"Voce ja esta na versao atual ({APP_VERSION}).")
            return

        if not getattr(sys, "frozen", False):
            messagebox.showinfo(
                "Atualizacao disponivel",
                f"Nova versao encontrada: {latest_version}\n\nPara testar a atualizacao automatica, abra o .exe instalado.",
            )
            return

        details = f"Versao atual: {APP_VERSION}\nNova versao: {latest_version}"
        if notes:
            details += f"\n\nNovidades:\n{notes}"
        confirm = messagebox.askyesno("Atualizacao disponivel", f"{details}\n\nDeseja atualizar agora?")
        if confirm:
            self.download_and_apply_update(installer_url, latest_version)
        else:
            self.set_status(f"Atualizacao {latest_version} adiada pelo usuario.")

    def _handle_update_error(self, error_message: str, silent_errors: bool) -> None:
        self.is_checking_updates = False
        self.set_status("Nao foi possivel verificar atualizacoes.")
        if not silent_errors:
            messagebox.showerror("Falha ao verificar atualizacao", error_message)

    def download_and_apply_update(self, installer_url: str, target_version: str) -> None:
        self.set_status(f"Baixando atualizacao {target_version}...")
        threading.Thread(
            target=self._download_and_apply_update_worker,
            args=(installer_url, target_version),
            daemon=True,
        ).start()

    def _download_and_apply_update_worker(self, installer_url: str, target_version: str) -> None:
        staged_setup = UPDATES_DIR / f"Setup Buscador Cliente HeadCargo_{target_version}.exe"
        try:
            download_file(installer_url, staged_setup)
            self.after(0, lambda: self._launch_update_installer(staged_setup, target_version))
        except Exception as exc:
            self.after(0, lambda: messagebox.showerror("Falha ao baixar atualizacao", str(exc)))
            self.after(0, lambda: self.set_status("Falha ao baixar a nova versao."))

    def _launch_update_installer(self, staged_setup: Path, target_version: str) -> None:
        if not staged_setup.exists():
            messagebox.showerror("Falha ao atualizar", "O instalador baixado nao foi encontrado.")
            return
        try:
            subprocess.Popen(
                [
                    str(staged_setup),
                    "/VERYSILENT",
                    "/SUPPRESSMSGBOXES",
                    "/NORESTART",
                    "/SP-",
                    "/CLOSEAPPLICATIONS",
                    "/FORCECLOSEAPPLICATIONS",
                ],
                cwd=str(staged_setup.parent),
            )
        except OSError as exc:
            messagebox.showerror("Falha ao iniciar atualizacao", str(exc))
            return
        self.set_status(f"Atualizando para a versao {target_version}...")
        self.after(400, self.shutdown)

    def _start_hotkey_listener(self) -> None:
        self.hotkey_listener = HotkeyListener(lambda: self.after(0, self.palette.toggle_palette))
        self.hotkey_listener.start()
        self.hotkey_listener.ready.wait(1.0)
        if self.hotkey_listener.error_message:
            self.hotkey_var.set("Hotkey indisponivel")
            self.set_status(self.hotkey_listener.error_message)

    def _start_tray_icon(self) -> None:
        self.tray_icon = TrayIcon(self)
        self.tray_icon.start()
        self.tray_icon.ready.wait(2.0)

    def refresh_index(self) -> None:
        if self.is_scanning:
            return
        self.is_scanning = True
        self.refresh_config_labels()
        self.set_status(f"Indexando clientes em: {self.config_store.root_path}")
        self.count_var.set(f"v{APP_VERSION} | Clientes indexados: ...")
        threading.Thread(target=self._scan_worker, args=(self.config_store.root_path,), daemon=True).start()

    def _scan_worker(self, root_path: Path) -> None:
        try:
            entries = self.indexer.scan(root_path)
            self.after(0, lambda: self._finish_scan_success(entries))
        except Exception as exc:
            self.after(0, lambda: self._finish_scan_error(str(exc)))

    def _finish_scan_success(self, entries: list[ClientEntry]) -> None:
        self.is_scanning = False
        self.entries = entries
        self.count_var.set(f"v{APP_VERSION} | Clientes indexados: {len(entries)}")
        self.set_status(f"Indexacao concluida. {len(entries)} cliente(s) disponivel(is).")
        self.palette.update_results()

    def _finish_scan_error(self, error_message: str) -> None:
        self.is_scanning = False
        self.entries = []
        self.count_var.set(f"v{APP_VERSION} | Clientes indexados: 0")
        self.set_status(error_message)
        self.palette.update_results()

    def rank_entries(self, query: str) -> list[ClientEntry]:
        scored = []
        for entry in self.entries:
            score = self._score_entry(query, entry)
            if score > 0:
                scored.append((score, entry.display_name.lower(), entry))
        scored.sort(key=lambda item: (-item[0], item[1]))
        return [item[2] for item in scored]

    def _score_entry(self, query: str, entry: ClientEntry) -> int:
        name = entry.normalized_name
        if name == query:
            return 10000
        if name.startswith(query):
            return 9000 - len(name)
        if query in name:
            return 8000 - name.index(query)
        query_tokens = query.split()
        name_tokens = name.split()
        pos = 0
        total = 0
        for token in query_tokens:
            found = False
            for idx in range(pos, len(name_tokens)):
                if name_tokens[idx].startswith(token):
                    total += 180 - min(idx * 10, 80)
                    pos = idx + 1
                    found = True
                    break
            if not found:
                total = 0
                break
        if total:
            return 7000 + total
        compact = query.replace(" ", "")
        initials = "".join(token[0] for token in name_tokens if token)
        if initials.startswith(compact):
            return 6500 - len(initials)
        similarity = SequenceMatcher(None, query, name).ratio()
        return int(similarity * 5000) if similarity >= 0.45 else 0

    def prepare_client_files(self, entry: ClientEntry) -> None:
        username = self.config_store.preferred_username.strip()
        if not username:
            return
        self._update_server_dcn_user_list(entry.folder_path, username)

    def _update_server_dcn_user_list(self, client_folder: Path, username: str) -> None:
        server_path = client_folder / SERVER_DCN_NAME
        if not server_path.exists():
            return

        try:
            content, encoding = read_text_with_fallback(server_path)
        except OSError:
            return

        lines = content.splitlines(keepends=True)
        users_section_index = None
        list_line_index = None

        for index, line in enumerate(lines):
            stripped = line.strip()
            if stripped.lower() == "[users]":
                users_section_index = index
                continue
            if users_section_index is not None and stripped.startswith("[") and stripped.endswith("]") and index > users_section_index:
                break
            if users_section_index is not None and stripped.lower().startswith("list="):
                list_line_index = index
                break

        if list_line_index is not None:
            current_line = lines[list_line_index]
            newline = "\r\n" if current_line.endswith("\r\n") else "\n"
            current_users = parse_user_list(current_line.split("=", 1)[1].strip())
            deduped = [username]
            seen = {normalize_text(username)}
            for item in current_users:
                normalized_item = normalize_text(item)
                if normalized_item and normalized_item not in seen:
                    deduped.append(item)
                    seen.add(normalized_item)
            lines[list_line_index] = f"List={serialize_user_list(deduped)}{newline}"
        elif users_section_index is not None:
            insert_index = users_section_index + 1
            newline = "\r\n"
            lines.insert(insert_index, f"List={serialize_user_list([username])}{newline}")
        else:
            if lines and not lines[-1].endswith(("\n", "\r\n")):
                lines[-1] = lines[-1] + "\r\n"
            lines.extend(["[Users]\r\n", f"List={serialize_user_list([username])}\r\n"])

        try:
            server_path.write_text("".join(lines), encoding=encoding)
            self.set_status(f"Usuario '{username}' gravado em {SERVER_DCN_NAME}.")
        except OSError:
            pass

    def launch_entry(self, entry: ClientEntry) -> None:
        if not entry.exe_path.exists():
            messagebox.showerror("Arquivo nao encontrado", f"Nao achei o arquivo:\n{entry.exe_path}")
            return
        self.prepare_client_files(entry)
        try:
            process = subprocess.Popen([str(entry.exe_path)], cwd=str(entry.folder_path))
        except OSError as exc:
            messagebox.showerror("Erro ao executar", str(exc))
            return
        self.set_status(f"Abrindo: {entry.display_name}")
        self.palette.hide_palette()
        self.login_automator.apply_async(process.pid, self.config_store.login_preferences())

    def open_root_folder(self) -> None:
        if not self.config_store.root_path.exists():
            messagebox.showerror("Pasta nao encontrada", f"Nao achei a pasta:\n{self.config_store.root_path}")
            return
        subprocess.Popen(["explorer", str(self.config_store.root_path)])

    def shutdown(self) -> None:
        if self.hotkey_listener:
            self.hotkey_listener.stop()
        if self.tray_icon:
            self.tray_icon.stop()
        self.palette.destroy()
        self.destroy()


def bring_existing_instance_to_front() -> None:
    hwnd = user32.FindWindowW(None, APP_NAME)
    if hwnd:
        user32.ShowWindow(hwnd, 5)
        user32.SetForegroundWindow(hwnd)


def acquire_single_instance_mutex():
    mutex = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if not mutex:
        return None
    if kernel32.GetLastError() == 183:
        bring_existing_instance_to_front()
        user32.MessageBoxW(None, "O Buscador Cliente HeadCargo ja esta aberto.", APP_NAME, 0x40)
        return None
    return mutex


if __name__ == "__main__":
    mutex_handle = acquire_single_instance_mutex()
    if mutex_handle:
        app = LauncherApp()
        try:
            app.mainloop()
        finally:
            kernel32.ReleaseMutex(mutex_handle)
