import base64
import ctypes
import csv
import hashlib
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
from datetime import date, datetime
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from tkinter import filedialog, messagebox


APP_NAME = "Buscador Cliente HeadCargo"
APP_VERSION = "1.2.1"
APP_DIR = Path(__file__).resolve().parent
APPDATA_DIR = Path(os.environ.get("APPDATA", str(APP_DIR))) / "BuscadorClienteHeadCargo"
CONFIG_FILE = APPDATA_DIR / "launcher_clientes_onedrive_config.json"
UPDATES_DIR = APPDATA_DIR / "updates"
CLIENT_CACHE_DIR = APPDATA_DIR / "clientes_cache"
LOGO_FILE = "logo_buscador.png"
ICON_FILE = "logo_buscador.ico"
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
APP_USER_MODEL_ID = "HeadCargo.BuscadorClienteHeadCargo"

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
IMAGE_ICON = 1
LR_LOADFROMFILE = 0x0010

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
crypt32 = ctypes.windll.crypt32
WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_ssize_t, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.DefWindowProcW.restype = ctypes.c_ssize_t
kernel32.LocalFree.argtypes = [ctypes.c_void_p]
kernel32.LocalFree.restype = ctypes.c_void_p


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


class DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_byte)),
    ]


crypt32.CryptProtectData.argtypes = [ctypes.POINTER(DATA_BLOB), wintypes.LPCWSTR, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(DATA_BLOB)]
crypt32.CryptProtectData.restype = wintypes.BOOL
crypt32.CryptUnprotectData.argtypes = [ctypes.POINTER(DATA_BLOB), ctypes.POINTER(wintypes.LPWSTR), ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(DATA_BLOB)]
crypt32.CryptUnprotectData.restype = wintypes.BOOL


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
    password: str


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


def slugify_filename(value: str) -> str:
    normalized = normalize_text(value).replace(" ", "_").strip("_")
    return normalized or "cliente"


def strip_ini_section(content: str, section_name: str) -> str:
    target_header = f"[{section_name.lower()}]"
    filtered = []
    inside_target = False
    for line in content.splitlines(keepends=True):
        stripped = line.strip()
        is_header = stripped.startswith("[") and stripped.endswith("]")
        if is_header and inside_target:
            inside_target = False
        if stripped.lower() == target_header:
            inside_target = True
            continue
        if not inside_target:
            filtered.append(line)
    return "".join(filtered)


def extract_ini_section(content: str, section_name: str) -> str:
    target_header = f"[{section_name.lower()}]"
    collected = []
    inside_target = False
    for line in content.splitlines(keepends=True):
        stripped = line.strip()
        is_header = stripped.startswith("[") and stripped.endswith("]")
        if is_header and inside_target and stripped.lower() != target_header:
            break
        if stripped.lower() == target_header:
            inside_target = True
        if inside_target:
            collected.append(line)
    return "".join(collected)


def create_data_blob(data: bytes) -> tuple[DATA_BLOB, ctypes.Array]:
    if not data:
        buffer = ctypes.create_string_buffer(1)
        blob = DATA_BLOB(0, ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)))
        return blob, buffer
    buffer = ctypes.create_string_buffer(data, len(data))
    blob = DATA_BLOB(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)))
    return blob, buffer


def protect_secret(plain_text: str) -> str:
    if not plain_text:
        return ""
    data = plain_text.encode("utf-8")
    input_blob, input_buffer = create_data_blob(data)
    output_blob = DATA_BLOB()
    if not crypt32.CryptProtectData(ctypes.byref(input_blob), APP_NAME, None, None, None, 0, ctypes.byref(output_blob)):
        raise ctypes.WinError()
    try:
        encrypted = ctypes.string_at(output_blob.pbData, output_blob.cbData)
        return base64.b64encode(encrypted).decode("ascii")
    finally:
        if output_blob.pbData:
            kernel32.LocalFree(output_blob.pbData)
        del input_buffer


def unprotect_secret(protected_text: str) -> str:
    if not protected_text:
        return ""
    encrypted = base64.b64decode(protected_text.encode("ascii"))
    input_blob, input_buffer = create_data_blob(encrypted)
    output_blob = DATA_BLOB()
    description = wintypes.LPWSTR()
    if not crypt32.CryptUnprotectData(ctypes.byref(input_blob), ctypes.byref(description), None, None, None, 0, ctypes.byref(output_blob)):
        raise ctypes.WinError()
    try:
        decrypted = ctypes.string_at(output_blob.pbData, output_blob.cbData)
        return decrypted.decode("utf-8")
    finally:
        if output_blob.pbData:
            kernel32.LocalFree(output_blob.pbData)
        if description:
            kernel32.LocalFree(description)
        del input_buffer


def load_windows_icon(filename: str, size: int = 32) -> int | None:
    icon_path = resource_path(filename)
    if not icon_path.exists():
        return None
    handle = user32.LoadImageW(None, str(icon_path), IMAGE_ICON, size, size, LR_LOADFROMFILE)
    return handle or None


def set_app_user_model_id() -> None:
    try:
        shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
    except Exception:
        pass


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
            "encrypted_password": "",
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
        return LoginPreferences(self.preferred_username, self.mark_save_password, self.saved_password)

    @property
    def encrypted_password(self) -> str:
        return str(self.data.get("encrypted_password", "")).strip()

    @encrypted_password.setter
    def encrypted_password(self, value: str) -> None:
        self.data["encrypted_password"] = value.strip()
        self.save()

    @property
    def saved_password(self) -> str:
        if not self.encrypted_password:
            return ""
        try:
            return unprotect_secret(self.encrypted_password)
        except Exception:
            return ""

    def store_password(self, plain_text: str) -> None:
        self.encrypted_password = protect_secret(plain_text)

    def clear_saved_password(self) -> None:
        self.encrypted_password = ""

    def has_saved_password(self) -> bool:
        return bool(self.saved_password)

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
        if not (preferences.username or preferences.password or preferences.mark_save_password):
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
        if password_hwnd and preferences.password:
            self._apply_password(password_hwnd, preferences.password)
        if checkbox_hwnd:
            self._set_checkbox_state(checkbox_hwnd, preferences.mark_save_password)
        if password_hwnd:
            user32.SetFocus(password_hwnd)
        summary = "Login preparado no cliente."
        if preferences.username and preferences.password:
            summary = f"Usuario '{preferences.username}' e senha local preparados no login."
        elif preferences.username:
            summary = f"Usuario '{preferences.username}' preparado no login."
        elif preferences.password:
            summary = "Senha local preparada no login."
        self.app.after(0, lambda: self.app.set_status(summary))

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
        if not username:
            return
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

    def _apply_password(self, hwnd: int, password: str) -> None:
        user32.SendMessageW(hwnd, WM_SETTEXT, 0, password)


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
        self.hicon = self.app.get_native_icon_handle(16) or user32.LoadIconW(None, IDI_APPLICATION)
        wc = WNDCLASSW()
        wc.lpfnWndProc = self.wnd_proc
        wc.hInstance = instance
        wc.lpszClassName = self.class_name
        wc.hIcon = self.app.get_native_icon_handle(32) or self.hicon
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
        self.geometry("700x720")
        self.resizable(False, False)
        self.configure(bg=BG_APP)
        self.transient(app)
        self.grab_set()
        if getattr(app, "app_icon", None):
            self.iconphoto(True, app.app_icon)
        self.path_var = tk.StringVar(value=str(app.config_store.root_path))
        self.username_var = tk.StringVar(value=app.config_store.preferred_username)
        self.password_var = tk.StringVar()
        self.password_var.trace_add("write", lambda *_: self._on_password_changed())
        self.save_password_var = tk.BooleanVar(value=app.config_store.mark_save_password)
        self.password_status_var = tk.StringVar()
        self.clear_saved_password_requested = False
        self._refresh_password_status()
        self._build_ui()
        self._fit_to_content()
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
        tk.Label(card, text="Senha padrao do HeadCargo", font=("Segoe UI Semibold", 10), fg=TEXT_MAIN, bg=BG_PANEL).pack(anchor="w")
        tk.Entry(card, textvariable=self.password_var, show="*", font=("Segoe UI", 10), relief="flat", bd=0, fg=TEXT_MAIN, bg=BG_RESULTS, insertbackground=TEXT_MAIN).pack(fill="x", pady=(6, 8), ipady=8)
        tk.Label(card, textvariable=self.password_status_var, font=("Segoe UI", 9), fg=TEXT_MUTED, bg=BG_PANEL, justify="left", anchor="w", wraplength=610).pack(fill="x")
        password_actions = tk.Frame(card, bg=BG_PANEL)
        password_actions.pack(fill="x", pady=(10, 0))
        tk.Button(password_actions, text="Limpar senha salva neste PC", command=self._mark_clear_saved_password, font=("Segoe UI", 9), relief="flat", bd=0, padx=14, pady=8, fg=TEXT_MAIN, bg="#e2e8f0", activebackground="#cbd5e1", activeforeground=TEXT_MAIN).pack(side="left")
        tk.Checkbutton(card, text="Marcar 'Salvar senha' do HeadCargo automaticamente", variable=self.save_password_var, font=("Segoe UI", 10), fg=TEXT_MAIN, bg=BG_PANEL, activebackground=BG_PANEL, activeforeground=TEXT_MAIN, selectcolor=BG_PANEL).pack(anchor="w", pady=(12, 0))
        tk.Label(card, text="A senha digitada aqui fica criptografada somente neste computador. O OneDrive do cliente nao recebe essa senha.", font=("Segoe UI", 9), fg=TEXT_MUTED, bg=BG_PANEL, justify="left", wraplength=610).pack(anchor="w", pady=(8, 0))
        tk.Label(card, text="Atualizacoes automáticas", font=("Segoe UI Semibold", 10), fg=TEXT_MAIN, bg=BG_PANEL).pack(anchor="w", pady=(16, 0))
        update_card = tk.Frame(card, bg=BG_RESULTS, padx=12, pady=12, highlightbackground=BORDER, highlightthickness=1)
        update_card.pack(fill="x", pady=(6, 0))
        tk.Label(update_card, text="Origem: GitHub / kauanlauer/Buscador-Cliente", font=("Consolas", 9), fg=TEXT_MAIN, bg=BG_RESULTS, justify="left", anchor="w").pack(fill="x")
        tk.Label(update_card, text="O programa verifica novas versoes todo dia e voce tambem pode usar o botao 'Verificar atualizacao' no painel principal.", font=("Segoe UI", 9), fg=TEXT_MUTED, bg=BG_RESULTS, justify="left", wraplength=590).pack(anchor="w", pady=(8, 0))
        buttons = tk.Frame(card, bg=BG_PANEL)
        buttons.pack(fill="x", pady=(24, 0))
        tk.Button(buttons, text="Salvar", command=self._save, font=("Segoe UI Semibold", 10), relief="flat", bd=0, padx=18, pady=10, fg="#ffffff", bg=ACCENT, activebackground=ACCENT_HOVER, activeforeground="#ffffff").pack(side="left")
        tk.Button(buttons, text="Cancelar", command=self.destroy, font=("Segoe UI", 10), relief="flat", bd=0, padx=18, pady=10, fg=TEXT_MAIN, bg="#e2e8f0", activebackground="#cbd5e1", activeforeground=TEXT_MAIN).pack(side="left", padx=(10, 0))

    def _fit_to_content(self) -> None:
        self.update_idletasks()
        width = min(max(self.winfo_reqwidth(), 700), max(self.winfo_screenwidth() - 80, 700))
        height = min(max(self.winfo_reqheight(), 720), max(self.winfo_screenheight() - 80, 560))
        self.geometry(f"{width}x{height}")

    def _center(self) -> None:
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = max((self.winfo_screenwidth() - width) // 2, 0)
        y = max((self.winfo_screenheight() - height) // 6, 20)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _select_path(self) -> None:
        selected = filedialog.askdirectory(title="Selecione a pasta raiz dos clientes", initialdir=self.path_var.get() or str(self.app.config_store.root_path))
        if selected:
            self.path_var.set(selected)

    def _refresh_password_status(self) -> None:
        if self.clear_saved_password_requested:
            self.password_status_var.set("A senha salva neste PC sera removida quando voce clicar em Salvar.")
            return
        if self.app.config_store.has_saved_password():
            self.password_status_var.set("Ja existe uma senha salva e criptografada neste PC. Se voce digitar outra senha acima, ela substitui a atual.")
        else:
            self.password_status_var.set("Nenhuma senha salva neste PC. Se voce preencher o campo acima, ela sera gravada somente nesta maquina.")

    def _mark_clear_saved_password(self) -> None:
        self.password_var.set("")
        self.clear_saved_password_requested = True
        self._refresh_password_status()

    def _on_password_changed(self) -> None:
        if self.password_var.get():
            self.clear_saved_password_requested = False
            self.password_status_var.set("Uma nova senha sera salva e criptografada neste PC quando voce clicar em Salvar.")
        elif not self.clear_saved_password_requested:
            self._refresh_password_status()

    def _save(self) -> None:
        selected_path = Path(self.path_var.get().strip())
        if not selected_path.exists():
            messagebox.showerror("Pasta invalida", "Selecione uma pasta valida.")
            return
        self.app.config_store.root_path = selected_path
        self.app.config_store.preferred_username = self.username_var.get().strip()
        password_text = self.password_var.get()
        try:
            if password_text:
                self.app.config_store.store_password(password_text)
            elif self.clear_saved_password_requested:
                self.app.config_store.clear_saved_password()
        except Exception as exc:
            messagebox.showerror("Falha ao salvar senha", str(exc))
            return
        self.app.config_store.mark_save_password = self.save_password_var.get()
        self.app.refresh_config_labels()
        self.app.refresh_index()
        self.app.set_status("Configuracoes salvas.")
        self.destroy()



class SettingsWindow(tk.Toplevel):
    def __init__(self, app) -> None:
        super().__init__(app)
        self.app = app
        self.title(f"{APP_NAME} - Configuracoes")
        self.resizable(False, False)
        self.configure(bg=BG_APP)
        self.transient(app)
        self.grab_set()
        if getattr(app, "app_icon", None):
            self.iconphoto(True, app.app_icon)
        self.path_var = tk.StringVar(value=str(app.config_store.root_path))
        self.username_var = tk.StringVar(value=app.config_store.preferred_username)
        self.password_var = tk.StringVar()
        self.password_var.trace_add("write", lambda *_: self._on_password_changed())
        self.save_password_var = tk.BooleanVar(value=app.config_store.mark_save_password)
        self.password_status_var = tk.StringVar()
        self.clear_saved_password_requested = False
        self.scroll_canvas = None
        self.scroll_window = None
        self._refresh_password_status()
        self._build_ui()
        self._fit_to_content()
        self._center()

    def _build_ui(self) -> None:
        container = tk.Frame(self, bg=BG_APP)
        container.pack(fill="both", expand=True)

        content_shell = tk.Frame(container, bg=BG_APP, padx=18, pady=18)
        content_shell.pack(fill="both", expand=True, pady=(0, 8))

        self.scroll_canvas = tk.Canvas(content_shell, bg=BG_APP, highlightthickness=0, bd=0)
        scrollbar = tk.Scrollbar(content_shell, orient="vertical", command=self.scroll_canvas.yview)
        self.scroll_canvas.configure(yscrollcommand=scrollbar.set)
        self.scroll_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        scroll_body = tk.Frame(self.scroll_canvas, bg=BG_APP)
        self.scroll_window = self.scroll_canvas.create_window((0, 0), window=scroll_body, anchor="nw")
        scroll_body.bind("<Configure>", self._update_scroll_region)
        self.scroll_canvas.bind("<Configure>", self._sync_scroll_width)
        self.bind("<MouseWheel>", self._on_mousewheel)

        card = tk.Frame(scroll_body, bg=BG_PANEL, padx=18, pady=18, highlightbackground=BORDER, highlightthickness=1)
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

        tk.Label(card, text="Senha padrao do HeadCargo", font=("Segoe UI Semibold", 10), fg=TEXT_MAIN, bg=BG_PANEL).pack(anchor="w")
        tk.Entry(card, textvariable=self.password_var, show="*", font=("Segoe UI", 10), relief="flat", bd=0, fg=TEXT_MAIN, bg=BG_RESULTS, insertbackground=TEXT_MAIN).pack(fill="x", pady=(6, 8), ipady=8)
        tk.Label(card, textvariable=self.password_status_var, font=("Segoe UI", 9), fg=TEXT_MUTED, bg=BG_PANEL, justify="left", anchor="w", wraplength=610).pack(fill="x")

        password_actions = tk.Frame(card, bg=BG_PANEL)
        password_actions.pack(fill="x", pady=(10, 0))
        tk.Button(password_actions, text="Limpar senha salva neste PC", command=self._mark_clear_saved_password, font=("Segoe UI", 9), relief="flat", bd=0, padx=14, pady=8, fg=TEXT_MAIN, bg="#e2e8f0", activebackground="#cbd5e1", activeforeground=TEXT_MAIN).pack(side="left")

        tk.Checkbutton(card, text="Marcar 'Salvar senha' do HeadCargo automaticamente", variable=self.save_password_var, font=("Segoe UI", 10), fg=TEXT_MAIN, bg=BG_PANEL, activebackground=BG_PANEL, activeforeground=TEXT_MAIN, selectcolor=BG_PANEL).pack(anchor="w", pady=(12, 0))
        tk.Label(card, text="A senha digitada aqui fica criptografada somente neste computador. O OneDrive do cliente nao recebe essa senha.", font=("Segoe UI", 9), fg=TEXT_MUTED, bg=BG_PANEL, justify="left", wraplength=610).pack(anchor="w", pady=(8, 0))

        tk.Label(card, text="Atualizacoes automaticas", font=("Segoe UI Semibold", 10), fg=TEXT_MAIN, bg=BG_PANEL).pack(anchor="w", pady=(16, 0))
        update_card = tk.Frame(card, bg=BG_RESULTS, padx=12, pady=12, highlightbackground=BORDER, highlightthickness=1)
        update_card.pack(fill="x", pady=(6, 0))
        tk.Label(update_card, text="Origem: GitHub / kauanlauer/Buscador-Cliente", font=("Consolas", 9), fg=TEXT_MAIN, bg=BG_RESULTS, justify="left", anchor="w").pack(fill="x")
        tk.Label(update_card, text="O programa verifica novas versoes todo dia e voce tambem pode usar o botao 'Verificar atualizacao' no painel principal.", font=("Segoe UI", 9), fg=TEXT_MUTED, bg=BG_RESULTS, justify="left", wraplength=590).pack(anchor="w", pady=(8, 0))

        footer = tk.Frame(container, bg=BG_APP, padx=18, pady=18)
        footer.pack(fill="x", side="bottom")
        footer_card = tk.Frame(footer, bg=BG_PANEL, padx=18, pady=14, highlightbackground=BORDER, highlightthickness=1)
        footer_card.pack(fill="x")
        tk.Button(footer_card, text="Salvar", command=self._save, font=("Segoe UI Semibold", 10), relief="flat", bd=0, padx=18, pady=10, fg="#ffffff", bg=ACCENT, activebackground=ACCENT_HOVER, activeforeground="#ffffff").pack(side="left")
        tk.Button(footer_card, text="Cancelar", command=self.destroy, font=("Segoe UI", 10), relief="flat", bd=0, padx=18, pady=10, fg=TEXT_MAIN, bg="#e2e8f0", activebackground="#cbd5e1", activeforeground=TEXT_MAIN).pack(side="left", padx=(10, 0))

    def _fit_to_content(self) -> None:
        self.update_idletasks()
        width = min(max(self.winfo_reqwidth(), 700), max(self.winfo_screenwidth() - 80, 700))
        height = min(max(self.winfo_reqheight(), 620), max(self.winfo_screenheight() - 80, 560))
        self.geometry(f"{width}x{height}")

    def _center(self) -> None:
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = max((self.winfo_screenwidth() - width) // 2, 0)
        y = max((self.winfo_screenheight() - height) // 6, 20)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _update_scroll_region(self, _event=None) -> None:
        if self.scroll_canvas:
            self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))

    def _sync_scroll_width(self, event) -> None:
        if self.scroll_canvas is not None and self.scroll_window is not None:
            self.scroll_canvas.itemconfigure(self.scroll_window, width=event.width)

    def _on_mousewheel(self, event) -> None:
        if self.scroll_canvas:
            self.scroll_canvas.yview_scroll(int(-event.delta / 120), "units")

    def _select_path(self) -> None:
        selected = filedialog.askdirectory(title="Selecione a pasta raiz dos clientes", initialdir=self.path_var.get() or str(self.app.config_store.root_path))
        if selected:
            self.path_var.set(selected)

    def _refresh_password_status(self) -> None:
        if self.clear_saved_password_requested:
            self.password_status_var.set("A senha salva neste PC sera removida quando voce clicar em Salvar.")
            return
        if self.app.config_store.has_saved_password():
            self.password_status_var.set("Ja existe uma senha salva e criptografada neste PC. Se voce digitar outra senha acima, ela substitui a atual.")
        else:
            self.password_status_var.set("Nenhuma senha salva neste PC. Se voce preencher o campo acima, ela sera gravada somente nesta maquina.")

    def _mark_clear_saved_password(self) -> None:
        self.password_var.set("")
        self.clear_saved_password_requested = True
        self._refresh_password_status()

    def _on_password_changed(self) -> None:
        if self.password_var.get():
            self.clear_saved_password_requested = False
            self.password_status_var.set("Uma nova senha sera salva e criptografada neste PC quando voce clicar em Salvar.")
        elif not self.clear_saved_password_requested:
            self._refresh_password_status()

    def _save(self) -> None:
        selected_path = Path(self.path_var.get().strip())
        if not selected_path.exists():
            messagebox.showerror("Pasta invalida", "Selecione uma pasta valida.")
            return
        self.app.config_store.root_path = selected_path
        self.app.config_store.preferred_username = self.username_var.get().strip()
        password_text = self.password_var.get()
        try:
            if password_text:
                self.app.config_store.store_password(password_text)
            elif self.clear_saved_password_requested:
                self.app.config_store.clear_saved_password()
        except Exception as exc:
            messagebox.showerror("Falha ao salvar senha", str(exc))
            return
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
        set_app_user_model_id()
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
        self.local_password_text = tk.StringVar()
        self.update_source_var = tk.StringVar()
        self.update_result_text = tk.StringVar(value="Atualizacao: aguardando primeira verificacao.")
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
        if "--minimized" in sys.argv:
            self.after(1200, self.minimize_to_taskbar)

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
        icon_path = resource_path(ICON_FILE)
        if icon_path.exists():
            try:
                self.iconbitmap(default=str(icon_path))
            except Exception:
                pass
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
        tk.Label(config_card, textvariable=self.local_password_text, font=("Segoe UI", 10), fg=TEXT_MUTED, bg=BG_PANEL, anchor="w").pack(fill="x", pady=(6, 0))
        tk.Label(config_card, text="Origem de atualizacao", font=("Segoe UI Semibold", 10), fg=TEXT_MAIN, bg=BG_PANEL).pack(anchor="w", pady=(12, 0))
        tk.Label(config_card, textvariable=self.update_source_var, font=("Consolas", 9), fg=TEXT_MUTED, bg=BG_PANEL, justify="left", anchor="w", wraplength=480).pack(fill="x", pady=(6, 0))
        tk.Label(config_card, textvariable=self.update_result_text, font=("Segoe UI", 9), fg=TEXT_SOFT, bg=BG_PANEL, justify="left", anchor="w", wraplength=480).pack(fill="x", pady=(8, 0))
        actions = tk.Frame(container, bg=BG_APP)
        actions.pack(fill="x", pady=(14, 0))
        tk.Button(actions, text="Abrir busca", command=self.show_palette, font=("Segoe UI Semibold", 10), relief="flat", bd=0, padx=16, pady=10, fg="#ffffff", bg=ACCENT, activebackground=ACCENT_HOVER, activeforeground="#ffffff").pack(side="left")
        tk.Button(actions, text="Configuracoes", command=self.open_settings, font=("Segoe UI", 10), relief="flat", bd=0, padx=16, pady=10, fg=TEXT_MAIN, bg="#e2e8f0", activebackground="#cbd5e1", activeforeground=TEXT_MAIN).pack(side="left", padx=(10, 0))
        tk.Button(actions, text="Reindexar", command=self.refresh_index, font=("Segoe UI", 10), relief="flat", bd=0, padx=16, pady=10, fg=TEXT_MAIN, bg="#e2e8f0", activebackground="#cbd5e1", activeforeground=TEXT_MAIN).pack(side="left", padx=(10, 0))
        tk.Button(actions, text="Verificar atualizacao", command=self.check_for_updates, font=("Segoe UI", 10), relief="flat", bd=0, padx=16, pady=10, fg=TEXT_MAIN, bg="#e2e8f0", activebackground="#cbd5e1", activeforeground=TEXT_MAIN).pack(side="left", padx=(10, 0))
        hint_card = tk.Frame(container, bg=BG_PANEL, padx=20, pady=16, highlightbackground=BORDER, highlightthickness=1)
        hint_card.pack(fill="both", expand=True, pady=(14, 0))
        tk.Label(hint_card, text="Como usar", font=("Segoe UI Semibold", 10), fg=TEXT_MAIN, bg=BG_PANEL).pack(anchor="w")
        tk.Label(hint_card, text="1. Deixe o buscador aberto.\n2. Use Alt+Espaco.\n3. Digite o cliente e pressione Enter.\n4. O programa abre uma copia local do cliente, preenche o usuario e pode preencher a senha deste PC.", font=("Segoe UI", 10), fg=TEXT_MUTED, bg=BG_PANEL, justify="left", anchor="w").pack(fill="x", pady=(8, 8))
        tk.Label(hint_card, textvariable=self.status_var, font=("Segoe UI", 9), fg=TEXT_SOFT, bg=BG_PANEL, justify="left", anchor="w").pack(fill="x")

    def refresh_config_labels(self) -> None:
        self.folder_var.set(str(self.config_store.root_path))
        self.user_var.set(self.config_store.preferred_username or "Nao definido")
        if self.config_store.mark_save_password:
            self.save_password_text.set("Salvar senha do HeadCargo: marcar automaticamente na copia local do cliente.")
        else:
            self.save_password_text.set("Salvar senha do HeadCargo: nao marcar automaticamente.")
        self.local_password_text.set("Senha local neste PC: salva e criptografada." if self.config_store.has_saved_password() else "Senha local neste PC: nao salva.")
        self.update_source_var.set("GitHub: kauanlauer/Buscador-Cliente")

    def get_native_icon_handle(self, size: int = 32) -> int | None:
        return load_windows_icon(ICON_FILE, size)

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
                messagebox.showinfo("Atualizacao", "Ja existe uma verificacao de atualizacao em andamento. Aguarde alguns segundos.")
            return
        self.is_checking_updates = True
        self.set_status("Verificando atualizacoes no GitHub...")
        self.update_result_text.set("Atualizacao: consultando o GitHub...")
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
            checked_at = datetime.now().strftime("%d/%m/%Y %H:%M")
            message = f"GitHub verificado com sucesso em {checked_at}. Versao instalada {APP_VERSION}. Versao no repositorio {latest_version}."
            self.set_status(message)
            self.update_result_text.set(message)
            if not silent_no_update:
                messagebox.showinfo("Atualizacao", f"Repositorio lido com sucesso.\n\nVersao instalada: {APP_VERSION}\nVersao no GitHub: {latest_version}\n\nSeu programa ja esta atualizado.")
            return

        if not getattr(sys, "frozen", False):
            self.update_result_text.set(f"Nova versao encontrada no GitHub: {latest_version}.")
            messagebox.showinfo(
                "Atualizacao disponivel",
                f"Nova versao encontrada: {latest_version}\n\nPara testar a atualizacao automatica, abra o .exe instalado.",
            )
            return

        details = f"Versao atual: {APP_VERSION}\nNova versao: {latest_version}"
        if notes:
            details += f"\n\nNovidades:\n{notes}"
        self.update_result_text.set(f"Nova versao encontrada no GitHub: {latest_version}.")
        confirm = messagebox.askyesno("Atualizacao disponivel", f"{details}\n\nDeseja atualizar agora?")
        if confirm:
            self.download_and_apply_update(installer_url, latest_version)
        else:
            self.set_status(f"Atualizacao {latest_version} adiada pelo usuario.")
            self.update_result_text.set(f"Atualizacao {latest_version} encontrada no GitHub e adiada pelo usuario.")

    def _handle_update_error(self, error_message: str, silent_errors: bool) -> None:
        self.is_checking_updates = False
        self.set_status("Nao foi possivel verificar atualizacoes.")
        self.update_result_text.set("Atualizacao: falha ao ler o repositorio do GitHub.")
        if not silent_errors:
            messagebox.showerror("Falha ao verificar atualizacao", error_message)

    def download_and_apply_update(self, installer_url: str, target_version: str) -> None:
        self.set_status(f"Baixando atualizacao {target_version}...")
        self.update_result_text.set(f"Baixando a versao {target_version} direto do GitHub...")
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
            self.after(0, lambda: self.update_result_text.set("Atualizacao: falha ao baixar o instalador do GitHub."))

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
        self.update_result_text.set(f"Atualizacao iniciada para a versao {target_version}.")
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

    def prepare_client_workspace(self, entry: ClientEntry) -> tuple[Path, Path]:
        source_folder = entry.folder_path
        local_folder = self._local_client_folder(source_folder, entry.display_name)
        self._sync_client_folder(source_folder, local_folder)
        username = self.config_store.preferred_username.strip()
        if username:
            self._update_server_dcn_user_list(local_folder, username)
        relative_exe = entry.exe_path.relative_to(source_folder)
        local_exe = local_folder / relative_exe
        if not local_exe.exists():
            raise FileNotFoundError(f"Nao achei o Desktop.exe local em: {local_exe}")
        return local_folder, local_exe

    def _local_client_folder(self, source_folder: Path, display_name: str) -> Path:
        folder_hash = hashlib.sha1(str(source_folder).encode("utf-8")).hexdigest()[:10]
        local_name = f"{slugify_filename(display_name)}_{folder_hash}"
        target = CLIENT_CACHE_DIR / local_name
        target.mkdir(parents=True, exist_ok=True)
        return target

    def _sync_client_folder(self, source_folder: Path, local_folder: Path) -> None:
        for current_root, dirnames, filenames in os.walk(source_folder):
            current_root_path = Path(current_root)
            relative_root = current_root_path.relative_to(source_folder)
            target_root = local_folder / relative_root
            target_root.mkdir(parents=True, exist_ok=True)
            for dirname in dirnames:
                (target_root / dirname).mkdir(parents=True, exist_ok=True)
            for filename in filenames:
                source_file = current_root_path / filename
                target_file = target_root / filename
                if source_file.name.lower() == SERVER_DCN_NAME.lower():
                    self._sync_local_server_dcn(source_file, target_file)
                    continue
                if self._should_copy_file(source_file, target_file):
                    shutil.copy2(source_file, target_file)

    def _should_copy_file(self, source_file: Path, target_file: Path) -> bool:
        if not target_file.exists():
            return True
        source_stat = source_file.stat()
        target_stat = target_file.stat()
        return source_stat.st_size != target_stat.st_size or int(source_stat.st_mtime) > int(target_stat.st_mtime)

    def _sync_local_server_dcn(self, source_file: Path, target_file: Path) -> None:
        try:
            source_content, encoding = read_text_with_fallback(source_file)
        except OSError:
            return
        sanitized = strip_ini_section(source_content, "SaveData").rstrip()
        if target_file.exists():
            try:
                local_content, _ = read_text_with_fallback(target_file)
            except OSError:
                local_content = ""
            local_save_data = extract_ini_section(local_content, "SaveData").strip()
            if local_save_data:
                sanitized = f"{sanitized}\r\n{local_save_data}\r\n"
            else:
                sanitized = f"{sanitized}\r\n"
        else:
            sanitized = f"{sanitized}\r\n"
        target_file.parent.mkdir(parents=True, exist_ok=True)
        target_file.write_text(sanitized, encoding=encoding)

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
            self.set_status(f"Usuario '{username}' preparado na copia local de {SERVER_DCN_NAME}.")
        except OSError:
            pass

    def launch_entry(self, entry: ClientEntry) -> None:
        if not entry.exe_path.exists():
            messagebox.showerror("Arquivo nao encontrado", f"Nao achei o arquivo:\n{entry.exe_path}")
            return
        try:
            local_folder, local_exe = self.prepare_client_workspace(entry)
            process = subprocess.Popen([str(local_exe)], cwd=str(local_folder))
        except (OSError, FileNotFoundError) as exc:
            messagebox.showerror("Erro ao executar", str(exc))
            return
        self.set_status(f"Abrindo copia local segura de: {entry.display_name}")
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
