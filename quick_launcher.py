"""
quick_launcher.py - A simple and customizable quick launcher application for Windows.

author: chin shinrei
version: 1.0.0
date: 2025-6-30
"""

import os
import sys
import json
import tkinter as tk
from tkinter import messagebox, simpledialog, colorchooser, ttk, font as tkfont
from PIL import Image, ImageTk, ImageDraw, ImageFont
import io
import webbrowser
import threading
import configparser
import winreg
import ctypes
from ctypes import wintypes
from urllib.parse import urlparse
import requests
import pythoncom
from pystray import Icon, Menu, MenuItem as item
from win32com.client import Dispatch
import logging
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import shutil
import subprocess
import copy

# --- アプリケーション設定 ---
logging.basicConfig(filename='app_errors.log', level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')

if getattr(sys, 'frozen', False):
    # PyInstallerでパッケージ化された場合
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # スクリプトとして実行されている場合
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")
LINKS_FILE = os.path.join(BASE_DIR, "links.json")
PROFILES_DIR = os.path.join(BASE_DIR, "profiles")
DEFAULT_PROFILE_NAME = "(default)"

DEFAULT_SETTINGS = {
    'font': 'Yu Gothic UI',
    'size': 11,
    'font_color': '#000000',
    'bg': '#f0f0f0',
    'border_color': '#666666', # デフォルトのボーダー色
    'use_online_favicon': False,
    "current_profile": "(default)"
}

# --- ユーティリティ関数 ---
def get_profile_path(profile_name):
    """指定されたプロファイルのディレクトリパスを返す。なければ作成する。"""
    path = os.path.join(PROFILES_DIR, profile_name)
    os.makedirs(path, exist_ok=True)
    return path

def get_all_profile_names():
    """profilesディレクトリ内のすべてのプロファイル名（ディレクトリ名）を取得する"""
    if not os.path.isdir(PROFILES_DIR):
        return []
    return sorted([d for d in os.listdir(PROFILES_DIR) if os.path.isdir(os.path.join(PROFILES_DIR, d))])

# --- ctypesのグローバル定義 ---
shell32 = ctypes.windll.shell32
user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32
gdi32.DeleteObject.restype = wintypes.BOOL

class RECT(ctypes.Structure):
    _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG),
                ("right", wintypes.LONG), ("bottom", wintypes.LONG)]

# SystemParametersInfoWで作業領域を取得するための定義
SPI_GETWORKAREA = 0x0030
user32.SystemParametersInfoW.argtypes = [wintypes.UINT, wintypes.UINT, ctypes.c_void_p, wintypes.UINT]
user32.SystemParametersInfoW.restype = wintypes.BOOL

class SHFILEINFO(ctypes.Structure):
    _fields_ = [("hIcon", wintypes.HICON), ("iIcon", ctypes.c_int), ("dwAttributes", wintypes.DWORD),
                ("szDisplayName", ctypes.c_wchar * 260), ("szTypeName", ctypes.c_wchar * 80)]

class ICONINFO(ctypes.Structure):
    _fields_ = [('fIcon', wintypes.BOOL), ('xHotspot', wintypes.DWORD), ('yHotspot', wintypes.DWORD),
                ('hbmMask', wintypes.HBITMAP), ('hbmColor', wintypes.HBITMAP)]

class BITMAP(ctypes.Structure):
    _fields_ = [('bmType', wintypes.LONG), ('bmWidth', wintypes.LONG), ('bmHeight', wintypes.LONG),
                ('bmWidthBytes', wintypes.LONG), ('bmPlanes', wintypes.WORD), ('bmBitsPixel', wintypes.WORD),
                ('bmBits', wintypes.LPVOID)]

shell32.SHGetFileInfoW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, ctypes.POINTER(SHFILEINFO), wintypes.UINT, wintypes.UINT]
shell32.SHGetFileInfoW.restype = wintypes.HANDLE
user32.GetIconInfo.argtypes = [wintypes.HICON, ctypes.POINTER(ICONINFO)]
user32.GetIconInfo.restype = wintypes.BOOL
user32.DestroyIcon.argtypes = [wintypes.HICON]
user32.DestroyIcon.restype = wintypes.BOOL
gdi32.GetObjectW.argtypes = [wintypes.HANDLE, wintypes.INT, wintypes.LPVOID]
gdi32.GetObjectW.restype = wintypes.INT
gdi32.DeleteObject.argtypes = [wintypes.HANDLE]
gdi32.DeleteObject.restype = wintypes.BOOL

# --- グローバルキャッシュ ---
_icon_cache = {}
_system_icon_cache = {}
_default_browser_icon = {}  # サイズごとにキャッシュ

# --- ヘルパー関数 ---

ICON_BASE64 = """\
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\
"""

def get_work_area():
    """
    Windows APIを呼び出し、タスクバーなどを除いたデスクトップの作業領域を取得する。
    戻り値: (left, top, right, bottom) のタプル
    """
    work_area_rect = RECT()
    # SystemParametersInfoWを呼び出して作業領域をwork_area_rectに格納
    if user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(work_area_rect), 0):
        return work_area_rect.left, work_area_rect.top, work_area_rect.right, work_area_rect.bottom
    else:
        # API呼び出しに失敗した場合のフォールバック
        screen_w = user32.GetSystemMetrics(0) # SM_CXSCREEN
        screen_h = user32.GetSystemMetrics(1) # SM_CYSCREEN
        return 0, 0, screen_w, screen_h

def lighten_color(hex_color, amount=0.4):
    """
    16進数の色コードを指定された量だけ明るくする。
    amount: 0.0 (元の色) から 1.0 (白) までの値。
    """
    try:
        # '#'を削除し、RGBの各値を取得
        hex_color = hex_color.lstrip('#')
        r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        
        # 各色成分を白に近づける
        # (255 - 元の色) * amount で白までの距離の何割進むかを計算
        r = int(r + (255 - r) * amount)
        g = int(g + (255 - g) * amount)
        b = int(b + (255 - b) * amount)
        
        # 0-255の範囲に収める
        r = max(0, min(255, r))
        g = max(0, min(255, g))
        b = max(0, min(255, b))
        
        # 新しい色を16進数コードで返す
        return f'#{r:02x}{g:02x}{b:02x}'
    except Exception:
        # 解析に失敗した場合は、フォールバックとしてグレーを返す
        return "#a0a0a0"

def round_to_step(value, step=4):
    """数値を指定されたステップに丸める（例: 15をstep=4で16に）"""
    return step * round(value / step)

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                return {**DEFAULT_SETTINGS, **json.load(f)}
        except Exception as e:
            logging.warning(f"Failed to load settings: {e}")
            # 読み込み失敗時は、デフォルト設定でファイルを作成し直す
            save_settings(DEFAULT_SETTINGS)
            return DEFAULT_SETTINGS.copy()
    else:
        # ファイルが存在しない場合は、デフォルト設定で新規作成する
        logging.info("Settings file not found, creating a new one.")
        save_settings(DEFAULT_SETTINGS)
        return DEFAULT_SETTINGS.copy()

def save_settings(settings):
    with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)

def load_links_data(profile_name=DEFAULT_PROFILE_NAME):
    profile_dir = get_profile_path(profile_name)
    links_file = os.path.join(profile_dir, "links.json")
    if os.path.exists(links_file):
        try:
            with open(links_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, list) and data and isinstance(data[0], dict) and 'group' in data[0]:
                    return data
                elif isinstance(data, list):
                    return [{"group": "マイリンク", "links": data}]
        except Exception as e:
            logging.warning(f"Failed to load links data file, creating a new one: {e}")
            # 読み込み失敗時は、デフォルトの空データでファイルを作成し直す
            default_links = [{"group": "マイリンク", "links": []}]
            save_links_data(default_links)
            return default_links
    else:
        default_links = [{"group": "マイリンク", "links": []}]
        save_links_data(default_links)
        return default_links

def save_links_data(links, profile_name=DEFAULT_PROFILE_NAME):
    profile_dir = get_profile_path(profile_name)
    links_file = os.path.join(profile_dir, "links.json")
    with open(links_file, 'w', encoding='utf-8') as f:
        json.dump(links, f, ensure_ascii=False, indent=2)

def open_link(path):
    try:
        if path.startswith(('http://', 'https://')):
            webbrowser.open(path)
        else:
            # コマンドライン引数付きの場合は subprocess で実行
            subprocess.Popen(path, shell=True)
    except Exception as e:
        logging.error(f"Failed to open link '{path}': {e}")
        messagebox.showerror("リンクエラー", f"リンクを開けませんでした:\n{path}")

def _create_fallback_icon(size=20):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse((2, 2, size - 3, size - 3), outline="#888", width=1)
    return ImageTk.PhotoImage(img)

# --- アイコン取得ロジック ---
def _hicon_to_photoimage(hIcon, size, destroy_after=True):
    """
    アイコンハンドル(HICON)をTkinter PhotoImageに変換する。
    どんなサイズのHICONでも、要求されたsizeで正しく描画・変換する。
    """
    tk_icon = None
    hdc = user32.GetDC(None)
    mem_dc = gdi32.CreateCompatibleDC(hdc)
    # ★★★修正点1: 作成するビットマップのサイズを、元のアイコンサイズではなく、
    #             目標の `size` にする。
    mem_bmp = gdi32.CreateCompatibleBitmap(hdc, size, size)
    gdi32.SelectObject(mem_dc, mem_bmp)

    try:
        # 背景を透明にするための準備（マゼンタで塗りつぶし）
        # この手法により、アルファチャンネルがないアイコンでも透過背景で描画できる
        brush = gdi32.CreateSolidBrush(0x00FF00FF) # BGR形式のマゼンタ
        rect_fill = RECT(0, 0, size, size)
        user32.FillRect(mem_dc, ctypes.byref(rect_fill), brush)
        gdi32.DeleteObject(brush)

        user32.DrawIconEx(mem_dc, 0, 0, hIcon, size, size, 0, None, 3)

        # メモリDCからビットマップデータを取得
        bmp_str = ctypes.create_string_buffer(size * size * 4)
        gdi32.GetBitmapBits(mem_bmp, len(bmp_str), bmp_str)

        # Pillow Imageに変換
        img = Image.frombuffer("RGBA", (size, size), bmp_str, "raw", "BGRA", 0, 1)

        # マゼンタの背景を透過に変換
        img = img.convert("RGBA")
        datas = img.getdata()
        new_data = []
        for item in datas:
            # itemは (R, G, B, A)
            if item[0] == 255 and item[1] == 0 and item[2] == 255:
                new_data.append((255, 255, 255, 0)) # 透明ピクセル
            else:
                new_data.append(item)
        img.putdata(new_data)
        
        tk_icon = ImageTk.PhotoImage(img)
        
    finally:
        # リソースの解放
        gdi32.DeleteObject(mem_bmp)
        gdi32.DeleteDC(mem_dc)
        user32.ReleaseDC(None, hdc)
        if destroy_after:
            user32.DestroyIcon(hIcon)
            
    return tk_icon

def get_system_folder_icon(size=16):
    """
    SHGetFileInfoWを使い、システムの標準フォルダアイコンを取得する。
    ヘルパー関数のバグ修正により、この方法が最も確実であると再確認された最終版。
    """
    key = ('folder_icon_shgfi', size) # キャッシュキーを明確に
    if key in _system_icon_cache:
        return _system_icon_cache[key]

    # --- 定義 ---
    SHGFI_ICON = 0x100
    SHGFI_SMALLICON = 0x1
    SHGFI_USEFILEATTRIBUTES = 0x10
    FILE_ATTRIBUTE_DIRECTORY = 0x10

    info = SHFILEINFO()
    flags = SHGFI_ICON | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES
    
    # まず実在するディレクトリ（C:\\Windows）で取得を試みる
    tk_icon = None
    try:
        res = shell32.SHGetFileInfoW(
            r"C:\\Windows",
            FILE_ATTRIBUTE_DIRECTORY,
            ctypes.byref(info),
            ctypes.sizeof(info),
            flags
        )
        if res and info.hIcon:
            tk_icon = _hicon_to_photoimage(info.hIcon, size, destroy_after=True)
    except Exception:
        tk_icon = None
    # 失敗時は従来通りダミー名で再試行
    if tk_icon is None:
        try:
            res = shell32.SHGetFileInfoW(
                "dummy_folder",
                FILE_ATTRIBUTE_DIRECTORY,
                ctypes.byref(info),
                ctypes.sizeof(info),
                flags
            )
            if res and info.hIcon:
                tk_icon = _hicon_to_photoimage(info.hIcon, size, destroy_after=True)
        except Exception:
            tk_icon = None
    # それでも失敗した場合はダミー
    if tk_icon is None:
        tk_icon = _create_fallback_icon(size)

    _system_icon_cache[key] = tk_icon
    return tk_icon

def get_system_warning_icon(size=16):
    """
    システムの標準的な「警告」アイコンを取得する。
    """
    key = ('warning_icon', size)
    if key in _system_icon_cache:
        return _system_icon_cache[key]
    
    tk_icon = None
    hIcon = 0
    try:
        # LoadIconWを使い、標準の警告アイコン(IDI_WARNING)を要求
        # IDI_WARNING = 32515
        hIcon = user32.LoadIconW(None, 32515)
        if hIcon:
            # 標準アイコンなのでハンドルは破棄しない
            tk_icon = _hicon_to_photoimage(hIcon, size, destroy_after=False)
    except Exception as e:
        logging.warning(f"Failed to get system warning icon: {e}")

    if tk_icon is None:
        tk_icon = _create_fallback_icon(size) # フォールバックも統一

    _system_icon_cache[key] = tk_icon
    return tk_icon

def get_file_icon(path, size=16):
    """ファイルパスからアイコンを取得する。パスが存在しない場合は警告アイコンを返す。"""

    # サイズに応じたフラグを決定する
    if size > 20:
        flags = 0x100 | 0x0
    else:
        flags = 0x100 | 0x1
    key = (path, size, flags)
    # まず、キャッシュを確認
    if key in _icon_cache:
        return _icon_cache[key]

    # ファイル/フォルダの存在を確認
    file_exists = os.path.exists(path)
    # ファイル・フォルダが存在しない場合は警告アイコンを返す
    if not file_exists:
        tk_icon = get_system_warning_icon(size)
        _icon_cache[key] = tk_icon
        return tk_icon

    info = SHFILEINFO()
    res = shell32.SHGetFileInfoW(path, 0, ctypes.byref(info), ctypes.sizeof(info), flags)

    tk_icon = None
    if res and info.hIcon:
        tk_icon = _hicon_to_photoimage(info.hIcon, size)
    else:
        # exeの場合はExtractIconExで直接抽出
        if path.lower().endswith('.exe'):
            try:
                large = ctypes.c_void_p()
                small = ctypes.c_void_p()
                num_icons = shell32.ExtractIconExW(path, 0, ctypes.byref(large), ctypes.byref(small), 1)
                hIcon = None
                if size > 20 and large.value:
                    hIcon = large.value
                elif small.value:
                    hIcon = small.value
                if hIcon:
                    tk_icon = _hicon_to_photoimage(hIcon, size, destroy_after=True)
            except Exception as e:
                logging.info(f"[get_file_icon] ExtractIconEx failed: {path}: {e}")
        # 拡張子からアイコン取得（jpg, mp4, txt, pdf等）
        if tk_icon is None:
            SHGFI_ICON = 0x100
            SHGFI_SMALLICON = 0x1
            SHGFI_USEFILEATTRIBUTES = 0x10
            ext = os.path.splitext(path)[1]
            if ext:
                dummy_name = f"dummy{ext}"
                attr = 0x80  # FILE_ATTRIBUTE_NORMAL
                flags2 = SHGFI_ICON | (SHGFI_SMALLICON if size <= 20 else 0) | SHGFI_USEFILEATTRIBUTES
                info2 = SHFILEINFO()
                res2 = shell32.SHGetFileInfoW(dummy_name, attr, ctypes.byref(info2), ctypes.sizeof(info2), flags2)
                if res2 and info2.hIcon:
                    tk_icon = _hicon_to_photoimage(info2.hIcon, size, destroy_after=True)
    if tk_icon is None:
        tk_icon = get_system_folder_icon(size)

    _icon_cache[key] = tk_icon
    return tk_icon

def get_web_icon(url, size=16):
    """
    URLからファビコンを取得する。
    settingsに応じてオンライン(Google)/オフライン(直接取得)を切り替える。
    """
    global _default_browser_icon

    if not url:
        # URLが無効な場合は、デフォルトブラウザアイコンを返すしかない
        if size not in _default_browser_icon:
            _default_browser_icon[size] = _get_or_create_default_browser_icon(size)
        return _default_browser_icon[size]

    domain = urlparse(url).netloc
    key = (domain, size)
    if key in _icon_cache: return _icon_cache[key]

    tk_icon = None
    
    # 1. オンラインモードを試す (設定がTrueの場合)
    if settings.get('use_online_favicon', True):
        try:
            response = requests.get(f"https://www.google.com/s2/favicons?domain={domain}&sz=32", timeout=2)
            response.raise_for_status()
            if response.content and len(response.content) > 100: # Googleのデフォルトアイコンでないことを確認
                img = Image.open(io.BytesIO(response.content)).convert("RGBA")
                tk_icon = ImageTk.PhotoImage(img.resize((size, size), Image.LANCZOS))
        except requests.RequestException:
            # オンラインでの取得に失敗した場合、オフラインモードにフォールバック
            logging.info(f"Online favicon fetch failed for {domain}, falling back to offline mode.")
            pass 
    
    # 2. オフラインモード (またはオンラインが失敗した場合) で、tk_iconがまだNoneなら実行
    if tk_icon is None:
        try:
            # イントラネット向けに証明書検証を無効にするオプションも考慮
            headers = {'User-Agent': 'Mozilla/5.0'}
            response = requests.get(url, headers=headers, timeout=3, verify=False)
            response.raise_for_status()

            # HTMLから<link>タグを探す
            soup = BeautifulSoup(response.text, 'html.parser')
            icon_link = soup.find('link', rel='icon') or soup.find('link', rel='shortcut icon')
            
            favicon_url = None
            if icon_link and icon_link.get('href'):
                # 見つかったhrefを絶対URLに変換
                favicon_url = urljoin(url, icon_link['href'])
            else:
                # 見つからなければ、ドメインルートのfavicon.icoを試す
                favicon_url = urljoin(url, '/favicon.ico')
            
            # ファビコン画像をダウンロード
            fav_response = requests.get(favicon_url, headers=headers, timeout=3, verify=False)
            fav_response.raise_for_status()
            if fav_response.content:
                img = Image.open(io.BytesIO(fav_response.content)).convert("RGBA")
                tk_icon = ImageTk.PhotoImage(img.resize((size, size), Image.LANCZOS))

        except Exception as e:
            # オフライン取得でも失敗した場合
            logging.info(f"Offline favicon fetch failed for {url}: {e}")
            pass

    # 3. 最終フォールバック
    if tk_icon is None:
        if size not in _default_browser_icon:
            _default_browser_icon[size] = _get_or_create_default_browser_icon(size)
        tk_icon = _default_browser_icon[size]

    _icon_cache[key] = tk_icon
    return tk_icon

def _get_or_create_default_browser_icon(size=16):
    """内部用のヘルパー。デフォルトブラウザアイコンを取得、失敗時は警告アイコン。"""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice") as key:
            prog_id = winreg.QueryValueEx(key, "ProgId")[0]
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"{prog_id}\shell\open\command") as key:
            cmd_path = winreg.QueryValueEx(key, "")[0]
        browser_path = cmd_path.split('"')[1]
        if os.path.exists(browser_path):
            return get_file_icon(browser_path, size)
        else:
            return get_system_warning_icon(size)
    except Exception:
        return get_system_warning_icon(size)
    
# --- GUIクラス ---

class ToolTip:
    def __init__(self, widget):
        self.widget = widget
        self.tip_window = None
        self.id = None
        self.x = self.y = 0

    def showtip(self, text):
        "Display text in tooltip window"
        self.text = text
        if self.tip_window or not self.text:
            return
        
        bbox_result = self.widget.bbox("current")
        if bbox_result is None:
            # マウスカーソル下にアイテムがなければ、ツールチップは表示しない
            return
        x, y, _, _ = bbox_result # "current"はCanvas上のマウスカーソル位置のアイテム
        x = x + self.widget.winfo_rootx() + 25
        y = y + self.widget.winfo_rooty() + 20
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tip_window
        self.tip_window = None
        if tw:
            tw.destroy()

class SettingsDialog(simpledialog.Dialog):
    def __init__(self, parent, current_settings):
        self.settings = current_settings.copy()
        super().__init__(parent, title="設定")

    def body(self, master):
        # ウィンドウの幅と高さをリサイズ不可に設定
        self.winfo_toplevel().resizable(False, False)

        # アイコン設定
        if 'app_icon' in globals() and app_icon:
            self.iconphoto(True, app_icon)
        font_families = sorted(set(tkfont.families()))
        size_list = [str(s) for s in range(8, 33)]
        
        tk.Label(master, text="フォント名").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.font_var = tk.StringVar(value=self.settings['font'])
        self.font_combo = ttk.Combobox(master, textvariable=self.font_var, values=font_families, state="readonly")
        self.font_combo.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        tk.Label(master, text="フォントサイズ").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.size_var = tk.StringVar(value=str(self.settings['size']))
        self.size_combo = ttk.Combobox(master, textvariable=self.size_var, values=size_list, state="readonly", width=5)
        self.size_combo.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        tk.Label(master, text="フォント色").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.font_color_btn = tk.Button(master, text=" 　　　　 ", bg=self.settings['font_color'], command=self.choose_font_color)
        self.font_color_btn.grid(row=2, column=1, sticky="w", padx=5, pady=5)

        tk.Label(master, text="背景色").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.bg_color_btn = tk.Button(master, text=" 　　　　 ", bg=self.settings['bg'], command=self.choose_bg_color)
        self.bg_color_btn.grid(row=3, column=1, sticky="w", padx=5, pady=5)

        tk.Label(master, text="ボーダー色").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.border_color_btn = tk.Button(master, text=" 　　　　 ", bg=self.settings['border_color'], command=self.choose_border_color)
        self.border_color_btn.grid(row=4, column=1, sticky="w", padx=5, pady=5)

        self.online_favicon_var = tk.BooleanVar(value=self.settings.get('use_online_favicon', True))
        favicon_check = tk.Checkbutton(master, text="Webサイトのアイコンをオンラインで取得する", variable=self.online_favicon_var)
        favicon_check.grid(row=5, column=0, columnspan=2, sticky="w", padx=5, pady=10)

        self.default_btn = tk.Button(master, text="デフォルトに戻す", command=self.reset_default)
        self.default_btn.grid(row=6, column=0, columnspan=2, pady=10)

        return self.font_combo

    def choose_border_color(self):
        color = colorchooser.askcolor(initialcolor=self.border_color_btn.cget('bg'))[1]
        if color:
            self.border_color_btn.config(bg=color)

    def choose_font_color(self):
        color = colorchooser.askcolor(initialcolor=self.font_color_btn.cget('bg'))[1]
        if color:
            self.font_color_btn.config(bg=color)

    def choose_bg_color(self):
        color = colorchooser.askcolor(initialcolor=self.bg_color_btn.cget('bg'))[1]
        if color:
            self.bg_color_btn.config(bg=color)

    def ok(self, event=None):
        self.settings['font'] = self.font_var.get()
        self.settings['size'] = int(self.size_var.get())
        self.settings['font_color'] = self.font_color_btn.cget('bg')
        self.settings['bg'] = self.bg_color_btn.cget('bg')
        self.settings['border_color'] = self.border_color_btn.cget('bg')
        self.settings['use_online_favicon'] = self.online_favicon_var.get()
        self.result = self.settings.copy()
        save_settings(self.result)
        super().ok()

    def reset_default(self):
        self.settings = DEFAULT_SETTINGS.copy()
        self.font_var.set(self.settings['font'])
        self.size_var.set(str(self.settings['size']))
        self.font_color_btn.config(bg=self.settings['font_color'])
        self.bg_color_btn.config(bg=self.settings['bg'])
        self.border_color_btn.config(bg=self.settings['border_color'])
        self.online_favicon_var.set(self.settings['use_online_favicon'])

# --- リンク編集画面 ---
class LinksEditDialog(tk.Toplevel):

    def __init__(self, parent, groups, settings):
        super().__init__(parent)

        # ★最重要ポイント1: transient を削除
        # 親(root)が非表示のため、transient を設定するとこのウィンドウも非表示になってしまう。
        # self.transient(parent) 

        self.title("リンクの編集")
        if 'app_icon' in globals() and app_icon:
            self.iconphoto(True, app_icon)

        # 元データをディープコピーして保持
        self.groups = [dict(g, links=[dict(l) for l in g.get('links', [])]) for g in groups]
        self.original_groups = [dict(g, links=[dict(l) for l in g.get('links', [])]) for g in groups]

        self.settings = settings
        self.link_row_height = 24 # デフォルト値
        self.link_icon_size = 16 # デフォルト値
        self.selected_group = 0
        self.selected_link = None
        self.icon_refs = []
        self.result = None
        self.modified = False  # 変更フラグ
        self.is_searching = False

        # ウィンドウクローズ時のハンドラ
        self.protocol("WM_DELETE_WINDOW", self.cancel)

        self.resizable(True, True)
        self.minsize(640, 420)

        # --- レイアウト設定 ---
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        main_pane = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_pane.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        # --- フォント定義 ---
        title_font = tkfont.Font(family=self.settings['font'], size=self.settings['size'] - 1)
        content_font = tkfont.Font(family=self.settings['font'], size=self.settings['size'])

        # =================================================================
        # === 左ペイン：グループ一覧 ===
        # =================================================================
        group_outer_frame = tk.Frame(main_pane, width=150)  # PanedWindow用コンテナ
        group_outer_frame.pack_propagate(False) # ★重要: このフレームが縮まないようにする
        group_outer_frame.grid_rowconfigure(0, weight=1)
        group_outer_frame.grid_columnconfigure(0, weight=1)

        group_pane = tk.LabelFrame(group_outer_frame, text="グループ", font=title_font, bd=1, padx=5, pady=5)
        group_pane.grid(row=0, column=0, sticky="nsew")
        
        # LabelFrame内部のGrid設定
        group_pane.grid_columnconfigure(0, weight=1)
        group_pane.grid_rowconfigure(1, weight=1) # Listboxの行を伸縮させる

        # --- 検索ボックス ---
        #search_frame = tk.Frame(group_pane, bd=1, relief=tk.SOLID, borderwidth=1)
        # 親フレームのデフォルト背景色を取得する
        try:
            # まずはEntryのスタイルから色を取得しようと試みる
            style = ttk.Style(self)
            bg_color = style.lookup('TEntry', 'fieldbackground')
            if not bg_color or bg_color in ("", "transparent"):
                # ttkのテーマによっては""や"transparent"が返る場合があるので、親ウィジェットの背景色を使う
                bg_color = self.cget('background')
        except tk.TclError:
            # 失敗した場合は、フォールバックとしてウィンドウの標準背景色を取得
            bg_color = self.cget('background')

        border_color = lighten_color(bg_color, 0.2)
        search_frame = tk.Frame(group_pane, bg=bg_color, highlightbackground=border_color, highlightthickness=1) 
        search_frame.grid(row=0, column=0, sticky="ew", pady=(2, 6))
        
        search_icon_label = tk.Label(search_frame, text="🔍", font=("Segoe UI Symbol", self.settings['size']))
        search_icon_label.pack(side="left", padx=(5, 0))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, style='Search.TEntry')
        self.search_entry.pack(side="left", fill="both", expand=True)
        
        style = ttk.Style(self)
        style.configure('Search.TEntry', borderwidth=0, relief='flat')

        search_frame.config(bg=bg_color)
        search_icon_label.config(bg=bg_color)
        self.search_var.trace_add("write", self._on_search_change)

        # --- グループリストボックス ---
        group_list_frame = tk.Frame(group_pane)
        group_list_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 4))
        
        group_scrollbar = tk.Scrollbar(group_list_frame, orient="vertical")
        self.group_listbox = tk.Listbox(group_list_frame, yscrollcommand=group_scrollbar.set, exportselection=False, font=content_font)

        # ★★★ placeを使って配置 ★★★
        group_scrollbar.place(relx=1.0, rely=0, relheight=1.0, anchor='ne')
        self.group_listbox.place(x=0, y=0, relwidth=1.0, relheight=1.0)
        # スクロールバーの分だけListboxの幅を狭める
        self.group_listbox.config(width=0) # これによりrelwidthが優先される
        self.group_listbox.place_configure(relwidth=1.0, bordermode='outside', width=-group_scrollbar.winfo_reqwidth())
        
        self.group_listbox.bind('<<ListboxSelect>>', self.on_group_select)

        self.no_results_label = tk.Label(group_list_frame, text="検索結果がありません", font=content_font, fg="gray")
        # ★★★ placeを使って重ねて配置 ★★★
        self.no_results_label.place(relx=0.5, rely=0.5, anchor='center')
        self.no_results_label.lower() # 最初は非表示（背面に）
        
        # --- グループ操作ボタン ---
        self.group_btns_frame = tk.Frame(group_pane)
        self.group_btns_frame.grid(row=2, column=0, sticky="ew")
        # (ボタンの作成とpackは変更なし)
        self.add_group_btn = tk.Button(self.group_btns_frame, text="追加", command=self.add_group)
        self.add_group_btn.pack(side="left")
        self.rename_group_btn = tk.Button(self.group_btns_frame, text="名変更", command=self.rename_group)
        self.rename_group_btn.pack(side="left")
        self.delete_group_btn = tk.Button(self.group_btns_frame, text="削除", command=self.delete_group)
        self.delete_group_btn.pack(side="left")
        self.move_group_up_btn = tk.Button(self.group_btns_frame, text="↑", command=self.move_group_up)
        self.move_group_up_btn.pack(side="left")
        self.move_group_down_btn = tk.Button(self.group_btns_frame, text="↓", command=self.move_group_down)
        self.move_group_down_btn.pack(side="left")

        main_pane.add(group_outer_frame)

        # =================================================================
        # === 右ペイン：リンク一覧 ===
        # =================================================================
        link_outer_frame = tk.Frame(main_pane)
        link_outer_frame.grid_rowconfigure(0, weight=1)
        link_outer_frame.grid_columnconfigure(0, weight=1)

        link_pane = tk.LabelFrame(link_outer_frame, text="リンク", font=title_font, bd=1, padx=5, pady=5)
        link_pane.grid(row=0, column=0, sticky="nsew")
        
        link_pane.grid_columnconfigure(0, weight=1)
        link_pane.grid_rowconfigure(1, weight=1) # Canvasの行を伸縮させる

        # --- アドレス入力ボックス ---
        link_addr_row = tk.Frame(link_pane)
        link_addr_row.grid(row=0, column=0, sticky="ew", pady=(2, 6))
        link_addr_row.grid_columnconfigure(0, weight=1)
        
        self.link_addr_var = tk.StringVar()
        self.link_addr_entry = ttk.Entry(link_addr_row, textvariable=self.link_addr_var, font=content_font)
        self.link_addr_entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        self.save_addr_btn = tk.Button(link_addr_row, text="保存", command=self.save_link_addr, height=1)
        self.save_addr_btn.grid(row=0, column=1, sticky="e")
        self.link_addr_entry.bind("<FocusIn>", self.on_link_addr_focus)
        
        # --- リンク一覧キャンバス ---
        #link_canvas_frame = tk.Frame(link_pane, bg="#ffffff", bd=1, relief=tk.SOLID)
        link_canvas_frame = tk.Frame(link_pane, bg="#ffffff")
        link_canvas_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 4))
        link_canvas_frame.grid_rowconfigure(0, weight=1)
        link_canvas_frame.grid_columnconfigure(0, weight=1)

        link_scrollbar = tk.Scrollbar(link_canvas_frame, orient="vertical")
        self.link_canvas = tk.Canvas(link_canvas_frame, bg="#ffffff", highlightthickness=0, yscrollcommand=link_scrollbar.set)
        self.link_canvas.grid(row=0, column=0, sticky="nsew")
        link_scrollbar.config(command=self.link_canvas.yview)
        link_scrollbar.grid(row=0, column=1, sticky="ns")
        self.link_canvas.bind("<Configure>", lambda e: self.refresh_link_list())
        self.link_canvas.bind("<MouseWheel>", self._on_link_canvas_mousewheel)
        self.link_canvas.bind("<Button-1>", self.on_link_canvas_click)
        self.link_canvas.bind("<Double-Button-1>", self.on_link_canvas_double)

        # --- リンク操作ボタン ---
        self.link_btns_frame = tk.Frame(link_pane)
        self.link_btns_frame.grid(row=2, column=0, sticky="ew")
        self.add_link_btn = tk.Button(self.link_btns_frame, text="追加", command=self.add_link)
        self.add_link_btn.pack(side="left")
        self.rename_link_btn = tk.Button(self.link_btns_frame, text="名変更", command=self.rename_link)
        self.rename_link_btn.pack(side="left")
        self.delete_link_btn = tk.Button(self.link_btns_frame, text="削除", command=self.delete_link)
        self.delete_link_btn.pack(side="left")
        self.move_link_up_btn = tk.Button(self.link_btns_frame, text="↑", command=self.move_link_up)
        self.move_link_up_btn.pack(side="left")
        self.move_link_down_btn = tk.Button(self.link_btns_frame, text="↓", command=self.move_link_down)
        self.move_link_down_btn.pack(side="left")

        main_pane.add(link_outer_frame)

        # --- OK/Cancelボタン ---
        button_frame = tk.Frame(self)
        #button_frame.grid(row=1, column=0, sticky="e", padx=10, pady=(5, 10))
        button_frame.grid(row=1, column=0, pady=(5, 10))
        tk.Button(button_frame, text="OK", width=10, command=self.ok).pack(side="left", padx=5)
        tk.Button(button_frame, text="キャンセル", width=10, command=self.cancel).pack(side="left")

        self.link_addr_entry.bind("<Return>", lambda e: self.save_link_addr())
        self.bind("<Escape>", self.cancel)
        self.bind("<Alt-F4>", self.cancel)

        # --- 初期化と表示処理 ---
        self.refresh_group_list()
        self.group_listbox.focus_set()
        self._update_buttons_state() # ボタンの初期状態を設定

        # # ★★★ PanedWindowの初期分割位置を設定 ★★★
        # self.update_idletasks() # ウィジェットのサイズを計算させる
        # try:
        #     # 全体の幅の約1/4を左ペインに割り当てる
        #     sash_position = self.winfo_width() // 4
        #     main_pane.sashpos(0, sash_position)
        # except tk.TclError:
        #     logging.warning("Failed to set initial sash position.")

        # ★最重要ポイント2: 親(root)に頼らず、スクリーンの中央に配置
        self.update_idletasks() # これでウィンドウの要求サイズが計算される
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        win_w = self.winfo_width()
        win_h = self.winfo_height()
        x = (screen_w // 2) - (win_w // 2)
        y = (screen_h // 2) - (win_h // 2)
        self.geometry(f'{win_w}x{win_h}+{x}+{y}')

        # ★最重要ポイント3: 明示的に表示し、モーダル化する
        self.deiconify() # ウィンドウを強制的に表示状態にする
        self.grab_set()  # 他のウィンドウを操作不可にする
        self.wait_window(self) # このウィンドウが閉じるまで待つ

    # # Canvas上のマウスイベントハンドラ
    # def on_canvas_motion(self, event):
    #     # "current"タグで、現在マウスカーソルがあるアイテムIDを取得
    #     # find_closestは最も近いアイテムを返すので、カーソル直下を取得できる
    #     item_id = self.link_canvas.find_closest(event.x, event.y)[0]
        
    #     if item_id in self.canvas_item_map:
    #         # マップに登録されたアイテムならツールチップを表示
    #         path = self.canvas_item_map[item_id]
    #         self.tooltip.showtip(path)
    #     else:
    #         # それ以外の場所ならツールチップを隠す
    #         self.tooltip.hidetip()

    # def on_canvas_leave(self, event):
    #     # Canvasからマウスが出たらツールチップを隠す
    #     self.tooltip.hidetip()

    # ok, cancel
    def ok(self, event=None):
        # 変更：常にマスターデータである original_groups を結果として返す
        self.result = self.original_groups
        self.destroy()

    def cancel(self, event=None):
        if self.modified:
            if not messagebox.askyesno("確認", "変更内容が保存されていません。破棄して閉じますか？", parent=self):
                return
        self.result = None
        self.destroy()

    def add_group(self):
        # このメソッドは検索中は呼ばれないが、念のためロジックを堅牢に
        name = simpledialog.askstring("グループ名", "新しいグループ名:", parent=self)
        if name:
            new_group = {'group': name, 'links': []}
            # 元データに追加
            self.original_groups.append(copy.deepcopy(new_group))
            # 表示データにも追加
            self.groups.append(copy.deepcopy(new_group))

            self.selected_group = len(self.groups) - 1
            self.refresh_group_list()
            self.refresh_link_list()
            self.modified = True

    def rename_group(self):
        if not self.groups or self.selected_group is None:
            return
        idx = self.selected_group
        
        # 変更前のグループ名を取得（これが元データを探すキーになる）
        original_name = self.groups[idx]['group']
        
        new_name = simpledialog.askstring("グループ名変更", "新しいグループ名:", initialvalue=original_name, parent=self)
        
        if new_name and new_name != original_name:
            # 表示用データを更新
            self.groups[idx]['group'] = new_name
            
            # 元データ(original_groups)も探し出して更新
            for g in self.original_groups:
                if g['group'] == original_name:
                    g['group'] = new_name
                    break # 見つけたら抜ける

            self.refresh_group_list()
            self.modified = True

    def delete_group(self):
        if not self.groups or self.selected_group is None:
            return
        
        group_to_delete_name = self.groups[self.selected_group]['group']

        if not messagebox.askyesno("確認", f"グループ '{group_to_delete_name}' を削除しますか？\n（中のリンクもすべて削除されます）", parent=self):
            return

        # 表示データから削除
        del self.groups[self.selected_group]
        
        # 元データ(original_groups)も探し出して削除
        self.original_groups = [g for g in self.original_groups if g['group'] != group_to_delete_name]

        self.selected_group = max(0, self.selected_group - 1)
        if not self.groups:
            self.selected_group = None

        self.refresh_group_list()
        self.refresh_link_list()
        self.modified = True

    def move_group_up(self):
        # 検索中は無効になっているはずだが、念のためチェック
        if self.is_searching: return
        
        idx = self.selected_group
        if idx > 0:
            # ★★★ original_groups を直接並べ替える ★★★
            self.original_groups[idx-1], self.original_groups[idx] = self.original_groups[idx], self.original_groups[idx-1]
            # 表示用データも同じように並べ替える
            self.groups = [dict(g, links=[dict(l) for l in g.get('links', [])]) for g in self.original_groups]
            
            self.selected_group -= 1
            self.refresh_group_list()
            self.modified = True

    def move_group_down(self):
        if self.is_searching: return
        
        idx = self.selected_group
        if idx < len(self.original_groups) - 1:
            # ★★★ original_groups を直接並べ替える ★★★
            self.original_groups[idx+1], self.original_groups[idx] = self.original_groups[idx], self.original_groups[idx+1]
            # 表示用データも同じように並べ替える
            self.groups = [dict(g, links=[dict(l) for l in g.get('links', [])]) for g in self.original_groups]

            self.selected_group += 1
            self.refresh_group_list()
            self.modified = True

    def add_link(self):
        if not self.groups or self.selected_group is None:
            return
        # クリップボードからデフォルト値取得
        clipboard_text = None
        try:
            clipboard_text = self.clipboard_get()
        except Exception:
            clipboard_text = None
        default_name = ""
        default_path = ""
        if isinstance(clipboard_text, str):
            text = clipboard_text.strip()
            if text.startswith("http://") or text.startswith("https://"):
                default_name = ""
                default_path = text
            else:
                base = os.path.basename(text)
                default_name = base
                default_path = text
        # それ以外（テキストでない場合）はデフォルト空
        name = simpledialog.askstring("リンク名", "新しいリンク名:", initialvalue=default_name, parent=self)
        if not name:  # Noneまたは空文字列
            return
        path = self.ask_dialog(self, "リンク先", "リンク先パスまたはURL:", initialvalue=default_path)
        if not path:  # Noneまたは空文字列
            return
        
        new_link_data = {'name': name, 'path': path}

        # --- ★★★ データ同期ロジック ★★★ ---
        # 1. 表示されているグループ名を取得
        current_group_name = self.groups[self.selected_group]['group']
        
        # 2. original_groups から該当グループを探し、そこに新しいリンクを追加
        for g in self.original_groups:
            if g['group'] == current_group_name:
                g['links'].append(new_link_data.copy())
                break
        # 3. 表示用のgroupsにも追加
        self.groups[self.selected_group]['links'].append(new_link_data.copy())
        self.selected_link = len(self.groups[self.selected_group]['links']) - 1

        # --- 追加したリンクのアイコンを個別にキャッシュ取得 ---
        try:
            if path.startswith('http'):
                get_web_icon(path, size=self.link_icon_size)
            else:
                get_file_icon(path, size=self.link_icon_size)
        except Exception as e:
            logging.info(f"[add_link] icon fetch failed: {path} : {e}")
        self.refresh_link_list()
        self._update_buttons_state()
        self.modified = True

    def rename_link(self):
        if not self.groups or self.selected_group is None or self.selected_link is None:
            return
        
        # --- ★★★ データ同期ロジック ★★★ ---
        # 1. 変更対象の情報を取得
        group_name = self.groups[self.selected_group]['group']
        link_idx = self.selected_link

        # ★重要: original_groups内の本当のオブジェクトを見つけるために、表示上のインデックスだけでなく、
        #        表示データと内容が一致するオブジェクトを探す必要があります。
        #        まず、表示されているリンクの情報を取得します。
        visible_link = self.groups[self.selected_group]['links'][link_idx]
        
        new_name = simpledialog.askstring("名前変更", "新しい名前:", initialvalue=visible_link['name'], parent=self)
        
        if new_name and new_name != visible_link['name']:
            # 2. original_groups から該当グループとリンクを探して更新
            for g in self.original_groups:
                if g['group'] == group_name:
                    # g['links']の中から、表示されているリンクと内容が一致するものを探す
                    for original_link in g['links']:
                        if original_link['name'] == visible_link['name'] and original_link['path'] == visible_link['path']:
                            original_link['name'] = new_name
                            break
                    break
            
            # 3. 表示用データも更新
            visible_link['name'] = new_name

            # 4. 画面を更新
            self.refresh_link_list()
            self.modified = True

    def delete_link(self):
        if not self.groups or self.selected_group is None or self.selected_link is None:
            return

        # --- ★★★ データ同期ロジック ★★★ ---
        # 1. 削除対象の情報を取得
        group_name = self.groups[self.selected_group]['group']
        link_to_delete = self.groups[self.selected_group]['links'][self.selected_link]
        
        # 2. original_groups から該当グループとリンクを探して削除
        for g in self.original_groups:
            if g['group'] == group_name:
                g['links'] = [link for link in g['links'] if not (link['name'] == link_to_delete['name'] and link['path'] == link_to_delete['path'])]
                break

        # 3. 表示用データから削除
        del self.groups[self.selected_group]['links'][self.selected_link]
        self.selected_link = None
        
        # 4. 画面を更新
        self.refresh_link_list()
        self._update_buttons_state()
        self.modified = True

    def move_link_up(self):
        if self.is_searching: return
        if not self.groups or self.selected_group is None or self.selected_link is None or self.selected_link == 0:
            return

        # ★★★ original_groups 内の該当リンクを直接並べ替える ★★★
        group_name = self.groups[self.selected_group]['group']
        link_idx = self.selected_link
        
        for g in self.original_groups:
            if g['group'] == group_name:
                g['links'][link_idx-1], g['links'][link_idx] = g['links'][link_idx], g['links'][link_idx-1]
                break
        
        # 表示用データも更新
        links = self.groups[self.selected_group]['links']
        links[link_idx-1], links[link_idx] = links[link_idx], links[link_idx-1]
        
        self.selected_link -= 1
        self.refresh_link_list()
        self.modified = True

    def move_link_down(self):
        if self.is_searching: return
        if not self.groups or self.selected_group is None or self.selected_link is None:
            return

        links = self.groups[self.selected_group]['links']
        link_idx = self.selected_link
        if link_idx >= len(links) - 1:
            return

        # ★★★ original_groups 内の該当リンクを直接並べ替える ★★★
        group_name = self.groups[self.selected_group]['group']
        
        for g in self.original_groups:
            if g['group'] == group_name:
                g['links'][link_idx+1], g['links'][link_idx] = g['links'][link_idx], g['links'][link_idx+1]
                break

        # 表示用データも更新
        links[link_idx+1], links[link_idx] = links[link_idx], links[link_idx]
        
        self.selected_link += 1
        self.refresh_link_list()
        self.modified = True

    def refresh_link_list(self):
        self.link_canvas.delete("all")
        self.icon_refs.clear()
        # self.canvas_item_map.clear()

        font_name_tuple = (self.settings['font'], self.settings['size'])
        font_path_tuple = (self.settings['font'], self.settings['size'])
        color_path = "#888888"

        name_font_obj = tkfont.Font(font=font_name_tuple)
        font_metrics = name_font_obj.metrics()
        font_height = font_metrics.get('linespace', font_metrics.get('height', 16))
        padding = 8
        self.link_row_height = font_height + padding

        self.link_icon_size = font_metrics.get('ascent', 16)
        self.link_icon_size = round_to_step(self.link_icon_size, step=4)
        self.link_icon_size = max(12, min(self.link_icon_size, 32))

        if not self.groups or self.selected_group is None or self.selected_group >= len(self.groups):
            self.link_addr_entry.config(state="disabled")
            self.save_addr_btn.config(state="disabled")
            self.link_addr_var.set("")
            return
        links = self.groups[self.selected_group]['links']
        y = 2
        canvas_width = self.link_canvas.winfo_width() or 360
        for i, link in enumerate(links):
            path = link['path']
            icon = None
            if path.startswith('http'):
                domain = urlparse(path).netloc
                key = (domain, self.link_icon_size)
                icon = _icon_cache.get(key)
            else:
                if self.link_icon_size > 20:
                    flags = 0x100 | 0x0
                else:
                    flags = 0x100 | 0x1
                key = (path, self.link_icon_size, flags)
                icon = _icon_cache.get(key)
            if not icon:
                icon = _create_fallback_icon(self.link_icon_size)
            if icon:
                self.link_canvas.create_image(4, y + self.link_row_height // 2, image=icon, anchor="w")
                self.icon_refs.append(icon)
            # アイコンの右側余白を1px、左側を4pxに
            name_x_start = 4 + self.link_icon_size + 1
            # 省略せずフルテキストで表示
            display_name = link['name']
            if self.selected_link == i:
                self.link_canvas.create_rectangle(0, y, canvas_width, y+self.link_row_height, outline="#3399ff", width=2)
            name_id = self.link_canvas.create_text(name_x_start, y + self.link_row_height // 2,
                                                   text=display_name, anchor="w", font=font_name_tuple)
            name_width = name_font_obj.measure(display_name)
            path_x_start = name_x_start + name_width + 25
            path_id = self.link_canvas.create_text(path_x_start, y + self.link_row_height // 2,
                                                   text=link['path'], anchor="w", font=font_path_tuple, fill=color_path)
            y += self.link_row_height
        self.link_canvas.config(scrollregion=(0,0,canvas_width,y))
        if self.selected_link is not None and 0 <= self.selected_link < len(links):
            self.link_addr_entry.config(state="normal")
            self.save_addr_btn.config(state="normal")
            self.link_addr_var.set(links[self.selected_link]['path'])
        else:
            self.link_addr_entry.config(state="disabled")
            self.save_addr_btn.config(state="disabled")
            self.link_addr_var.set("")

    def refresh_group_list(self):
        self.group_listbox.delete(0, tk.END)
        for g in self.groups:
            self.group_listbox.insert(tk.END, " " + g['group'])  # 先頭にスペースで左余白
        if self.groups:
            self.group_listbox.select_set(self.selected_group)
        # グループリスト更新時はリンクリストを空に
        self.refresh_link_list()

    def on_group_select(self, event):
        idx = self.group_listbox.curselection()
        if idx:
            self.selected_group = idx[0]
            self.selected_link = None
            self.refresh_link_list()
            self.link_canvas.config(state="normal")

    def on_link_canvas_click(self, event):
        # 現在表示されているグループのリンク一覧を取得
        if not self.groups or self.selected_group is None:
            return
        links = self.groups[self.selected_group]['links']

        # クリックされたY座標から、何番目のリンクかを計算
        idx = (event.y - 2) // self.link_row_height
        
        if 0 <= idx < len(links):
            # 有効なリンクがクリックされた場合
            if self.selected_link == idx:
                # すでに選択されている項目を再度クリックした場合は選択を解除
                self.selected_link = None
            else:
                self.selected_link = idx
            self.link_addr_var.set(links[self.selected_link]['path'] if self.selected_link is not None else "")
        else:
            # リンク以外の場所（空白領域）がクリックされた場合は選択を解除
            self.selected_link = None
            self.link_addr_var.set("")
            
        # 選択状態が変わったので、Canvasを再描画してハイライトを更新
        self.refresh_link_list()
        self._update_buttons_state()

    def on_link_canvas_double(self, event):
        idx = (event.y - 2) // self.link_row_height
        links = self.groups[self.selected_group]['links']
        if 0 <= idx < len(links):
            path = links[idx]['path']
            open_link(path)

    def save_link_addr(self):
        if not self.groups or self.selected_group is None or self.selected_link is None:
            return

        # --- ★★★ データ同期ロジック ★★★ ---
        # 1. 変更対象の情報を取得
        group_name = self.groups[self.selected_group]['group']
        visible_link = self.groups[self.selected_group]['links'][self.selected_link]
        old_path = visible_link['path']
        new_path = self.link_addr_var.get().strip()

        if not new_path or new_path == old_path:
            return
            
        # 2. original_groups から該当グループとリンクを探して更新
        for g in self.original_groups:
            if g['group'] == group_name:
                # 名前と古いパスでユニークに特定
                for original_link in g['links']:
                    if original_link['name'] == visible_link['name'] and original_link['path'] == old_path:
                        original_link['path'] = new_path
                        break
                break

        # 3. 表示用データも更新
        visible_link['path'] = new_path

        # キャッシュクリア（旧パス・新パス両方）
        for p in (old_path, new_path):
            key = (p, self.link_icon_size)
            if key in _icon_cache:
                del _icon_cache[key]
        # 新しいパスのアイコンを取得
        try:
            if new_path.startswith('http'):
                get_web_icon(new_path, size=self.link_icon_size)
            else:
                get_file_icon(new_path, size=self.link_icon_size)
        except Exception as e:
            logging.info(f"[save_link_addr] icon fetch failed: {new_path} ({self.link_icon_size}px): {e}")

        # 4. 画面を更新
        self.refresh_link_list()
        self.modified = True

    def on_link_addr_focus(self, event):
        self.link_addr_entry.icursor(tk.END)

    def _on_link_canvas_mousewheel(self, event):
        scroll_info = self.link_canvas.yview()
        if scroll_info[0] == 0.0 and scroll_info[1] == 1.0:
            return
    
        if event.num == 4:
            delta = -1
        elif event.num == 5:
            delta = 1
        else:
            delta = -1 * (event.delta // 120)
            
        self.link_canvas.yview_scroll(delta, "units")

    def _on_search_change(self, *args):
        query = self.search_var.get().strip()
        if query:
            self.is_searching = True
            self._perform_search(query)
        else:
            self.is_searching = False
            # 元のリストに戻す
            self.groups = [dict(g, links=[dict(l) for l in g.get('links', [])]) for g in self.original_groups]
            self.selected_group = 0 if self.groups else None
            
            self.no_results_label.lower() # ラベルを背面に
            
        self.refresh_group_list()
        self._update_buttons_state()

    def _perform_search(self, query):
        """実際に検索処理を実行し、表示用データを生成する"""
        query_lower = query.lower()
        search_results = []
        
        for group_data in self.original_groups:
            matched_links = []
            group_name_match = query_lower in group_data['group'].lower()
            
            for link_data in group_data.get('links', []):
                if group_name_match:
                    matched_links.append(link_data.copy())
                    continue
                
                if query_lower in link_data['name'].lower() or query_lower in link_data['path'].lower():
                    matched_links.append(link_data.copy())
            
            if matched_links:
                search_results.append({'group': group_data['group'], 'links': matched_links})

        self.groups = search_results
        self.selected_group = 0 if self.groups else None

        if not self.groups:
            self.no_results_label.lift() # ラベルを前面に
        else:
            self.no_results_label.lower() # ラベルを背面に

    def _update_buttons_state(self):
        """現在の状態に応じて、すべてのボタンの有効/無効を切り替える"""
        
        # --- 基本的な状態判定 ---
        # 検索中か？
        is_searching = self.is_searching
        # グループリストに表示項目があるか？
        has_groups = bool(self.groups)
        # 何かグループが選択されているか？
        group_selected = self.selected_group is not None and has_groups
        # 何かリンクが選択されているか？
        link_selected = self.selected_link is not None and group_selected
        
        # --- 状態変数 ---
        # 検索中は無効、それ以外は有効
        search_dependent_state = "disabled" if is_searching else "normal"
        # グループが選択されていれば有効
        group_dependent_state = "normal" if group_selected else "disabled"
        # リンクが選択されていれば有効
        link_dependent_state = "normal" if link_selected else "disabled"

        # --- グループ操作ボタンの状態を更新 ---
        self.add_group_btn.config(state=search_dependent_state)
        self.move_group_up_btn.config(state=search_dependent_state)
        self.move_group_down_btn.config(state=search_dependent_state)
        
        # 名変更と削除は、グループが選択されていれば検索中でも有効
        self.rename_group_btn.config(state=group_dependent_state)
        self.delete_group_btn.config(state=group_dependent_state)

        # --- リンク操作ボタンの状態を更新 ---
        # リンクの移動は検索中は無効
        self.move_link_up_btn.config(state=search_dependent_state)
        self.move_link_down_btn.config(state=search_dependent_state)

        # リンクの追加は、グループが選択されていれば検索中でも有効
        self.add_link_btn.config(state=group_dependent_state)
        
        # リンクの名変更と削除は、リンクが選択されていれば検索中でも有効
        self.rename_link_btn.config(state=link_dependent_state)
        self.delete_link_btn.config(state=link_dependent_state)

    # 入力ダイアログをカスタムしてEntry幅を指定
    @staticmethod
    def ask_dialog(parent, title, prompt, initialvalue=""):
        dialog = tk.Toplevel(parent)
        dialog.title(title)
        if 'app_icon' in globals() and app_icon:
            dialog.iconphoto(True, app_icon)
        dialog.transient(parent)

        # promptを左寄せ
        prompt_label = tk.Label(dialog, text=prompt, anchor="w", justify="left")
        prompt_label.pack(fill="x", padx=10, pady=(10, 2))
        var = tk.StringVar(value=initialvalue)
        entry = tk.Entry(dialog, textvariable=var, width=75)
        entry.pack(padx=10, pady=(0, 10))
        entry.focus_set()
        result = []
        def on_ok(event=None):
            result.append(var.get())
            dialog.destroy()
        def on_cancel(event=None):
            dialog.destroy()
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=(0, 10))
        ok_btn = tk.Button(btn_frame, text="OK", width=10, command=on_ok)
        ok_btn.pack(side="left", padx=5)
        cancel_btn = tk.Button(btn_frame, text="キャンセル", width=10, command=on_cancel)
        cancel_btn.pack(side="left", padx=5)
        dialog.bind("<Return>", on_ok)
        dialog.bind("<Escape>", on_cancel)
        # 親画面中央に正確に表示
        dialog.update_idletasks()
        try:
            parent_x = parent.winfo_rootx()
            parent_y = parent.winfo_rooty()
            parent_w = parent.winfo_width()
            parent_h = parent.winfo_height()
            if parent_w < 10 or parent_h < 10:
                raise Exception()
        except Exception:
            screen_w = dialog.winfo_screenwidth()
            screen_h = dialog.winfo_screenheight()
            win_w = dialog.winfo_width()
            win_h = dialog.winfo_height()
            x = (screen_w - win_w) // 2
            y = (screen_h - win_h) // 2
        else:
            win_w = dialog.winfo_width()
            win_h = dialog.winfo_height()
            x = parent_x + (parent_w - win_w) // 2
            y = parent_y + (parent_h - win_h) // 2
        dialog.geometry(f"+{x}+{y}")
        dialog.deiconify()
        dialog.grab_set()  # ここでgrab_set
        parent.wait_window(dialog)
        return result[0] if result else None

class LinkPopup(tk.Toplevel):
    # --- レイアウト定数 ---
    # ICON_COLUMN_WIDTH = 24  # ←固定値を廃止
    TEXT_LEFT_PADDING = 0   # アイコンとテキストの間の隙間
    
    def __init__(self, master, settings, profile_name):
        super().__init__(master)
        if app_icon:
            self.iconphoto(True, app_icon)
        self.profile_name = profile_name
        self.settings = settings.copy()
        self.group_row_height = 24 # デフォルト値
        self.icon_size = 16 # デフォルト値
        self.overrideredirect(True)
        self.withdraw()
        self.bind("<FocusOut>", lambda e: self.withdraw())
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.canvas.pack(expand=True, fill="both", padx=1, pady=1)
        self.canvas.bind("<Motion>", self.on_motion)
        self.canvas.bind("<Leave>", self.on_leave)
        
        self.group_map = []
        self.link_items = {}
        self.hover_group = None
        self.link_popup = None
        self._leave_after_id = None
        #self.folder_icon = get_system_folder_icon(size=16)
        self.folder_icon = None
        self.arrow_icon = None

        self.reload_links()
        self.apply_settings(self.settings)
        LinkPopup.current_icon_size = self.icon_size

    def apply_settings(self, settings):
        self.settings = settings.copy()
        border_color = self.settings.get('border_color', DEFAULT_SETTINGS['border_color'])
        content_bg_color = self.settings.get('bg', DEFAULT_SETTINGS['bg'])
        self.config(bg=border_color)
        self.canvas.config(bg=content_bg_color)

        font_main = tkfont.Font(family=self.settings['font'], size=self.settings['size'])
        font_metrics = font_main.metrics()
        font_height = font_metrics.get('linespace', font_metrics.get('height', 16))
        padding = 4  # 上下の余白（合計）
        self.group_row_height = font_height + padding

        self.icon_size = font_metrics.get('ascent', 16)
        self.icon_size = round_to_step(self.icon_size, step=4)
        self.icon_size = max(12, min(self.icon_size, 32))
        self.ICON_COLUMN_WIDTH = self.icon_size + 5  # 左4px+アイコン+右1px
        # フォルダアイコンを必ず取得
        self.folder_icon = get_system_folder_icon(size=self.icon_size)
        self.draw_list()

    def reload_profile(self, profile_name):
        self.profile_name = profile_name
        # settingsは共有なので再読み込みは不要。ただし、必要ならここで読み込んでも良い。
        # self.settings = load_settings() 
        self.reload_links() # 新しいプロファイルのlinks.jsonを読み込む
        self.apply_settings(self.settings) # 見た目を更新

    def reload_links(self):
        self.link_items.clear()
        self.group_map.clear()
        groups_data = load_links_data(self.profile_name)
        if groups_data is None:
            links_os = self._load_os_links()
            groups_data = [{"group": "マイリンク", "links": [{"name": n, "path": p} for n, p in links_os.items()]}]
            save_links_data(groups_data)
        for group in groups_data:
            self.link_items[group['group']] = group['links']
            self.group_map.append(group['group'])

    def _load_os_links(self):
        pythoncom.CoInitialize()
        links = {}
        links_folder = os.path.join(os.environ['USERPROFILE'], 'Links')
        if not os.path.isdir(links_folder): return {}
        for filename in os.listdir(links_folder):
            filepath = os.path.join(links_folder, filename)
            name, ext = os.path.splitext(filename)
            target = None
            try:
                if ext == '.url':
                    config = configparser.ConfigParser(); config.read(filepath, encoding='utf-8-sig')
                    target = config.get('InternetShortcut', 'URL')
                elif ext == '.lnk':
                    shell = Dispatch("WScript.Shell"); shortcut = shell.CreateShortCut(filepath)
                    target = shortcut.TargetPath
            except Exception as e: logging.warning(f"OS link parse error: {e}")
            if target: links[name] = target
        return dict(sorted(links.items()))

    def show(self):
        self.reload_links()
        self.draw_list()
        self.update_idletasks() # ウィンドウサイズを計算させる

        # --- ここからが座標調整ロジック ---
        # タスクバーを除いた作業領域の座標を取得
        wa_left, wa_top, wa_right, wa_bottom = get_work_area()

        win_w = self.winfo_width()
        win_h = self.winfo_height()

        # 理想の表示座標を計算（カーソルの少し上に表示）
        x = self.winfo_pointerx() - win_w // 2
        y = self.winfo_pointery() - win_h - 10 

        # X座標の調整 (作業領域の左右にはみ出さないように)
        if x + win_w > wa_right:
            x = wa_right - win_w - 5 # 右端に余裕を持たせる
        if x < wa_left:
            x = wa_left + 5 # 左端に余裕を持たせる

        # Y座標の調整 (作業領域の上下にはみ出さないように)
        # まず、上にはみ出す場合
        if y < wa_top:
            # カーソルの下側に表示位置を変更
            y = self.winfo_pointery() + 20
        
        # 次に、下にはみ出す場合（タスクバーを考慮）
        if y + win_h > wa_bottom:
            # 作業領域の下端ぴったりに合わせる
            y = wa_bottom - win_h - 5 # 下端に余裕を持たせる
        # --- ここまでが座標調整ロジック ---

        self.geometry(f"+{x}+{y}")
        self.deiconify()
        self.lift()
        self.focus_force()

    def draw_list(self):
        self.canvas.delete("all")
        self.canvas.image_refs = [] 
        RIGHT_PADDING = 15 
        ARROW_AREA_WIDTH = 20
        font_main = (self.settings['font'], self.settings['size'])
        font_metrics = tkfont.Font(font=font_main)
        max_group_width = 220  # グループ名の最大幅(px)
        # --- ここでarrow_char/font_arrowを定義 ---
        try:
            arrow_font_size = max(12, self.icon_size - 2)
            font_arrow = tkfont.Font(family='Marlett', size=arrow_font_size)
            arrow_char = '4'
        except tk.TclError:
            font_arrow = font_metrics
            arrow_char = '>'
        maxlen = max((min(font_metrics.measure(g), max_group_width) for g in self.group_map), default=0)
        canvas_w = min(max(self.ICON_COLUMN_WIDTH + self.TEXT_LEFT_PADDING + maxlen + RIGHT_PADDING, 120), 400)
        canvas_h = len(self.group_map) * self.group_row_height
        self.canvas.config(width=canvas_w, height=canvas_h)
        main_font_color = self.settings['font_color']
        arrow_color = lighten_color(main_font_color, amount=0.6)
        y = 0
        for i, group in enumerate(self.group_map):
            bg_color = "#eaf6ff" if self.hover_group == i else self.settings['bg']
            self.canvas.create_rectangle(0, y, canvas_w, y + self.group_row_height, fill=bg_color, outline="")
            if self.folder_icon:
                icon_x = 4 + self.icon_size // 2
                self.canvas.create_image(icon_x, y + self.group_row_height  // 2, image=self.folder_icon, anchor="center")
                self.canvas.image_refs.append(self.folder_icon)
            # グループ名を省略表示
            display_group = ellipsize_text(group, font_metrics, max_group_width)
            self.canvas.create_text(
                self.ICON_COLUMN_WIDTH + self.TEXT_LEFT_PADDING + 1, 
                y + self.group_row_height  // 2, 
                text=display_group, 
                anchor="w", 
                font=font_main, 
                fill=main_font_color
            )
            if self.link_items.get(group):
                arrow_x = canvas_w - (ARROW_AREA_WIDTH // 2)
                self.canvas.create_text(
                    arrow_x, 
                    y + self.group_row_height  // 2, 
                    text=arrow_char, 
                    anchor="center", 
                    font=font_arrow, 
                    fill=arrow_color
                )
            y += self.group_row_height

    def on_motion(self, event):
        row_h = self.group_row_height 
        hover = event.y // row_h if 0 <= event.y < len(self.group_map) * row_h else None
        if hover != self.hover_group:
            self.hover_group = hover
            self.draw_list()
            self.show_link_popup(hover)

    def on_leave(self, event):
        if hasattr(self, '_leave_after_id') and self._leave_after_id: self.after_cancel(self._leave_after_id)
        self._leave_after_id = self.after(250, self._delayed_hide)

    def _on_link_popup_leave(self):
        # on_leaveとほぼ同じ処理。タイマーを開始する
        if hasattr(self, '_leave_after_id') and self._leave_after_id:
            self.after_cancel(self._leave_after_id)
        self._leave_after_id = self.after(250, self._delayed_hide)

    def _delayed_hide(self):
        if not self._point_in_window(self.winfo_pointerx(), self.winfo_pointery(), self) and \
           not (self.link_popup and self._point_in_window(self.winfo_pointerx(), self.winfo_pointery(), self.link_popup)):
            self.hover_group = None
            if self.link_popup: self.link_popup.destroy(); self.link_popup = None
            self.draw_list()
        self._leave_after_id = None
        
    def show_link_popup(self, group_idx):
        if self.link_popup: self.link_popup.destroy(); self.link_popup = None
        if group_idx is None or not (0 <= group_idx < len(self.group_map)): return
        links = self.link_items.get(self.group_map[group_idx], [])
        if not links: return

        popup = tk.Toplevel(self)
        popup.overrideredirect(True); popup.attributes("-topmost", True)

        border_color = self.settings.get('border_color', DEFAULT_SETTINGS['border_color'])
        content_bg_color = self.settings.get('bg', DEFAULT_SETTINGS['bg'])

        popup.config(bg=border_color)
        frame = tk.Frame(popup, bg=content_bg_color)
        frame.pack(expand=True, fill="both", padx=1, pady=1)

        font_link = (self.settings['font'], self.settings['size'])
        font_underline = tkfont.Font(font=font_link); font_underline.config(underline=True)
        font_link_obj = tkfont.Font(font=font_link)
        max_link_width = 320  # リンク名の最大幅(px)
        popup.icon_refs = []

        for i, link in enumerate(links):
            row = tk.Frame(frame, bg=self.settings['bg'])
            row.grid(row=i, column=0, sticky="ew")
            frame.grid_columnconfigure(0, weight=1)
            row.grid_columnconfigure(0, minsize=self.ICON_COLUMN_WIDTH)
            row.grid_columnconfigure(1, weight=1)
            path = link['path']
            if path.startswith('http'):
                domain = urlparse(path).netloc
                key = (domain, self.icon_size)
                icon = _icon_cache.get(key)
                if not icon:
                    icon = _create_fallback_icon(self.icon_size)
            else:
                if self.icon_size > 20:
                    flags = 0x100 | 0x0
                else:
                    flags = 0x100 | 0x1
                key = (path, self.icon_size, flags)
                icon = _icon_cache.get(key)
                if not icon:
                    icon = _create_fallback_icon(self.icon_size)
            icon_label = tk.Label(row, image=icon, bg=self.settings['bg'])
            icon_label.grid(row=0, column=0, sticky="nse", padx=(0, self.TEXT_LEFT_PADDING))
            popup.icon_refs.append(icon)
            # リンク名を省略表示
            display_name = ellipsize_text(link['name'], font_link_obj, max_link_width)
            text_label = tk.Label(row, text=display_name, anchor="w", bg=self.settings['bg'], font=font_link, fg=self.settings['font_color'])
            text_label.grid(row=0, column=1, sticky="w")
            def on_enter(e, lbl=text_label): lbl.config(font=font_underline)
            def on_leave(e, lbl=text_label): lbl.config(font=font_link)
            def on_click(e, p=path): self.open_and_close(p)
            for w in [row, icon_label, text_label]:
                w.bind("<Enter>", on_enter); w.bind("<Leave>", on_leave); w.bind("<Button-1>", on_click)
        self.update_idletasks()
        popup.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        parent_x = self.winfo_rootx()
        parent_y = self.winfo_rooty()
        parent_w = self.winfo_width()
        popup_w = popup.winfo_width()
        popup_h = popup.winfo_height()
        x = parent_x + parent_w
        y = parent_y + group_idx * self.group_row_height 
        if x + popup_w > screen_w:
            x = parent_x - popup_w
        if y + popup_h > screen_h:
            y = screen_h - popup_h - 5
        if y < 0:
            y = 5
        popup.geometry(f"+{x}+{y}")
        self.link_popup = popup
        popup.bind("<Leave>", lambda e: self._on_link_popup_leave())

    def open_and_close(self, path):
        open_link(path)
        if self.link_popup: self.link_popup.destroy(); self.link_popup = None
        self.withdraw()

    def _point_in_window(self, x, y, win):
        try:
            return (win.winfo_rootx() <= x < win.winfo_rootx() + win.winfo_width() and
                    win.winfo_rooty() <= y < win.winfo_rooty() + win.winfo_height())
        except tk.TclError: return False

class ProfileManagerDialog(simpledialog.Dialog):
    def __init__(self, parent, current_profile):
        self.current_profile = current_profile
        self.profiles = get_all_profile_names()
        self.result = None
        super().__init__(parent, "プロファイルの管理")

    def body(self, master):
        if 'app_icon' in globals() and app_icon:
            self.iconphoto(True, app_icon)
        
        self.resizable(False, False)
        
        list_frame = tk.Frame(master, bd=1, relief=tk.SOLID)
        list_frame.pack(padx=10, pady=10, expand=True, fill="both")
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        
        self.listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, exportselection=False, height=10)
        self.listbox.pack(side="left", expand=True, fill="both")
        scrollbar.config(command=self.listbox.yview)

        for p_name in self.profiles:
            display_name = p_name
            if p_name == self.current_profile:
                display_name += " (現在)"
            self.listbox.insert(tk.END, display_name)
            if p_name == self.current_profile:
                self.listbox.itemconfig(tk.END, {'bg': '#e0e0e0'})
        
        # 初期選択
        try:
            current_idx = self.profiles.index(self.current_profile)
            self.listbox.select_set(current_idx)
        except ValueError:
            pass

        btn_frame = tk.Frame(master)
        btn_frame.pack(padx=10, pady=(0, 10), fill="x")
        
        tk.Button(btn_frame, text="選択して切り替え", command=self.switch_profile).pack(side="left", padx=2)
        tk.Button(btn_frame, text="追加", command=self.add_profile).pack(side="left", padx=2)
        tk.Button(btn_frame, text="名前変更", command=self.rename_profile).pack(side="left", padx=2)
        tk.Button(btn_frame, text="削除", command=self.delete_profile).pack(side="left", padx=2)
        
        return self.listbox

    def get_selected_profile(self):
        selected_indices = self.listbox.curselection()
        if not selected_indices:
            return None
        return self.profiles[selected_indices[0]]

    def switch_profile(self):
        selected = self.get_selected_profile()
        if selected:
            self.result = {"action": "switch", "profile": selected}
            self.ok()

    def add_profile(self):
        new_name = simpledialog.askstring("新規プロファイル", "新しいプロファイル名:", parent=self)
        if new_name and new_name not in self.profiles:
            self.result = {"action": "add", "profile": new_name}
            self.ok()
        elif new_name:
            messagebox.showerror("エラー", "その名前は既に使用されています。", parent=self)

    def rename_profile(self):
        # 1. プロファイルが選択されているかチェック
        selected = self.get_selected_profile()
        if not selected:
            messagebox.showwarning("警告", "名前を変更するプロファイルを選択してください。", parent=self)
            return
            
        # 2. デフォルトプロファイルは変更不可
        if selected == DEFAULT_PROFILE_NAME:
            messagebox.showerror("エラー", "デフォルトプロファイルの名前は変更できません。", parent=self)
            return
            
        # 3. 新しい名前をユーザーから入力してもらう
        new_name = simpledialog.askstring("名前の変更", f"'{selected}' の新しい名前:", initialvalue=selected, parent=self)
        
        # 4. 入力の検証
        if not new_name or new_name == selected:
            # キャンセルされたか、名前が変わっていない場合は何もしない
            return
            
        if new_name in self.profiles:
            # 新しい名前が既に存在する場合
            messagebox.showerror("エラー", f"プロファイル名 '{new_name}' は既に使用されています。", parent=self)
            return

        # 5. 結果を辞書に格納してダイアログを閉じる
        self.result = {"action": "rename", "old": selected, "new": new_name}
        self.ok()

    def delete_profile(self):
        selected = self.get_selected_profile()
        if not selected:
            messagebox.showwarning("警告", "プロファイルを選択してください。", parent=self)
            return
        if selected == DEFAULT_PROFILE_NAME:
            messagebox.showerror("エラー", "デフォルトプロファイルは削除できません。", parent=self)
            return
            
        if messagebox.askyesno("確認", f"プロファイル '{selected}' を削除しますか？\nこの操作は元に戻せません。", parent=self):
            self.result = {"action": "delete", "profile": selected}
            self.ok()
            
    def buttonbox(self):
        # OK/Cancelボタンは不要なので、何もしないようにオーバーライド
        pass

def ellipsize_text(text, font_obj, max_width):
    if font_obj.measure(text) <= max_width:
        return text
    for i in range(len(text), 0, -1):
        s = text[:i] + '...'
        if font_obj.measure(s) <= max_width:
            return s
    return '...'

# --- アプリケーション実行部 ---
def main():
    try:
        # Windows 8.1以降で推奨されるPer-Monitor V2 DPI Awarenessを試す
        # PROCESS_PER_MONITOR_DPI_AWARE = 2
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        logging.info("SetProcessDpiAwareness(2) succeeded.")
    except (AttributeError, OSError):
        # shcore.dllがない古いWindows (7, 8) 向けにフォールバック
        try:
            # System-wide DPI Awarenessを有効にする
            user32.SetProcessDPIAware()
            logging.info("SetProcessDPIAware() succeeded.")
        except (AttributeError, OSError):
            logging.warning("Failed to set DPI awareness.")

    global root, popup, settings, app_icon

    # --- ★★★ アプリケーション起動時の処理 ★★★ ---
    # 1. 設定ファイルを読み込む
    settings = load_settings()

    # 2. デフォルトプロファイルが存在するか確認し、なければ作成
    if not os.path.exists(get_profile_path(DEFAULT_PROFILE_NAME)):
        # デフォルトの空のlinks.jsonを作成
        save_links_data([{"group": "マイリンク", "links": []}], DEFAULT_PROFILE_NAME)
        
    current_profile_name = settings.get('current_profile', DEFAULT_PROFILE_NAME)

    def run_tray():
        # --- 1. アイコン画像の準備 (フォールバックを含む) ---
        image = None
        try:
            # まずは外部のicon.pngファイルを試す
            icon_path = os.path.join(BASE_DIR, "icon.png")
            if os.path.exists(icon_path):
                image = Image.open(icon_path)
            else:
                # ファイルがなければBase64から復元
                import base64
                image_data = base64.b64decode(ICON_BASE64)
                image = Image.open(io.BytesIO(image_data))
        except Exception as e:
            logging.error(f"Failed to load icon image, falling back to default. Error: {e}")
        
        # 最終的なフォールバック: すべての試みが失敗した場合
        if image is None:
            image = Image.new('RGB', (64, 64), 'blue')
            draw = ImageDraw.Draw(image)
            # フォントを読み込んで、より綺麗に描画する（任意だが推奨）
            try:
                font = ImageFont.truetype("yugothic.ttf", 32)
            except IOError:
                font = ImageFont.load_default()
            draw.text((15, 12), "QL", font=font, fill="white")

        # --- 2. メニュー項目のアクション定義 ---
        def show_popup_action(icon=None): root.after(0, popup.show)
        def edit_links_action(icon=None): root.after(0, open_links_editor)
        def profile_action(icon=None): root.after(0, open_profile_manager)
        def settings_action(icon=None): root.after(0, open_settings_dialog)
        def exit_action(icon, item):
            icon.stop()
            root.destroy()

        # --- 3. メニューとアイコンの作成・実行 ---
        try:
            menu = Menu(item('リンクを表示', show_popup_action, default=True), 
                        item('リンク編集', edit_links_action),
                        item('プロファイル管理', profile_action),
                        item('設定', settings_action), 
                        Menu.SEPARATOR, 
                        item('終了', exit_action))
        
            icon = Icon("QuickLauncher", image, "QuickLauncher", menu)
            icon.run()
        except Exception as e:
            # Iconの作成やrun()自体が失敗するような致命的なエラー
            logging.error(f"Failed to run tray icon: {e}", exc_info=True)
            if not root.winfo_exists(): return
            root.destroy()

    def open_links_editor():
        links_data = load_links_data(current_profile_name)
        dialog = LinksEditDialog(root, links_data, settings)
        if dialog.result is not None:
            save_links_data(dialog.result, current_profile_name)
            if popup:
                popup.reload_profile(current_profile_name)

    def open_settings_dialog():
        dialog = SettingsDialog(root, settings)
        if dialog.result is not None:
            settings.clear()
            settings.update(dialog.result)
            if popup: popup.apply_settings(settings)

    def open_profile_manager():
        nonlocal current_profile_name # 外側の変数を変更するためにnonlocalを宣言
        
        dialog = ProfileManagerDialog(root, current_profile_name)
        result = dialog.result
        
        if result:
            action = result['action']
            
            if action == 'switch':
                current_profile_name = result['profile']
            elif action == 'add':
                new_profile = result['profile']
                save_links_data([{"group": "マイリンク", "links": []}], new_profile)
                current_profile_name = new_profile
            elif action == 'rename':
                old_profile_name = result['old']
                new_profile_name = result['new']
                old_path = get_profile_path(old_profile_name)
                new_path = os.path.join(PROFILES_DIR, new_profile_name)
                # 念のため、移動先が本当に存在しないか確認
                if os.path.exists(new_path):
                    messagebox.showerror("エラー", f"予期せぬエラー: '{new_profile_name}' は既に存在します。", parent=root)
                    return
                try:
                    # os.rename を使ってディレクトリの名前を直接変更
                    os.rename(old_path, new_path)
                    # 現在選択中のプロファイルがリネーム対象だった場合、
                    # current_profile_name も新しい名前に更新する
                    if current_profile_name == old_profile_name:
                        current_profile_name = new_profile_name
                    messagebox.showinfo("成功", f"プロファイル名を '{old_profile_name}' から '{new_profile_name}' に変更しました。", parent=root)
                except Exception as e:
                    logging.error(f"Failed to rename profile directory: {e}")
                    messagebox.showerror("エラー", f"プロファイルの名前変更中にエラーが発生しました。\n詳細はログファイルを確認してください。", parent=root)
                    return # エラーが起きたらリロード処理に進まない
            elif action == 'delete':
                profile_to_delete = result['profile']
                path_to_delete = get_profile_path(profile_to_delete)
                shutil.rmtree(path_to_delete) # ディレクトリごと削除
                if current_profile_name == profile_to_delete:
                    # 削除したのが現在使用中のプロファイルなら、デフォルトに戻す
                    current_profile_name = DEFAULT_PROFILE_NAME

            # 変更をsettings.jsonに保存し、アプリ全体をリロード
            settings['current_profile'] = current_profile_name
            save_settings(settings)
            reload_application_state(current_profile_name)

    def reload_application_state(profile_name):
        # LinkPopupに、新しいプロファイル名で再読み込みさせる
        if popup:
            popup.reload_profile(profile_name)
        
        # 事前キャッシュを再実行
        threading.Thread(target=lambda: preload_all_link_icons(profile_name), daemon=True).start()
        print(f"Switched to profile: {profile_name}")


    tray_thread = threading.Thread(target=run_tray, daemon=True)
    tray_thread.start()
    
    # --- 起動直後に全リンクのアイコンを事前キャッシュするバックグラウンドスレッド ---
    def preload_all_link_icons(profile_name):
        try:
            links_data = load_links_data(profile_name)
            if not links_data: return
            # 1. 現在のUIで実際に使われているサイズを取得
            # LinkPopupクラスが初期化されていれば、そこから正確なサイズを取得
            try:
                current_size = LinkPopup.current_icon_size
            except AttributeError:
                # まだLinkPopupがなければ、デフォルト設定から推測
                font_metrics = tkfont.Font(family=settings['font'], size=settings['size']).metrics()
                current_size = round_to_step(font_metrics.get('ascent', 16))

            # 2. キャッシュするサイズのリストを作成（現在のサイズを先頭に）
            standard_sizes = {8, 12, 16, 24, 32} # 一般的なサイズ
            preload_sizes = sorted(list({current_size} | standard_sizes))

            print(f"Preloading icons for sizes: {preload_sizes}") # デバッグ用

            for group in links_data:
                for link in group.get('links', []):
                    path = link.get('path', '')
                    if not path: continue
                    
                    for size in preload_sizes:
                        # ★ここでのアイコン取得はキャッシュ目的（UIには影響しない）
                        if path.startswith(('http://', 'https://')):
                            try:
                                get_web_icon(path, size=size)
                            except Exception: # ここでのエラーはログ不要（キャッシュ試行なので）
                                pass
                        else:
                            try:
                                get_file_icon(path, size=size)
                            except Exception:
                                pass
        except Exception as e:
            logging.warning(f"[preload] Preload thread failed: {e}")

    threading.Thread(target=lambda: preload_all_link_icons(current_profile_name), daemon=True).start()

    # 1. ルートウィンドウを先に作成する
    root = tk.Tk()
    root.withdraw() # すぐに非表示にする

    # 2. ルートウィンドウが作成された後で、アイコンを読み込む
    app_icon = None
    try:
        # まずは外部のicon.pngファイルを試す（カスタマイズ用）
        icon_path = os.path.join(BASE_DIR, "icon.png")
        if os.path.exists(icon_path):
            app_icon = tk.PhotoImage(file=icon_path)
        else:
            # ファイルがなければ、コードに埋め込まれたBase64データから復元
            app_icon = tk.PhotoImage(data=ICON_BASE64)
            
        root.iconphoto(True, app_icon)
    except tk.TclError:
        try:
            img = Image.open(icon_path)
            app_icon = ImageTk.PhotoImage(img)
            root.iconphoto(True, app_icon)
        except Exception as e:
            logging.warning(f"Failed to load application icon with Pillow: {e}")
    except Exception as e:
        logging.warning(f"icon.png not found or failed to load: {e}")
        
    # 3. アイコンを読み込んだ後で、それを利用するウィジェットを作成する
    popup = LinkPopup(root, settings, current_profile_name)
    
    # 4. メインループを開始
    try:
        root.mainloop()
    finally:
        # --- すべてのFileHandlerを明示的にclose & remove ---
        logger = logging.getLogger()
        handlers = logger.handlers[:]
        for handler in handlers:
            try:
                handler.close()
            except Exception:
                pass
            logger.removeHandler(handler)
        logging.shutdown()
    
if __name__ == "__main__":
    main()