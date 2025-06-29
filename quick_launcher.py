import os
import sys
import json
import tkinter as tk
from tkinter import messagebox, simpledialog, colorchooser, ttk, font as tkfont
from PIL import Image, ImageTk, ImageDraw
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

DEFAULT_SETTINGS = {
    'font': 'Yu Gothic UI',
    'size': 11,
    'font_color': '#000000',
    'bg': '#f0f0f0',
    'border_color': '#666666', # デフォルトのボーダー色
    'use_online_favicon': True,
}

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

def load_links_data():
    if os.path.exists(LINKS_FILE):
        try:
            with open(LINKS_FILE, 'r', encoding='utf-8') as f:
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
        # ファイルが存在しない場合は、デフォルトの空データで新規作成する
        logging.info("Links file not found, creating a new one.")
        default_links = [{"group": "マイリンク", "links": []}]
        save_links_data(default_links)
        return default_links

def save_links_data(links):
    with open(LINKS_FILE, 'w', encoding='utf-8') as f:
        json.dump(links, f, ensure_ascii=False, indent=2)

def open_link(path):
    try:
        if path.startswith(('http://', 'https://')):
            webbrowser.open(path)
        else:
            os.startfile(path)
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
            self.original_groups.append(new_group)
            # 表示データにも追加
            self.groups.append(new_group.copy())

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
        
        new_link = {'name': name, 'path': path}

        # --- ★★★ データ同期ロジック ★★★ ---
        # 1. 表示されているグループ名を取得
        current_group_name = self.groups[self.selected_group]['group']
        
        # 2. original_groups から該当グループを探し、そこに新しいリンクを追加
        for g in self.original_groups:
            if g['group'] == current_group_name:
                g['links'].append(new_link)
                break
        # 3. 表示用のgroupsにも追加
        self.groups[self.selected_group]['links'].append(new_link)
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
    
    def __init__(self, master, settings):
        super().__init__(master)
        if app_icon:
            self.iconphoto(True, app_icon)
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

    def reload_links(self):
        self.link_items.clear()
        self.group_map.clear()
        groups_data = load_links_data()
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
    settings = load_settings()

    def run_tray():
        try:
            try: image = Image.open(os.path.join(BASE_DIR, "icon.png"))
            except FileNotFoundError:
                image = Image.new('RGB', (64, 64), 'blue'); ImageDraw.Draw(image).text((10, 20), "QL", fill="white")
            
            def show_popup_action(icon=None): root.after(0, popup.show)
            def edit_links_action(icon=None): root.after(0, open_links_editor)
            def settings_action(icon=None): root.after(0, open_settings_dialog)
            def exit_action(icon=None):
                icon.stop()
                root.after(0, root.quit)

            menu = Menu(item('リンクを表示', show_popup_action, default=True), item('リンク編集', edit_links_action),
                        item('設定', settings_action), Menu.SEPARATOR, item('終了', exit_action))
            icon = Icon("QuickLauncher", image, "QuickLauncher", menu)
            icon.run()
        except Exception as e:
            logging.error(f"Tray icon failed: {e}", exc_info=True)
        if root: root.quit()

    def open_links_editor():
        links_data = load_links_data()
        if links_data is None: 
            links_data = [{"group": "マイリンク", "links": []}]
        
        # ダイアログをインスタンス化するだけで、__init__内のwait_windowで待機が始まる
        dialog = LinksEditDialog(root, links_data, settings)
        
        # ダイアログが閉じた後、結果を見て処理を続ける
        if dialog.result is not None:
            save_links_data(dialog.result)
            if popup: 
                popup.reload_links()

    def open_settings_dialog():
        dialog = SettingsDialog(root, settings)
        if dialog.result is not None:
            settings.clear()
            settings.update(dialog.result)
            if popup: popup.apply_settings(settings)

    tray_thread = threading.Thread(target=run_tray, daemon=True)
    tray_thread.start()
    
    # --- 起動直後に全リンクのアイコンを事前キャッシュするバックグラウンドスレッド ---
    def preload_all_link_icons():
        try:
            links_data = load_links_data()
            preload_sizes = [12, 16, 20, 24, 28, 32]
            for group in links_data:
                for link in group.get('links', []):
                    path = link.get('path', '')
                    for size in preload_sizes:
                        if path.startswith('http'):
                            try:
                                get_web_icon(path, size=size)
                            except Exception as e:
                                logging.info(f"[preload] get_web_icon failed: {path} ({size}px): {e}")
                        else:
                            try:
                                get_file_icon(path, size=size)
                            except Exception as e:
                                logging.info(f"[preload] get_file_icon failed: {path} ({size}px): {e}")
        except Exception as e:
            logging.warning(f"[preload] preload_all_link_icons failed: {e}")
    threading.Thread(target=preload_all_link_icons, daemon=True).start()

    # 1. ルートウィンドウを先に作成する
    root = tk.Tk()
    root.withdraw() # すぐに非表示にする

    # 2. ルートウィンドウが作成された後で、アイコンを読み込む
    app_icon = None
    try:
        icon_path = os.path.join(BASE_DIR, "icon.png")
        app_icon = tk.PhotoImage(file=icon_path)
        # ルートウィンドウ自身にもアイコンを設定しておく
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
    popup = LinkPopup(root, settings)
    
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