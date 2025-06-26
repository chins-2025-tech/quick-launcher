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
logging.basicConfig(filename='app_errors.log', level=logging.WARNING,
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
_default_browser_icon = None

# --- ヘルパー関数 ---
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
    リソース解放のタイミングを修正した最終確定版。
    """
    icon_info = ICONINFO()
    if not user32.GetIconInfo(hIcon, ctypes.byref(icon_info)):
        if destroy_after: user32.DestroyIcon(hIcon)
        return None

    hBitmap = icon_info.hbmColor
    if not hBitmap:
        if destroy_after: user32.DestroyIcon(hIcon)
        return None

    # --- ここからが画像データ抽出の本番 ---
    tk_icon = None
    try:
        bitmap = BITMAP()
        gdi32.GetObjectW(hBitmap, ctypes.sizeof(bitmap), ctypes.byref(bitmap))
        width, height = bitmap.bmWidth, bitmap.bmHeight

        hdc = user32.GetDC(None)
        mem_dc = gdi32.CreateCompatibleDC(hdc)
        mem_bmp = gdi32.CreateCompatibleBitmap(hdc, width, height)
        gdi32.SelectObject(mem_dc, mem_bmp)
        user32.DrawIconEx(mem_dc, 0, 0, hIcon, width, height, 0, None, 3) # DI_NORMAL
        
        bmp_str = ctypes.create_string_buffer(width * height * 4)
        gdi32.GetBitmapBits(mem_bmp, len(bmp_str), bmp_str)
        
        img = Image.frombuffer("RGBA", (width, height), bmp_str, "raw", "BGRA", 0, 1)
        
        # ★★★ PhotoImageを先に完全に作成する ★★★
        tk_icon = ImageTk.PhotoImage(img.resize((size, size), Image.LANCZOS))
        
        # リソース解放
        gdi32.DeleteObject(mem_bmp)
        gdi32.DeleteDC(mem_dc)
        user32.ReleaseDC(None, hdc)
    
    finally:
        # GDIオブジェクトと、必要であればアイコンハンドルを解放
        if hBitmap: gdi32.DeleteObject(hBitmap)
        if icon_info.hbmMask: gdi32.DeleteObject(icon_info.hbmMask)
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
    
    # ダミーのフォルダ属性で、Windowsに標準のフォルダアイコンを要求する
    res = shell32.SHGetFileInfoW(
        "dummy_folder",
        FILE_ATTRIBUTE_DIRECTORY,
        ctypes.byref(info),
        ctypes.sizeof(info),
        flags
    )
    
    tk_icon = None
    if res and info.hIcon:
        # 修正済みのヘルパー関数を呼び出す
        # SHGetFileInfoWで取得したハンドルは破棄する必要がある (destroy_after=True)
        tk_icon = _hicon_to_photoimage(info.hIcon, size, destroy_after=True)

    # 取得に失敗した場合のフォールバック
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

def get_file_icon(path, size=20):
    """ファイルパスからアイコンを取得する。パスが存在しない場合は警告アイコンを返す。"""

    # まず、ファイル/フォルダの存在を確認する
    if not os.path.exists(path):
        return get_system_warning_icon(size)

    key = (path, size)
    if key in _icon_cache: return _icon_cache[key]
    info = SHFILEINFO()
    res = shell32.SHGetFileInfoW(path, 0, ctypes.byref(info), ctypes.sizeof(info), 0x100 | 0x1)
    tk_icon = _hicon_to_photoimage(info.hIcon, size) if res and info.hIcon else get_system_folder_icon(size)
    _icon_cache[key] = tk_icon
    return tk_icon

def get_web_icon(url, size=20):
    """
    URLからファビコンを取得する。
    settingsに応じてオンライン(Google)/オフライン(直接取得)を切り替える。
    """
    global _default_browser_icon

    if not url:
        # URLが無効な場合は、デフォルトブラウザアイコンを返すしかない
        if _default_browser_icon is None: _default_browser_icon = _get_or_create_default_browser_icon(size)
        return _default_browser_icon

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
        if _default_browser_icon is None: _default_browser_icon = _get_or_create_default_browser_icon(size)
        tk_icon = _default_browser_icon

    _icon_cache[key] = tk_icon
    return tk_icon

def _get_or_create_default_browser_icon(size=20):
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
            # パスはあるがファイルがない異常事態
            return get_system_warning_icon(size)
    except Exception:
        # レジストリ検索失敗など
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
    LINK_ICON_SIZE = 18
    LINK_ROW_HEIGHT = 24

    def __init__(self, parent, groups, settings):
        super().__init__(parent)

        # ★最重要ポイント1: transient を削除
        # 親(root)が非表示のため、transient を設定するとこのウィンドウも非表示になってしまう。
        # self.transient(parent) 

        self.title("リンクの編集")
        if 'app_icon' in globals() and app_icon:
            self.iconphoto(True, app_icon)

        self.groups = [dict(group) for group in groups]
        self.settings = settings
        self.selected_group = 0
        self.selected_link = None
        self.icon_refs = []
        self.result = None

        self.resizable(True, True)
        self.minsize(550, 400)

        # --- レイアウト設定 ---
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        main_pane = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_pane.grid(row=0, column=0, sticky="nsew")

        title_font = (self.settings['font'], self.settings['size'] - 1)
        content_font = (self.settings['font'], self.settings['size'])

        # --- 左ペイン：グループ一覧 ---
        group_pane = tk.LabelFrame(main_pane, text="グループ", font=title_font, bd=1, padx=2, pady=2)
        group_pane.grid_rowconfigure(0, weight=1)
        group_pane.grid_columnconfigure(0, weight=1)
        main_pane.add(group_pane, weight=1)
        group_scrollbar = tk.Scrollbar(group_pane, orient="vertical")
        self.group_listbox = tk.Listbox(group_pane, yscrollcommand=group_scrollbar.set, exportselection=False, font=content_font)
        self.group_listbox.grid(row=0, column=0, sticky="nsew")
        group_scrollbar.config(command=self.group_listbox.yview)
        group_scrollbar.grid(row=0, column=1, sticky="ns")
        self.group_listbox.bind('<<ListboxSelect>>', self.on_group_select)
        group_btns_frame = tk.Frame(group_pane)
        group_btns_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
        tk.Button(group_btns_frame, text="追加", command=self.add_group).pack(side="left")
        tk.Button(group_btns_frame, text="名変更", command=self.rename_group).pack(side="left")
        tk.Button(group_btns_frame, text="削除", command=self.delete_group).pack(side="left")
        tk.Button(group_btns_frame, text="↑", command=self.move_group_up).pack(side="left")
        tk.Button(group_btns_frame, text="↓", command=self.move_group_down).pack(side="left")
        # --- 右ペイン：リンク一覧 ---
        link_pane = tk.LabelFrame(main_pane, text="リンク", font=title_font, bd=1, padx=2, pady=2)
        link_pane.grid_rowconfigure(1, weight=1)
        link_pane.grid_columnconfigure(0, weight=1)
        main_pane.add(link_pane, weight=4)
        link_addr_row = tk.Frame(link_pane)
        link_addr_row.grid(row=0, column=0, columnspan=2, sticky="ew", padx=2, pady=2)
        link_addr_row.grid_columnconfigure(0, weight=1)
        link_addr_row.grid_rowconfigure(0, weight=1) 
        self.link_addr_var = tk.StringVar()
        self.link_addr_entry = ttk.Entry(link_addr_row, textvariable=self.link_addr_var, font=content_font )
        self.link_addr_entry.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        self.save_addr_btn = tk.Button(link_addr_row, text="保存", command=self.save_link_addr)
        self.save_addr_btn.grid(row=0, column=1, sticky="e")
        self.link_addr_entry.bind("<FocusIn>", self.on_link_addr_focus)
        link_scrollbar = tk.Scrollbar(link_pane, orient="vertical")
        self.link_canvas = tk.Canvas(link_pane, bg="#ffffff", highlightthickness=0)
        self.link_canvas.grid(row=1, column=0, sticky="nsew")
        link_scrollbar.config(command=self.link_canvas.yview)
        link_scrollbar.grid(row=1, column=1, sticky="ns")
        self.link_canvas.configure(yscrollcommand=link_scrollbar.set)
        self.link_canvas.bind("<Configure>", lambda e: self.refresh_link_list())
        self.link_canvas.bind("<MouseWheel>", self._on_link_canvas_mousewheel)
        self.link_canvas.bind("<Button-1>", self.on_link_canvas_click)
        self.link_canvas.bind("<Double-Button-1>", self.on_link_canvas_double)
        #self.link_canvas.bind("<Motion>", self.on_canvas_motion) # マウス移動
        #self.link_canvas.bind("<Leave>", self.on_canvas_leave)   # マウスがCanvasから出た時
        #self.tooltip = ToolTip(self.link_canvas)
        #self.canvas_item_map = {} # {canvas_item_id: path_string}
        link_btns = tk.Frame(link_pane)
        link_btns.grid(row=2, column=0, columnspan=2, sticky="w", pady=(5,0))
        tk.Button(link_btns, text="追加", command=self.add_link).pack(side="left")
        tk.Button(link_btns, text="名変更", command=self.rename_link).pack(side="left")
        tk.Button(link_btns, text="削除", command=self.delete_link).pack(side="left")
        tk.Button(link_btns, text="↑", command=self.move_link_up).pack(side="left")
        tk.Button(link_btns, text="↓", command=self.move_link_down).pack(side="left")

        # --- OK/Cancelボタン ---
        button_frame = tk.Frame(self)
        #button_frame.grid(row=1, column=0, sticky="e", padx=10, pady=(5, 10))
        button_frame.grid(row=1, column=0, pady=(5, 10))
        tk.Button(button_frame, text="OK", width=10, command=self.ok).pack(side="left", padx=5)
        tk.Button(button_frame, text="キャンセル", width=10, command=self.cancel).pack(side="left")

        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)

        # --- 初期化と表示処理 ---
        self.refresh_group_list()
        self.group_listbox.focus_set()

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
        self.result = self.groups.copy()
        self.destroy()

    def cancel(self, event=None):
        self.result = None
        self.destroy()

    def add_group(self):
        name = simpledialog.askstring("グループ名", "新しいグループ名:", parent=self)
        if name:
            self.groups.append({'group': name, 'links': []})
            self.selected_group = len(self.groups) - 1
            self.refresh_group_list()
            self.refresh_link_list()

    def rename_group(self):
        if not self.groups:
            return
        idx = self.selected_group
        new_name = simpledialog.askstring("グループ名変更", "新しいグループ名:", initialvalue=self.groups[idx]['group'], parent=self)
        if new_name:
            self.groups[idx]['group'] = new_name
            self.refresh_group_list()

    def delete_group(self):
        if not self.groups:
            return
        idx = self.selected_group
        del self.groups[idx]
        self.selected_group = max(0, self.selected_group - 1)
        self.refresh_group_list()
        self.refresh_link_list()

    def move_group_up(self):
        idx = self.selected_group
        if idx > 0:
            self.groups[idx-1], self.groups[idx] = self.groups[idx], self.groups[idx-1]
            self.selected_group -= 1
            self.refresh_group_list()

    def move_group_down(self):
        idx = self.selected_group
        if idx < len(self.groups)-1:
            self.groups[idx+1], self.groups[idx] = self.groups[idx], self.groups[idx+1]
            self.selected_group += 1
            self.refresh_group_list()

    def add_link(self):
        if not self.groups:
            return
        name = simpledialog.askstring("リンク名", "新しいリンク名:", parent=self)
        if not name:
            return
        path = self.ask_dialog(self, "リンク先", "リンク先パスまたはURL:")
        if name and path:
            self.groups[self.selected_group]['links'].append({'name': name, 'path': path})
            self.selected_link = len(self.groups[self.selected_group]['links'])-1
            self.refresh_link_list()

    def rename_link(self):
        if not self.groups or self.selected_link is None:
            return
        links = self.groups[self.selected_group]['links']
        idx = self.selected_link
        new_name = simpledialog.askstring("名前変更", "新しい名前:", initialvalue=links[idx]['name'], parent=self)
        if new_name:
            links[idx]['name'] = new_name
            self.refresh_link_list()

    def delete_link(self):
        if not self.groups or self.selected_link is None:
            return
        links = self.groups[self.selected_group]['links']
        del links[self.selected_link]
        self.selected_link = None
        self.refresh_link_list()

    def move_link_up(self):
        if not self.groups or self.selected_link is None or self.selected_link == 0:
            return
        links = self.groups[self.selected_group]['links']
        i = self.selected_link
        links[i-1], links[i] = links[i], links[i-1]
        self.selected_link -= 1
        self.refresh_link_list()

    def move_link_down(self):
        if not self.groups or self.selected_link is None:
            return
        links = self.groups[self.selected_group]['links']
        i = self.selected_link
        if i == len(links)-1:
            return
        links[i+1], links[i] = links[i], links[i+1]
        self.selected_link += 1
        self.refresh_link_list()

    def refresh_link_list(self):
        self.link_canvas.delete("all")
        self.icon_refs.clear()
        # self.canvas_item_map.clear()

        font_name = (self.settings['font'], self.settings['size'])
        font_path = (self.settings['font'], self.settings['size'])
        color_path = "#888888"

        # テキスト幅計算用にフォントオブジェクトを取得
        name_font_obj = tkfont.Font(font=font_name)

        # 選択グループがなければ何も表示しない
        if not self.groups or self.selected_group is None or self.selected_group >= len(self.groups):
            self.link_addr_entry.config(state="disabled")
            self.save_addr_btn.config(state="disabled")
            self.link_addr_var.set("")
            return
        links = self.groups[self.selected_group]['links']
        y = 2
        # --- Canvasの幅を取得してテキスト描画幅を調整 ---
        canvas_width = self.link_canvas.winfo_width() or 360
        for i, link in enumerate(links):
            path = link['path']
            if path.startswith('http'):
                icon = get_web_icon(path, size=self.LINK_ICON_SIZE)
            else:
                icon = get_file_icon(path, size=self.LINK_ICON_SIZE)
            if icon:
                self.link_canvas.create_image(8, y + self.LINK_ROW_HEIGHT // 2, image=icon, anchor="w")
                self.icon_refs.append(icon)
            # 選択枠
            if self.selected_link == i:
                self.link_canvas.create_rectangle(0, y, canvas_width, y+self.LINK_ROW_HEIGHT, outline="#3399ff", width=2)
            
            # 1. リンク名を描画し、その幅を取得
            name_x_start = 32
            name_id = self.link_canvas.create_text(name_x_start, y + self.LINK_ROW_HEIGHT // 2,
                                                   text=link['name'], anchor="w", font=font_name)
            
            # 2. 描画したリンク名の実際の幅を計算
            # name_font_obj.measure() を使うことで、表示されるテキストのピクセル幅がわかる
            name_width = name_font_obj.measure(link['name'])
            
            # 3. パスの描画開始位置を計算
            path_x_start = name_x_start + name_width + 25 # リンク名の右端 + 25pxのスペース

            # 4. パスを描画
            path_id = self.link_canvas.create_text(path_x_start, y + self.LINK_ROW_HEIGHT // 2,
                                                   text=path, anchor="w", font=font_path, fill=color_path)
            
            # # 5. ★重要: 描画したアイテムのIDとパスをマップに登録
            # #    これにより、後でマウスが乗った時にどのパスか判定できる
            # self.canvas_item_map[name_id] = path
            # self.canvas_item_map[path_id] = path

            y += self.LINK_ROW_HEIGHT

        self.link_canvas.config(scrollregion=(0,0,canvas_width,y))
        # 選択リンクがあれば編集可、なければ不可
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
            self.group_listbox.insert(tk.END, g['group'])
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
        idx = (event.y - 2) // self.LINK_ROW_HEIGHT
        links = self.groups[self.selected_group]['links']
        if 0 <= idx < len(links):
            self.selected_link = idx
            self.link_addr_var.set(links[idx]['path'])
        else:
            self.selected_link = None
        self.refresh_link_list()

    def on_link_canvas_double(self, event):
        idx = (event.y - 2) // self.LINK_ROW_HEIGHT
        links = self.groups[self.selected_group]['links']
        if 0 <= idx < len(links):
            os.startfile(links[idx]['path'])

    def save_link_addr(self):
        if not self.groups or self.selected_link is None:
            return
        links = self.groups[self.selected_group]['links']
        links[self.selected_link]['path'] = self.link_addr_var.get()
        self.refresh_link_list()

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
    GROUP_ROW_HEIGHT = 24
    ICON_SIZE = 16
    ICON_COLUMN_WIDTH = 24  # アイコンを描画する領域の幅 (アイコンサイズ + 余白)
    TEXT_LEFT_PADDING = 0   # アイコンとテキストの間の隙間
    
    def __init__(self, master, settings):
        super().__init__(master)
        if app_icon:
            self.iconphoto(True, app_icon)
        self.settings = settings.copy()
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
        self.folder_icon = get_system_folder_icon(size=16)
        
        self.reload_links()
        self.apply_settings(self.settings)

    def apply_settings(self, settings):
        self.settings = settings.copy()
        border_color = self.settings.get('border_color', DEFAULT_SETTINGS['border_color'])
        content_bg_color = self.settings.get('bg', DEFAULT_SETTINGS['bg'])
        self.config(bg=border_color)
        self.canvas.config(bg=content_bg_color)
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
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        win_w = self.winfo_width()
        win_h = self.winfo_height()

        # 理想の表示座標を計算
        x = self.winfo_pointerx() - win_w // 2
        y = self.winfo_pointery() - win_h - 10 # カーソルの少し上に表示

        # X座標の調整 (画面の左右にはみ出さないように)
        if x + win_w > screen_w:
            x = screen_w - win_w - 5 # 右端に余裕を持たせる
        if x < 0:
            x = 5 # 左端に余裕を持たせる

        # Y座標の調整 (画面の上にはみ出さないように)
        # 下にはみ出すケースは、タスクバーがあるため通常は少ないが念のため考慮
        if y < 0:
            y = self.winfo_pointery() + 20 # カーソルの下側に表示位置を変更
        if y + win_h > screen_h:
            y = screen_h - win_h - 5 # 下端に余裕を持たせる
        # --- ここまでが座標調整ロジック ---

        self.geometry(f"+{x}+{y}")
        self.deiconify()
        self.lift()
        self.focus_force()

    def draw_list(self):
        self.canvas.delete("all")
        self.canvas.image_refs = [] 
        
        font_main = (self.settings['font'], self.settings['size'])
        font_metrics = tkfont.Font(font=font_main)
        maxlen = max((font_metrics.measure(g) for g in self.group_map), default=0)
        
        # --- 定数を使った計算 ---
        canvas_w = min(max(self.ICON_COLUMN_WIDTH + self.TEXT_LEFT_PADDING + maxlen, 120), 400)
        canvas_h = len(self.group_map) * self.GROUP_ROW_HEIGHT
        self.canvas.config(width=canvas_w, height=canvas_h)
        
        y = 0
        for i, group in enumerate(self.group_map):
            bg_color = "#eaf6ff" if self.hover_group == i else self.settings['bg']
            self.canvas.create_rectangle(0, y, canvas_w, y + self.GROUP_ROW_HEIGHT, fill=bg_color, outline="")
            
            if self.folder_icon:
                # アイコンをアイコン領域の中央に配置
                icon_x = self.ICON_COLUMN_WIDTH // 2
                self.canvas.create_image(icon_x, y + self.GROUP_ROW_HEIGHT // 2, image=self.folder_icon, anchor="center")
                self.canvas.image_refs.append(self.folder_icon)
            
            # テキストの位置を定数から計算
            text_x = self.ICON_COLUMN_WIDTH + self.TEXT_LEFT_PADDING
            self.canvas.create_text(text_x, y + self.GROUP_ROW_HEIGHT // 2, text=group, anchor="w", font=font_main, fill=self.settings['font_color'])
            
            y += self.GROUP_ROW_HEIGHT
        
    def on_motion(self, event):
        row_h = self.GROUP_ROW_HEIGHT
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
        popup.icon_refs = []

        # --- forループ内をgridを使うように変更 ---
        for i, link in enumerate(links):
            # 各行のコンテナとなるFrame
            row = tk.Frame(frame, bg=self.settings['bg'])
            row.grid(row=i, column=0, sticky="ew") # gridで行を管理
            frame.grid_columnconfigure(0, weight=1) # 幅を広げる設定

            # gridでカラムの幅を設定
            row.grid_columnconfigure(0, minsize=self.ICON_COLUMN_WIDTH)
            row.grid_columnconfigure(1, weight=1)

            # アイコンの取得
            path = link['path']
            icon = get_web_icon(path, self.ICON_SIZE) if path.startswith('http') else get_file_icon(path, self.ICON_SIZE)
            if not icon: icon = self.folder_icon # フォールバック
            
            icon_label = tk.Label(row, image=icon, bg=self.settings['bg'])
            # gridで配置 (column 0)
            icon_label.grid(row=0, column=0, sticky="e", padx=(0, self.TEXT_LEFT_PADDING))
            popup.icon_refs.append(icon)
            
            text_label = tk.Label(row, text=link['name'], anchor="w", bg=self.settings['bg'], font=font_link, fg=self.settings['font_color'])
            # gridで配置 (column 1)
            text_label.grid(row=0, column=1, sticky="w")
            
            def on_enter(e, lbl=text_label): lbl.config(font=font_underline)
            def on_leave(e, lbl=text_label): lbl.config(font=font_link)
            def on_click(e, p=path): self.open_and_close(p)
            
            for w in [row, icon_label, text_label]:
                w.bind("<Enter>", on_enter); w.bind("<Leave>", on_leave); w.bind("<Button-1>", on_click)

        self.update_idletasks() # 親ウィンドウの位置・サイズを確定させる
        popup.update_idletasks() # サブポップアップのサイズを計算させる

        # --- ここからが座標調整ロジック ---
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        parent_x = self.winfo_rootx()
        parent_y = self.winfo_rooty()
        parent_w = self.winfo_width()
        popup_w = popup.winfo_width()
        popup_h = popup.winfo_height()

        # 理想の表示座標を計算 (親ウィンドウの右側)
        x = parent_x + parent_w
        y = parent_y + group_idx * self.GROUP_ROW_HEIGHT

        # X座標の調整 (画面右端にはみ出す場合)
        if x + popup_w > screen_w:
            # 親ウィンドウの左側に表示位置を変更
            x = parent_x - popup_w
        
        # Y座標の調整 (画面下端にはみ出す場合)
        if y + popup_h > screen_h:
            # はみ出した分だけ上にずらす
            y = screen_h - popup_h - 5
        if y < 0:
            y = 5
        # --- ここまでが座標調整ロジック ---

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

# --- アプリケーション実行部 ---
def main():
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
            def exit_action(icon=None): icon.stop()

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
    root.mainloop()
    logging.shutdown()
    
if __name__ == "__main__":
    main()