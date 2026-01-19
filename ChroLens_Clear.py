### ChroLens_Clear 1.4
### 2025/11/02 By Lucienwooo
### pyinstaller --onedir --noconsole --add-data "Nekoneko.ico;." --icon=Nekoneko.ico --hidden-import=win32timezone ChroLens_Clear.py
"""
ChroLens_Clear - 視窗自動關閉工具
功能：
- 批次關閉指定視窗
- 支援模糊匹配視窗標題
- 延遲執行與重複執行
- 多語言支援（繁中/英文/日文）
"""

VERSION = "1.4.0"
GITHUB_REPO = "Lucienwooo/ChroLens_Clear"

import os
import json
from pathlib import Path
import ttkbootstrap as ttk
import win32gui
import win32con
import time
import tkinter as tk

# 版本管理模組
try:
    from version_manager import VersionManager
    from version_info_dialog import VersionInfoDialog
    VERSION_MANAGER_AVAILABLE = True
except ImportError:
    print("版本管理模組未安裝，版本檢查功能將停用")
    VERSION_MANAGER_AVAILABLE = False

# ===== 常數定義 =====
APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(APP_DIR, "config.json")
ICON_PATH = os.path.join(APP_DIR, "Nekoneko.ico")

# 系統視窗排除關鍵字
SYSTEM_WINDOW_KEYWORDS = [
    "設定", "Windows 輸入體驗", "Windows Input Experience", 
    "Program Manager", "LockApp", "Default IME", "Shell_TrayWnd",
    "Cortana", "Microsoft Text Input Application", "SearchUI", 
    "Start menu", "Task Switching", "RuntimeBroker", "TextInputHost"
]

# ===== 語言管理 =====
LANGUAGES = {
    "繁體中文": {
        "search_window": "搜尋視窗",
        "close_window_count": "關閉視窗數",
        "window_title": "視窗名稱",
        "execution_key": "執行快捷鍵",
        "delay": "延遲執行(秒)",
        "auto_run": "啟動時執行",
        "repeat": "次重複後自動關閉",
        "repeat_tip": "輸入N，將會延遲N秒執行一次，重複N次後自動關閉\n(每次間隔等於延遲秒數)\n輸入0代表只執行一次",
        "save": "儲存",
        "execute": "執行",
        "close": "關閉",
        "language": "Language",
        "default_window_name": "",
        "search_window_title": "搜尋視窗",
    },
    "English": {
        "search_window": "Search",
        "close_window_count": "Windows",
        "window_title": "Window Title",
        "execution_key": "Hotkey",
        "delay": "Delay (sec)",
        "auto_run": "Auto Run on Startup",
        "repeat": "Repeat N Times Then Auto Close",
        "repeat_tip": "Enter N: execute once every N seconds, repeat N times then auto close.\n(Interval equals delay seconds)\n0 means execute once only",
        "save": "Save",
        "execute": "Execute",
        "close": "Close",
        "language": "Language",
        "default_window_name": "",
        "search_window_title": "Search Windows",
    },
    "日本語": {
        "search_window": "検索",
        "close_window_count": "閉じるウィンドウ数",
        "window_title": "ウィンドウ名",
        "execution_key": "実行ホットキー",
        "delay": "遅延実行(秒)",
        "auto_run": "起動時に自動実行",
        "repeat": "N回繰り返し後自動終了",
        "repeat_tip": "Nを入力するとN秒ごとに実行し、N回後に自動終了します\n(間隔は遅延秒数と同じ)\n0は1回のみ実行",
        "save": "保存",
        "execute": "実行",
        "close": "閉じる",
        "language": "Language",
        "default_window_name": "",
        "search_window_title": "ウィンドウ検索",
    }
}

def load_language():
    """載入語言設定"""
    if Path(CONFIG_FILE).exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
            return config.get("language", "繁體中文")
        except Exception as e:
            print(f"載入語言設定失敗: {e}")
    return "繁體中文"

def save_language(lang):
    """儲存語言設定"""
    try:
        config = {}
        if Path(CONFIG_FILE).exists():
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
        config["language"] = lang
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"語言儲存失敗: {e}")

# 初始化語言
current_language = load_language()
L = LANGUAGES[current_language]

def set_language(lang):
    """切換語言並更新 UI"""
    global L
    L = LANGUAGES[lang]
    save_language(lang)
    
    # 更新 UI 文字
    try:
        language_combobox.set(L["language"])
        search_btn.config(text=L["search_window"])
        close_window_count_label.config(text=L["close_window_count"])
        window_title_label.config(text=L["window_title"])
        execution_key_label.config(text=L["execution_key"])
        delay_label.config(text=L["delay"])
        auto_run_checkbutton.config(text=L["auto_run"])
        repeat_label.config(text=L["repeat"])
        save_button.config(text=L["save"])
        execute_button.config(text=L["execute"])
    except Exception as e:
        print(f"更新 UI 語言失敗: {e}")


# ===== 視窗操作函數 =====
def get_all_visible_windows():
    """取得所有可見視窗標題（排除系統視窗）"""
    titles = []
    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title and title.strip():
                # 過濾系統視窗
                if not any(keyword.lower() in title.lower() for keyword in SYSTEM_WINDOW_KEYWORDS):
                    titles.append(title)
    try:
        win32gui.EnumWindows(enum_handler, None)
    except Exception as e:
        print(f"列舉視窗失敗: {e}")
    return titles

def close_window_by_keyword(keyword):
    """根據關鍵字關閉視窗（部分匹配）"""
    if not keyword or not keyword.strip():
        return []
    
    keyword_lower = keyword.lower().strip()
    closed_windows = []
    
    def enum_handler(hwnd, result):
        try:
            if win32gui.IsWindowVisible(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                if keyword_lower in window_text.lower():
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    result.append(window_text)
        except Exception as e:
            print(f"關閉視窗失敗: {e}")
    
    try:
        win32gui.EnumWindows(enum_handler, closed_windows)
    except Exception as e:
        print(f"列舉視窗時發生錯誤: {e}")
    
    return closed_windows

# ===== 設定檔管理 =====
def save_config():
    """儲存設定檔"""
    try:
        config = {
            "language": current_language,
            "num_windows": entry_num_windows.get(),
            "window_titles": [entry.get() for entry in entry_windows],
            "execution_key": execution_key_var.get(),
            "auto_run": auto_run_var.get(),
            "delay": delay_var.get(),
            "repeat": repeat_var.get(),
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"儲存設定檔失敗: {e}")

def load_config():
    """載入設定檔"""
    if not Path(CONFIG_FILE).exists():
        return
    
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)
        
        entry_num_windows.set(config.get("num_windows", "1"))
        execution_key_var.set(config.get("execution_key", "F8"))
        auto_run_var.set(config.get("auto_run", 0))
        delay_var.set(config.get("delay", "0"))
        repeat_var.set(config.get("repeat", "0"))
        
        window_titles = config.get("window_titles", [])
        for i, title in enumerate(window_titles):
            if i < len(entry_windows):
                entry_windows[i].delete(0, "end")
                entry_windows[i].insert(0, title)
        
        update_window_inputs()
    except Exception as e:
        print(f"載入設定檔失敗: {e}")

# ===== 執行邏輯 =====
def execute_close_windows():
    """執行關閉視窗的主邏輯"""
    try:
        delay = max(0, min(600, int(delay_var.get())))
    except ValueError:
        delay = 0
    
    try:
        repeat = max(0, int(repeat_var.get()))
    except ValueError:
        repeat = 0
    
    try:
        num_windows = int(entry_num_windows.get())
        window_titles = [
            entry_windows[i].get().strip() 
            for i in range(num_windows) 
            if i < len(entry_windows) and entry_windows[i].get().strip()
        ]
        
        if not window_titles:
            return
        
        def do_close(count=0):
            """遞迴執行關閉視窗"""
            for title in window_titles:
                closed = close_window_by_keyword(title)
                if closed:
                    print(f"已關閉: {', '.join(closed)}")
            
            count += 1
            
            # repeat = 0 代表只執行一次，不關閉主程式
            if repeat == 0:
                return
            
            # 尚未達到重複次數，繼續排程
            if count < repeat:
                root.after(delay * 1000, lambda: do_close(count))
            else:
                # 達到重複次數，關閉主程式
                root.after(1000, root.destroy)
        
        # 延遲後開始執行
        if delay > 0:
            root.after(delay * 1000, lambda: do_close(0))
        else:
            do_close(0)
            
    except ValueError as e:
        print(f"執行失敗: {e}")

# ===== UI 回調函數 =====
def on_save():
    """儲存按鈕回調"""
    save_config()
    print("設定已儲存")

def on_execute():
    """執行按鈕回調"""
    execute_close_windows()

def update_execution_key(*args):
    """更新執行快捷鍵"""
    # 注意：這個程式不使用 keyboard 模組進行全域熱鍵
    # F8 等功能鍵僅作為顯示用途
    save_config()

def update_window_inputs(*args):
    """更新視窗輸入欄位數量"""
    try:
        num = min(10, max(1, int(entry_num_windows.get())))
        if num != int(entry_num_windows.get()):
            entry_num_windows.set(str(num))
        
        for i in range(10):
            if i < num:
                entry_windows[i].grid()
            else:
                entry_windows[i].grid_remove()
        save_config()
    except ValueError:
        pass

# ===== UI 初始化 =====
root = ttk.Window(themename="superhero")
root.tk_setPalette(background="#172B4B")
root.title("ChroLens_Clear 1.4")

# 設定主視窗 icon
try:
    if os.path.exists(ICON_PATH):
        root.iconbitmap(ICON_PATH)
except Exception as e:
    print(f"設定圖示失敗: {e}")

# 設定樣式
style = ttk.Style()
style.configure("TLabel", foreground="#F5D07D", font=("微軟正黑體", 14, "bold"))
style.configure("TButton", foreground="#F5D07D", font=("微軟正黑體", 14, "bold"))

# ===== 視窗搜尋對話框 =====
def open_window_search_dialog():
    """開啟視窗搜尋對話框"""
    search_win = tk.Toplevel(root)
    search_win.title(L["search_window_title"])
    
    # 設定 icon
    try:
        if os.path.exists(ICON_PATH):
            search_win.iconbitmap(ICON_PATH)
    except Exception as e:
        print(f"搜尋視窗設定圖示失敗: {e}")
    
    # 取得主視窗位置與大小，搜尋視窗顯示在右側
    root.update_idletasks()
    w = root.winfo_width()
    h = root.winfo_height()
    x = root.winfo_x()
    y = root.winfo_y()
    search_win.geometry(f"{w}x{h}+{x + w + 20}+{y}")
    search_win.transient(root)
    search_win.focus_set()
    search_win.resizable(True, True)

    # 建立滾動區域
    frame = ttk.Frame(search_win)
    frame.pack(fill="both", expand=True)
    canvas = tk.Canvas(frame, highlightthickness=0)
    canvas.pack(side="left", fill="both", expand=True)
    vsb = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
    vsb.pack(side="right", fill="y")
    canvas.configure(yscrollcommand=vsb.set)
    inner_frame = ttk.Frame(canvas)
    inner_id = canvas.create_window((0, 0), window=inner_frame, anchor="nw")
    def on_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(inner_id, width=canvas.winfo_width())
    inner_frame.bind("<Configure>", on_configure)

    # 拖移相關變數
    dragged_title = {"title": None}
    drag_label_popup = {"win": None}

    def on_label_drag_start(event, title):
        dragged_title["title"] = title
        # 懸浮視窗
        if drag_label_popup["win"]:
            drag_label_popup["win"].destroy()
        drag_label_popup["win"] = tk.Toplevel(search_win)
        drag_label_popup["win"].overrideredirect(True)
        drag_label_popup["win"].attributes("-topmost", True)
        label = ttk.Label(drag_label_popup["win"], text=title, background="#222", foreground="#fff", font=("微軟正黑體", 12))
        label.pack()
        def follow_mouse(ev):
            x = ev.x_root + 10
            y = ev.y_root + 10
            drag_label_popup["win"].geometry(f"+{x}+{y}")
        search_win.bind("<Motion>", follow_mouse)
        def on_drop(ev):
            # 檢查滑鼠下方是否為主程式的視窗名稱欄位
            for entry in entry_windows:
                x1 = entry.winfo_rootx()
                y1 = entry.winfo_rooty()
                x2 = x1 + entry.winfo_width()
                y2 = y1 + entry.winfo_height()
                if x1 <= ev.x_root <= x2 and y1 <= ev.y_root <= y2:
                    entry.delete(0, tk.END)
                    entry.insert(0, title)
            if drag_label_popup["win"]:
                drag_label_popup["win"].destroy()
                drag_label_popup["win"] = None
            search_win.unbind("<Motion>")
            search_win.unbind("<ButtonRelease-1>")
            dragged_title["title"] = None
        search_win.bind("<ButtonRelease-1>", on_drop)

    # 顯示所有視窗標題
    windows = get_all_visible_windows()
    if windows:
        for row, title in enumerate(windows):
            lbl = ttk.Label(
                inner_frame, text=title, anchor="w", 
                font=("微軟正黑體", 12), background="#222", foreground="#fff"
            )
            lbl.grid(row=row, column=0, sticky="ew", padx=4, pady=2)
            lbl.bind("<ButtonPress-1>", lambda e, t=title: on_label_drag_start(e, t))
    else:
        lbl = ttk.Label(
            inner_frame, text="未找到可用視窗", anchor="center",
            font=("微軟正黑體", 14), background="#222", foreground="#ff6b6b"
        )
        lbl.grid(row=0, column=0, sticky="ew", padx=4, pady=20)

    # 關閉按鈕
    ttk.Button(search_win, text=L["close"], command=search_win.destroy).pack(pady=8)


# ===== 版本管理功能 =====
def check_for_updates():
    """檢查更新"""
    if not VERSION_MANAGER_AVAILABLE:
        tk.messagebox.showinfo("提示", "版本管理功能未啟用")
        return
    
    try:
        # 初始化版本管理器
        version_manager = VersionManager(GITHUB_REPO, VERSION,
            logger=lambda msg: print(f"[版本管理] {msg}")
        )
        
        # 開啟版本資訊對話框
        dialog = VersionInfoDialog(
            parent=root,
            version_manager=version_manager,
            current_version=VERSION,
            on_update_callback=on_update_complete
        )
    except Exception as e:
        tk.messagebox.showerror("錯誤", f"檢查更新失敗：{e}")

def on_update_complete():
    """更新完成回調"""
    tk.messagebox.showinfo("提示", "更新完成！請重新啟動應用程式。")

def show_about():
    """顯示關於資訊 (使用統一對話框)"""
    check_for_updates()


# row 0: 語言下拉選單 + 搜尋視窗按鈕 + 關於 + 檢查更新
language_var = ttk.StringVar(value=L["language"])
language_combobox = ttk.Combobox(root, textvariable=language_var, values=["日本語", "English", "繁體中文"], width=8, state="readonly")
language_combobox.grid(row=0, column=0, padx=5, pady=5, sticky="w")
def on_language_selected(event):
    set_language(language_var.get())
language_combobox.bind("<<ComboboxSelected>>", on_language_selected)

search_btn = ttk.Button(root, text=L["search_window"], style="info.TButton", command=open_window_search_dialog, width=8)
search_btn.grid(row=0, column=1, padx=5, pady=5, sticky="e")

# 添加關於和檢查更新按鈕
about_btn = ttk.Button(root, text="關於", style="secondary.TButton", command=show_about, width=6)
about_btn.grid(row=0, column=2, padx=5, pady=5, sticky="e")

update_btn = ttk.Button(root, text="檢查更新", style="success.TButton", command=check_for_updates, width=8)
update_btn.grid(row=0, column=3, padx=5, pady=5, sticky="e")

# row 1 開始往下排
close_window_count_label = ttk.Label(root, text=L["close_window_count"])
close_window_count_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
entry_num_windows = ttk.Combobox(root, values=[str(i) for i in range(1, 11)], width=5, state="readonly")
entry_num_windows.grid(row=1, column=1, padx=5, pady=5, sticky="w")
entry_num_windows.set("1")
entry_num_windows.bind("<<ComboboxSelected>>", update_window_inputs)

window_title_label = ttk.Label(root, text=L["window_title"])
window_title_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
entry_windows = []
for i in range(10):
    entry = ttk.Entry(root, width=20)
    entry.grid(row=2 + i, column=1, padx=5, pady=5, sticky="w")
    entry_windows.append(entry)
    if i > 0:
        entry.grid_remove()
    # 新增右鍵清空功能
    def clear_entry(event, e=entry):
        e.delete(0, tk.END)
    entry.bind("<Button-3>", clear_entry)
entry_windows[0].insert(0, L["default_window_name"])

execution_key_label = ttk.Label(root, text=L["execution_key"])
execution_key_label.grid(row=12, column=0, padx=5, pady=5, sticky="w")
execution_key_var = ttk.StringVar(value=execution_key)
execution_key_combobox = ttk.Combobox(root, textvariable=execution_key_var, values=[f"F{i}" for i in range(1, 13)], width=5, state="readonly")
execution_key_combobox.grid(row=12, column=1, padx=5, pady=5, sticky="w")
execution_key_combobox.bind("<<ComboboxSelected>>", update_execution_key)

delay_label = ttk.Label(root, text=L["delay"])
delay_label.grid(row=13, column=0, padx=5, pady=5, sticky="w")
delay_var = ttk.StringVar(value="0")
delay_entry = ttk.Entry(root, textvariable=delay_var, width=4)
delay_entry.grid(row=13, column=1, padx=5, pady=5, sticky="w")

# 常駐模式選項區塊
persistent_frame = ttk.LabelFrame(root, borderwidth=2)
persistent_frame.grid(row=16, column=0, columnspan=2, padx=5, pady=5, sticky="we")

auto_run_var = ttk.IntVar(value=0)
auto_run_checkbutton = ttk.Checkbutton(
    persistent_frame, text=L["auto_run"], variable=auto_run_var, command=save_config, style="PersistentGreen.TCheckbutton"
)
auto_run_checkbutton.grid(row=0, column=0, columnspan=2, pady=5, sticky="w")

repeat_var = ttk.StringVar(value="1")
repeat_entry = ttk.Entry(persistent_frame, textvariable=repeat_var, width=4)
repeat_entry.grid(row=1, column=0, padx=5, pady=5, sticky="w")
repeat_label = ttk.Label(persistent_frame, text=L["repeat"], style="PersistentGreen.TLabel")
repeat_label.grid(row=1, column=1, padx=5, pady=5, sticky="w")
def show_tip(event):
    x = event.x_root + 10
    y = event.y_root + 10
    global tip_window
    tip_window = ttk.Toplevel(root)
    tip_window.wm_overrideredirect(True)
    tip_window.geometry(f"+{x}+{y}")
    label = ttk.Label(tip_window, text=L["repeat_tip"], background="#FFFACD", foreground="#000")
    label.pack()
def hide_tip(event):
    global tip_window
    if tip_window:
        tip_window.destroy()
        tip_window = None
tip_window = None
repeat_label.bind("<Enter>", show_tip)
repeat_label.bind("<Leave>", hide_tip)

# 儲存與執行按鈕
button_frame = ttk.Frame(root)
button_frame.grid(row=101, column=0, columnspan=2, pady=10)
save_button = ttk.Button(button_frame, text=L["save"], style="info.TButton", command=on_save, width=10)
save_button.pack(side="left", padx=10)
execute_button = ttk.Button(button_frame, text=L["execute"], style="success.TButton", command=on_execute, width=10)
execute_button.pack(side="left", padx=10)

# ===== 程式啟動邏輯 =====
# 載入設定檔
load_config()

# Auto-run 勾選時執行
def on_auto_run_checked(*args):
    """當 auto_run 被勾選時執行"""
    if auto_run_var.get():
        # 延遲執行，讓 UI 完全初始化
        root.after(500, execute_close_windows)

auto_run_var.trace_add("write", on_auto_run_checked)

# 如果啟動時 auto_run 已勾選，自動執行
if auto_run_var.get():
    root.after(500, execute_close_windows)

# 啟動主迴圈
root.mainloop()


