### ChroLens_Clear1.3 
### 2025/06/23 By Lucienwooo
### pyinstaller --onedir --noconsole --add-data "Nekoneko.ico;." --icon=Nekoneko.ico --hidden-import=win32timezone ChroLens_Clear.py
##### 尚無問題
import os
import json
from pathlib import Path
import ttkbootstrap as ttk
import win32gui
import win32con
import time
import tkinter as tk  # 新增

# 修正：設定檔路徑永遠指向程式所在資料夾
APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(APP_DIR, "config.json")

def set_language(lang):
    global L
    L = LANGUAGES[lang]
    save_language(lang)
    # 更新 UI 文字
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
    # 預設視窗名稱
    entry_windows[0].delete(0, "end")
    entry_windows[0].insert(0, L["default_window_name"])


def on_save():
    save_config()

def on_execute():
    generate_and_run_ahk()

def close_window_by_title(title):
    def enum_handler(hwnd, result):
        if win32gui.IsWindowVisible(hwnd):
            window_text = win32gui.GetWindowText(hwnd)
            if window_text == title:
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                result.append(hwnd)
    result = []
    win32gui.EnumWindows(enum_handler, result)
    return result

def close_window_by_title_partial(keyword):
    keyword_lower = keyword.lower()
    def enum_handler(hwnd, result):
        if win32gui.IsWindowVisible(hwnd):
            window_text = win32gui.GetWindowText(hwnd)
            if keyword_lower in window_text.lower():
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                result.append(hwnd)
    result = []
    win32gui.EnumWindows(enum_handler, result)
    return result

# 儲存與載入 config 的函式不變

def early_auto_run():
    """在 UI 初始化後排程自動關閉功能（不阻塞 UI）"""
    if not Path(CONFIG_FILE).exists():
        return
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        config = json.load(f)
    if config.get("auto_run", 0):
        window_titles = config.get("window_titles", [])
        try:
            repeat = int(config.get("repeat", 1))
        except Exception:
            repeat = 1
        try:
            interval = int(config.get("interval", 0))
        except Exception:
            interval = 0
        try:
            delay = int(config.get("delay", 0))
        except Exception:
            delay = 0
        try:
            auto_close_sec = int(config.get("auto_close", 0))
        except Exception:
            auto_close_sec = 0

        def do_close(count=0):
            for title in window_titles:
                if title.strip():
                    close_window_by_title_partial(title)
                    if interval > 0:
                        time.sleep(interval)
            count += 1
            if repeat == 0:
                return
            if count >= repeat:
                # 自動關閉主程式
                sec = auto_close_sec
                if sec:
                    if sec < 1:
                        sec = 1
                    elif sec > 99:
                        sec = 99
                    root.after(sec * 1000, root.destroy)
                return
            # 下一次
            root.after(delay * 1000, lambda: do_close(count))

        # 先立即執行一次
        do_close(0)
        # 若 repeat > 1，才繼續排程
        if repeat > 1:
            root.after(delay * 1000, lambda: do_close(1))

# 先執行功能
# early_auto_run()  # <-- 移除這一行

# 設定 icon 路徑（建議用絕對路徑，避免 pyinstaller 打包後找不到）
ICON_PATH = os.path.join(APP_DIR, "Nekoneko.ico")

# 語言對應字典
LANGUAGES = {
    "繁體中文": {
        "search_window": "搜尋視窗",
        "close_window_count": "關閉視窗數",
        "window_title": "視窗名稱",
        "execution_key": "執行快捷鍵",
        "delay": "延遲執行",
        "auto_run": "啟動時執行",
        "repeat": "次重複後自動關閉",
        "repeat_tip": "輸入N，將會延遲N秒執行一次，重複N次後自動關閉\n(每次間隔等於延遲秒數)",
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
        "repeat": "Auto Close After N Times",
        "repeat_tip": "Enter N: will execute once every N seconds, repeat N times then auto close.\n(Interval equals delay seconds)",
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
        "repeat_tip": "Nを入力するとN秒ごとに実行し、N回後に自動終了します\n(間隔は遅延秒数と同じ)",
        "save": "保存",
        "execute": "実行",
        "close": "閉じる",
        "language": "Language",
        "default_window_name": "",
        "search_window_title": "ウィンドウ検索",
    }
}

def save_language(lang):
    try:
        if Path(CONFIG_FILE).exists():
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
        else:
            config = {}
        config["language"] = lang
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print("語言儲存失敗：", e)

def load_language():
    if Path(CONFIG_FILE).exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)
        return config.get("language", "繁體中文")
    return "繁體中文"

current_language = load_language()
L = LANGUAGES[current_language]

# 再初始化 UI
root = ttk.Window(themename="superhero")
root.tk_setPalette(background="#172B4B")
root.title("ChroLens_Clear1.3")

# 設定主視窗 icon
try:
    if os.path.exists(ICON_PATH):
        root.iconbitmap(ICON_PATH)
    else:
        print("圖示檔案不存在，將不設置圖示。")
except Exception as e:
    print("設定圖示失敗：", e)

style = ttk.Style()
style.configure("TLabel", foreground="#F5D07D", font=("微軟正黑體", 14, "bold"))
style.configure("TButton", foreground="#F5D07D", font=("微軟正黑體", 14, "bold"))

execution_key = "F8"

def save_config():
    config = {
        "num_windows": entry_num_windows.get(),
        "window_titles": [entry.get() for entry in entry_windows],
        "execution_key": execution_key_var.get(),
        "auto_run": auto_run_var.get(),
        # "auto_close": auto_close_var.get(),  # 移除自動退出
        "delay": delay_var.get(),
        "repeat": repeat_var.get(),
        # "interval": interval_var.get(),      # 秮除關閉間隔
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

def load_config():
    if not Path(CONFIG_FILE).exists():
        return
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        config = json.load(f)
    entry_num_windows.set(config.get("num_windows", "1"))
    execution_key_var.set(config.get("execution_key", "F8"))
    auto_run_var.set(config.get("auto_run", 0))
    # auto_close_var.set(config.get("auto_close", ""))  # 移除自動退出
    delay_var.set(config.get("delay", "0"))
    repeat_var.set(config.get("repeat", "1"))
    # interval_var.set(config.get("interval", "0"))     # 秮除關閉間隔
    window_titles = config.get("window_titles", [])
    for i, title in enumerate(window_titles):
        if i < len(entry_windows):
            entry_windows[i].delete(0, "end")
            entry_windows[i].insert(0, title)
    update_window_inputs()
    if auto_run_var.get():
        root.after(1000, generate_and_run_ahk)

def generate_and_run_ahk():
    try:
        delay = int(delay_var.get())
        if delay < 0:
            delay = 0
        elif delay > 600:
            delay = 600
    except ValueError:
        delay = 0

    try:
        repeat = int(repeat_var.get())
        if repeat < 0:
            repeat = 0
    except ValueError:
        repeat = 0

    try:
        num_windows = int(entry_num_windows.get())
        window_titles = [entry_windows[i].get() for i in range(num_windows) if entry_windows[i].get().strip()]
        if not window_titles:
            return

        def do_close(count=0):
            for title in window_titles:
                if title.strip():
                    close_window_by_title_partial(title)
            count += 1
            if repeat == 0:
                # 0 代表不自動關閉，只執行一次，不關閉主程式
                pass
            elif count < repeat:
                root.after(delay * 1000, lambda: do_close(count))
            else:
                root.after(1000, root.destroy)

        if delay > 0:
            root.after(delay * 1000, lambda: do_close(0))
        else:
            do_close(0)

    except ValueError:
        pass

def update_execution_key(*args):
    global execution_key
    execution_key = execution_key_var.get()
    if keyboard:
        keyboard.clear_all_hotkeys()
        keyboard.add_hotkey(execution_key, generate_and_run_ahk)
    save_config()

def update_window_inputs(*args):
    try:
        num = int(entry_num_windows.get())
        if num > 10:
            num = 10
            entry_num_windows.set("10")
        for i in range(10):
            if i < num:
                entry_windows[i].grid()
            else:
                entry_windows[i].grid_remove()
        save_config()
    except ValueError:
        pass


# === 新增「搜尋視窗」按鈕在 row=0 ===
def open_window_search_dialog():
    # 建立一個與主視窗同高、寬度相同的新 Toplevel 視窗，位置在主視窗右側
    search_win = tk.Toplevel(root)
    search_win.title("搜尋視窗")
    # 設定 icon 與主程式一致
    try:
        if os.path.exists(ICON_PATH):
            search_win.iconbitmap(ICON_PATH)
    except Exception as e:
        print("搜尋視窗設定圖示失敗：", e)
    # 取得主視窗位置與大小，搜尋視窗顯示在右側
    root.update_idletasks()
    w = root.winfo_width()
    h = root.winfo_height()
    x = root.winfo_x()
    y = root.winfo_y()
    search_win.geometry(f"{w}x{h}+{x + w + 20}+{y}")
    search_win.transient(root)
    # 不要 grab_set，這樣主視窗可以互動
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

    # 取得所有可見視窗標題，排除系統/隱藏視窗
    import win32gui
    def get_window_titles():
        titles = []
        exclude_keywords = [
            "設定", "Windows 輸入體驗", "Program Manager", "LockApp", "Default IME", "Shell_TrayWnd",
            "Cortana", "Microsoft Text Input Application", "SearchUI", "Start menu", "Task Switching"
        ]
        def enum_handler(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                # 過濾空白、系統視窗
                if title and title.strip():
                    if not any(ex in title for ex in exclude_keywords):
                        titles.append(title)
        win32gui.EnumWindows(enum_handler, None)
        return titles

    # 拖移相關
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
    for row, title in enumerate(get_window_titles()):
        lbl = ttk.Label(inner_frame, text=title, anchor="w", font=("微軟正黑體", 12), background="#222", foreground="#fff")
        lbl.grid(row=row, column=0, sticky="ew", padx=4, pady=2)
        lbl.bind("<ButtonPress-1>", lambda e, t=title: on_label_drag_start(e, t))

    # 關閉按鈕
    ttk.Button(search_win, text="關閉", command=search_win.destroy).pack(pady=8)


# row 0: 語言下拉選單 + 搜尋視窗按鈕
language_var = ttk.StringVar(value=L["language"])
language_combobox = ttk.Combobox(root, textvariable=language_var, values=["日本語", "English", "繁體中文"], width=8, state="readonly")
language_combobox.grid(row=0, column=0, padx=5, pady=5, sticky="w")
def on_language_selected(event):
    set_language(language_var.get())
language_combobox.bind("<<ComboboxSelected>>", on_language_selected)

search_btn = ttk.Button(root, text=L["search_window"], style="info.TButton", command=open_window_search_dialog, width=8)
search_btn.grid(row=0, column=1, padx=5, pady=5, sticky="e")

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

# 預設值
entry_num_windows.set("1")
execution_key_var.set("F8")
auto_run_var.set(0)
delay_var.set("0")
repeat_var.set("0")
for entry in entry_windows:
    entry.delete(0, "end")  # 不填入預設視窗名稱

# 啟動時自動執行一次
def startup_execute():
    generate_and_run_ahk()
root.after(100, startup_execute)

# auto_run 勾選時立即執行
def on_auto_run_checked(*args):
    if auto_run_var.get():
        generate_and_run_ahk()
auto_run_var.trace_add("write", on_auto_run_checked)

load_config()
early_auto_run()
root.mainloop()


