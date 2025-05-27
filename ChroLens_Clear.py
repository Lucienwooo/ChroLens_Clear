### ChroLens_Clear 1.0.0 
### 2025/05/26 By Lucienwooo
### pyinstaller --onefile --noconsole --add-data "Nekoneko.ico;." --icon=Nekoneko.ico --hidden-import=win32timezone ChroLens_Clear.py
##### 檢查若名稱完全相同時是否會全部關閉，或是只關閉一個
import os
import json
from pathlib import Path
import ttkbootstrap as ttk
import win32gui
import win32con
import time
import pygetwindow as gw

try:
    import keyboard
except ImportError:
    keyboard = None
    print("Error: 'keyboard' 庫未安裝。請使用 'pip install keyboard' 進行安裝。")

CONFIG_FILE = "config.json"

root = ttk.Window(themename="superhero")
root.tk_setPalette(background="#172B4B")
root.title("ChroLens_Clear v1.0")

icon_path_a = "./Nekoneko.ico"
icon_path_b = r"C:/Users/Lucien/Documents/GitHub/ChroLens_Clear/Nekoneko.ico"

if os.path.exists(icon_path_a):
    root.iconbitmap(icon_path_a)
elif os.path.exists(icon_path_b):
    root.iconbitmap(icon_path_b)
else:
    print("圖示檔案不存在，將不設置圖示。")

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
        "auto_close": auto_close_var.get(),
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
    auto_close_var.set(config.get("auto_close", 0))
    window_titles = config.get("window_titles", [])
    for i, title in enumerate(window_titles):
        if i < len(entry_windows):
            entry_windows[i].delete(0, "end")
            entry_windows[i].insert(0, title)
    update_window_inputs()
    update_execution_key()  # 新增：載入時綁定快捷鍵
    if auto_run_var.get():
        root.after(1000, generate_and_run_ahk)
    if auto_close_var.get():
        root.after(3000, root.destroy)

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
        repeat = 1

    try:
        num_windows = int(entry_num_windows.get())
        window_titles = [entry_windows[i].get() for i in range(num_windows) if entry_windows[i].get().strip()]
        if not window_titles:
            return
        try:
            interval = int(interval_var.get())
            if interval < 0:
                interval = 0
            elif interval > 99:
                interval = 99
        except ValueError:
            interval = 0

        def do_close():
            count = 0
            while True:
                for title in window_titles:
                    close_window_by_title_partial(title)
                    if interval > 0:
                        time.sleep(interval)
                count += 1
                if repeat == 0:
                    continue
                if count >= repeat:
                    break

        if delay > 0:
            root.after(delay * 1000, do_close)
        else:
            do_close()

        try:
            auto_close_sec = int(auto_close_var.get())
            if auto_close_sec < 1:
                auto_close_sec = 1
            elif auto_close_sec > 99:
                auto_close_sec = 99
            root.after(auto_close_sec * 1000, root.destroy)
        except Exception:
            pass

    except ValueError:
        pass

def bring_window_to_top_by_title_partial(keyword):
    """只用檔名（不含副檔名）比對視窗標題"""
    keyword = os.path.splitext(keyword)[0].lower()
    found = False
    for w in gw.getAllWindows():
        if w.title and keyword in w.title.lower():
            try:
                w.activate()
                w.restore()
                w.alwaysOnTop = True
                found = True
            except Exception as e:
                print(f"無法設為最上層: {w.title} {e}")
    return found

def bring_all_to_top():
    try:
        num_windows = int(entry_num_windows.get())
        window_titles = [entry_windows[i].get() for i in range(num_windows) if entry_windows[i].get().strip()]
        if not window_titles:
            return
        for title in window_titles:
            bring_window_to_top_by_title_partial(title)
    except Exception as e:
        print(f"Bring to top error: {e}")

def update_execution_key(*args):
    global execution_key
    execution_key = execution_key_var.get()
    if keyboard:
        keyboard.clear_all_hotkeys()
        # 綁定 bring_all_to_top 到快捷鍵
        keyboard.add_hotkey(execution_key, bring_all_to_top)
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

ttk.Label(root, text="關閉視窗數").grid(row=0, column=0, padx=5, pady=5, sticky="w")
entry_num_windows = ttk.Combobox(root, values=[str(i) for i in range(1, 11)], width=5, state="readonly")
entry_num_windows.grid(row=0, column=1, padx=5, pady=5, sticky="w")
entry_num_windows.set("1")
entry_num_windows.bind("<<ComboboxSelected>>", update_window_inputs)

ttk.Label(root, text="視窗名稱").grid(row=1, column=0, padx=5, pady=5, sticky="w")
entry_windows = []
for i in range(10):
    entry = ttk.Entry(root, width=20)
    entry.grid(row=1 + i, column=1, padx=5, pady=5, sticky="w")
    entry_windows.append(entry)
    if i > 0:
        entry.grid_remove()
entry_windows[0].insert(0, "chrome")

ttk.Label(root, text="執行快捷鍵").grid(row=11, column=0, padx=5, pady=5, sticky="w")
execution_key_var = ttk.StringVar(value=execution_key)
execution_key_combobox = ttk.Combobox(root, textvariable=execution_key_var, values=[f"F{i}" for i in range(1, 13)], width=5, state="readonly")
execution_key_combobox.grid(row=11, column=1, padx=5, pady=5, sticky="w")
execution_key_combobox.bind("<<ComboboxSelected>>", update_execution_key)

ttk.Label(root, text="延遲執行").grid(row=12, column=0, padx=5, pady=5, sticky="w")
delay_var = ttk.StringVar(value="0")
delay_entry = ttk.Entry(root, textvariable=delay_var, width=4)
delay_entry.grid(row=12, column=1, padx=5, pady=5, sticky="w")

ttk.Label(root, text="關閉間隔").grid(row=13, column=0, padx=5, pady=5, sticky="w")
interval_var = ttk.StringVar(value="0")
interval_entry = ttk.Entry(root, textvariable=interval_var, width=4)
interval_entry.grid(row=13, column=1, padx=5, pady=5, sticky="w")

# 常駐模式選項區塊
style.configure("Persistent.TLabel", foreground="#FFFFFF", font=("微軟正黑體", 12, "bold"))
style.configure("PersistentGreen.TLabel", foreground="#198754", font=("微軟正黑體", 14, "bold"))
style.configure("PersistentGreen.TCheckbutton", foreground="#198754", font=("微軟正黑體", 14, "bold"))

persistent_frame = ttk.LabelFrame(root, borderwidth=2)
persistent_frame.grid(row=15, column=0, columnspan=2, padx=5, pady=5, sticky="we")

auto_run_var = ttk.IntVar(value=0)
auto_run_checkbutton = ttk.Checkbutton(
    persistent_frame, text="程式啟動時自動執行", variable=auto_run_var, command=save_config, style="PersistentGreen.TCheckbutton"
)
auto_run_checkbutton.grid(row=0, column=0, columnspan=2, pady=5, sticky="w")

# 自動退出此工具（輸入框在左，標籤在右，含懸浮提示）
auto_close_var = ttk.StringVar(value="")  
auto_close_entry = ttk.Entry(persistent_frame, textvariable=auto_close_var, width=4)
auto_close_entry.grid(row=1, column=0, padx=5, pady=5, sticky="w")
auto_close_label = ttk.Label(persistent_frame, text="自動退出此工具", style="PersistentGreen.TLabel")
auto_close_label.grid(row=1, column=1, padx=5, pady=5, sticky="w")
def show_auto_close_tip(event):
    x = event.x_root + 10
    y = event.y_root + 10
    global tip_window_auto_close
    tip_window_auto_close = ttk.Toplevel(root)
    tip_window_auto_close.wm_overrideredirect(True)
    tip_window_auto_close.geometry(f"+{x}+{y}")
    label = ttk.Label(tip_window_auto_close, text="可輸入0~60秒\n沒數字=不自動退出\n0=馬上退出", background="#FFFACD", foreground="#000")
    label.pack()
def hide_auto_close_tip(event):
    global tip_window_auto_close
    if tip_window_auto_close:
        tip_window_auto_close.destroy()
        tip_window_auto_close = None
tip_window_auto_close = None
auto_close_label.bind("<Enter>", show_auto_close_tip)
auto_close_label.bind("<Leave>", hide_auto_close_tip)

# 重複執行次數（輸入框在左，標籤在右，含懸浮提示）
repeat_var = ttk.StringVar(value="1")
repeat_entry = ttk.Entry(persistent_frame, textvariable=repeat_var, width=4)
repeat_entry.grid(row=2, column=0, padx=5, pady=5, sticky="w")
repeat_label = ttk.Label(persistent_frame, text="重複執行次數", style="PersistentGreen.TLabel")
repeat_label.grid(row=2, column=1, padx=5, pady=5, sticky="w")
def show_tip(event):
    x = event.x_root + 10
    y = event.y_root + 10
    global tip_window
    tip_window = ttk.Toplevel(root)
    tip_window.wm_overrideredirect(True)
    tip_window.geometry(f"+{x}+{y}")
    label = ttk.Label(tip_window, text="輸入 0 會無線重複", background="#FFFACD", foreground="#000")
    label.pack()
def hide_tip(event):
    global tip_window
    if tip_window:
        tip_window.destroy()
        tip_window = None
tip_window = None
repeat_label.bind("<Enter>", show_tip)
repeat_label.bind("<Leave>", hide_tip)

# 直接接續儲存與執行按鈕
def on_save():
    save_config()

def on_execute():
    generate_and_run_ahk()

button_frame = ttk.Frame(root)
button_frame.grid(row=100, column=0, columnspan=2, pady=10)
save_button = ttk.Button(button_frame, text="儲存", style="info.TButton", command=on_save, width=10)
save_button.pack(side="left", padx=10)
execute_button = ttk.Button(button_frame, text="執行", style="success.TButton", command=on_execute, width=10)
execute_button.pack(side="left", padx=10)

load_config()
root.mainloop()

