import os
import json
from pathlib import Path
import ttkbootstrap as ttk  # 使用 ttkbootstrap 來創建窗口樣式

try:
    import keyboard  # 嘗試導入 keyboard 庫
except ImportError:
    keyboard = None
    print("Error: 'keyboard' 庫未安裝。請使用 'pip install keyboard' 進行安裝。")

# 配置檔案名稱
CONFIG_FILE = "config.json"

# 創建 Tkinter 視窗，使用 ttkbootstrap 提供的樣式
root = ttk.Window(themename="superhero")  # 使用 ttkbootstrap 主題

# 設置視窗的背景色
root.tk_setPalette(background="#172B4B")

# 設置窗口標題和圖標
root.title("ChroLens_Clear.exe")

# 嘗試設置應用程式的圖示
icon_path_a = "./Nekoneko.ico"  # 相對路徑
icon_path_b = r"C:/Users/Lucien/Documents/GitHub/ChroLens_Clear/Nekoneko.ico"  # 絕對路徑

if os.path.exists(icon_path_a):
    root.iconbitmap(icon_path_a)  # 設置相對路徑的圖示
elif os.path.exists(icon_path_b):
    root.iconbitmap(icon_path_b)  # 設置絕對路徑的圖示
else:
    print("圖示檔案不存在，將不設置圖示。")

# 設置字體和顏色
style = ttk.Style()
style.configure("TLabel", foreground="#F5D07D", font=("微軟正黑體", 14, "bold"))
style.configure("TButton", foreground="#F5D07D", font=("微軟正黑體", 14, "bold"))

# 全局變數儲存執行按鍵設定
execution_key = "F8"  # 預設執行按鍵

# 保存用戶設定
def save_config():
    config = {
        "num_windows": entry_num_windows.get(),
        "window_titles": [entry.get() for entry in entry_windows],
        "execution_key": execution_key_var.get(),
        "auto_run": auto_run_var.get(),  # 是否啟用開啟後自動執行
        "auto_close": auto_close_var.get(),  # 是否啟用3秒後自動關閉
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

# 加載用戶設定
def load_config():
    if not Path(CONFIG_FILE).exists():
        return

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        config = json.load(f)

    # 還原設定
    entry_num_windows.set(config.get("num_windows", "1"))
    execution_key_var.set(config.get("execution_key", "F8"))
    auto_run_var.set(config.get("auto_run", 0))
    auto_close_var.set(config.get("auto_close", 0))  # 加載 "3秒後自動關閉" 設定

    window_titles = config.get("window_titles", [])
    for i, title in enumerate(window_titles):
        if i < len(entry_windows):
            entry_windows[i].delete(0, "end")
            entry_windows[i].insert(0, title)

    # 更新視窗名稱輸入框的數量
    update_window_inputs()

    # 如果啟用了自動執行，則延遲1秒後執行動作
    if auto_run_var.get():
        root.after(1000, generate_and_run_ahk)  # 延遲1秒執行

    # 如果啟用了3秒後自動關閉，則設置3秒後關閉程式
    if auto_close_var.get():
        root.after(3000, root.destroy)

# AHK 腳本生成和執行邏輯
def generate_and_run_ahk():
    try:
        # 獲取用戶輸入
        num_windows = int(entry_num_windows.get())
        window_titles = [entry_windows[i].get() for i in range(num_windows) if entry_windows[i].get().strip()]

        if not window_titles:
            return

        # 動態生成 AHK 腳本
        ahk_script = """#NoTrayIcon  ; 隱藏托盤圖示

; 嘗試執行窗口關閉操作，捕捉任何錯誤並忽略
try {
"""
        for title in window_titles:
            ahk_script += f'    WinClose("{title}")\n'
            ahk_script += "    Sleep(200)  ; 等待 0.2 秒，執行下一輪關閉操作\n"

        ahk_script += """} catch {
    ; 如果出現錯誤，這裡捕捉錯誤並忽略，無視顯示
}
"""

        # 保存 AHK 腳本
        ahk_file = Path("temp_close_windows.ahk")
        ahk_file.write_text(ahk_script, encoding="utf-8")

        # 執行 AHK 腳本
        os.system(f"start /B {ahk_file}")

    except ValueError:
        pass

# 更新執行按鍵的熱鍵
def update_execution_key(*args):
    global execution_key
    execution_key = execution_key_var.get()
    if keyboard:
        keyboard.clear_all_hotkeys()  # 清除現有的熱鍵
        keyboard.add_hotkey(execution_key, generate_and_run_ahk)
    save_config()  # 保存配置

# 動態生成視窗名稱輸入框
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
        save_config()  # 保存配置
    except ValueError:
        pass

# 創建 UI 元素
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
entry_windows[0].insert(0, "MuMu操作錄製")  # 預設第一個輸入框內容

# 設置執行按鍵
ttk.Label(root, text="設定執行按鍵").grid(row=11, column=0, padx=5, pady=5, sticky="w")
execution_key_var = ttk.StringVar(value=execution_key)
execution_key_combobox = ttk.Combobox(root, textvariable=execution_key_var, values=[f"F{i}" for i in range(1, 13)], width=5, state="readonly")
execution_key_combobox.grid(row=11, column=1, padx=5, pady=5, sticky="w")
execution_key_combobox.bind("<<ComboboxSelected>>", update_execution_key)

# 自動執行選項
auto_run_var = ttk.IntVar(value=0)  # 默認為未勾選
auto_run_checkbutton = ttk.Checkbutton(
    root, text="開啟後自動執行", variable=auto_run_var, command=save_config
)
auto_run_checkbutton.grid(row=13, column=0, columnspan=2, pady=5, sticky="w")

# 3秒後自動關閉選項
auto_close_var = ttk.IntVar(value=0)  # 默認為未勾選
auto_close_checkbutton = ttk.Checkbutton(
    root, text="3秒後自動關閉", variable=auto_close_var, command=save_config
)
auto_close_checkbutton.grid(row=14, column=0, columnspan=2, pady=5, sticky="w")

# "存檔"按鈕
save_button = ttk.Button(root, text="存檔", command=save_config, style="info.TButton")
save_button.grid(row=15, column=0, padx=5, pady=5, sticky="e")
save_button.config(width=8)

# 執行按鈕
execute_button = ttk.Button(root, text="執行", command=generate_and_run_ahk, style="success.TButton")
execute_button.grid(row=15, column=1, padx=5, pady=5, sticky="w")
execute_button.config(width=8)

# 加載配置
load_config()

# 運行主循環
root.mainloop()
