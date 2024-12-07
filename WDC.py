import os
import ttkbootstrap as ttk  # 使用 ttkbootstrap 來創建窗口樣式
from tkinter import messagebox  # 用來顯示訊息框

try:
    import keyboard  # 嘗試導入 keyboard 庫
except ImportError:
    keyboard = None
    print("Error: 'keyboard' 庫未安裝。請使用 'pip install keyboard' 進行安裝。")

# 創建 Tkinter 視窗，使用 ttkbootstrap 提供的樣式
root = ttk.Window(themename="superhero")  # 使用 ttkbootstrap 主題

# 設置視窗的背景色
root.tk_setPalette(background="#172B4B")

# 設置窗口標題
root.title("WDC.exe")

# 設置本地圖示 (假設圖示檔案在專案根目錄)
icon_path = "./Nekoneko.ico"
if os.path.exists(icon_path):
    root.iconbitmap(icon_path)  # 設置本地圖示
else:
    print("圖示檔案不存在，未設置圖示。")

# 設置字體和顏色
root.option_add("*Font", "微軟正黑體 14 bold")  # 全局字體設置為粗體
root.option_add("*foreground", "#F5D07D")  # 設置字體顏色

# 全局變數儲存執行按鍵設定
execution_key = "F1"  # 預設執行按鍵

# AHK 腳本生成和執行邏輯
def generate_and_run_ahk():
    try:
        # 獲取用戶輸入
        num_windows = int(entry_num_windows.get())
        window_titles = [entry_windows[i].get() for i in range(num_windows) if entry_windows[i].get().strip()]

        if not window_titles:
            messagebox.showerror("Error", "請輸入至少一個視窗名稱！")
            return

        # 動態生成 AHK 腳本
        ahk_script = """#NoTrayIcon  ; 隱藏托盤圖示

; 嘗試執行窗口關閉操作，捕捉任何錯誤並忽略
try {
"""
        for title in window_titles:
            ahk_script += f'    WinClose("{title}")\n'
            ahk_script += "    Sleep(500)  ; 等待 0.5 秒，執行下一輪關閉操作\n"

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
        messagebox.showerror("Error", "請輸入有效的數字！")

# 更新執行按鍵的熱鍵
def update_execution_key(*args):
    global execution_key
    execution_key = execution_key_var.get()
    if keyboard:
        keyboard.clear_all_hotkeys()  # 清除現有的熱鍵
        keyboard.add_hotkey(execution_key, generate_and_run_ahk)
        print(f"全局熱鍵已設定為：{execution_key}")
    else:
        print("Error: 無法設定全局熱鍵，因為 'keyboard' 庫未安裝。")

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
    except ValueError:
        pass

# 創建 UI 元素
ttk.Label(root, text="關閉視窗數").grid(row=0, column=0, padx=10, pady=10, sticky="w")
entry_num_windows = ttk.Combobox(root, values=[str(i) for i in range(1, 11)], width=5, state="readonly")
entry_num_windows.grid(row=0, column=1, padx=5, pady=10)
entry_num_windows.set("1")
entry_num_windows.bind("<<ComboboxSelected>>", update_window_inputs)

ttk.Label(root, text="視窗名稱").grid(row=1, column=0, padx=10, pady=10, sticky="w")
entry_windows = []
for i in range(10):
    entry = ttk.Entry(root, width=15)
    entry.grid(row=1 + i, column=1, columnspan=2, padx=10, pady=5)
    entry_windows.append(entry)
    if i > 0:
        entry.grid_remove()
entry_windows[0].insert(0, "MuMu操作錄製")

# 設置執行按鍵
ttk.Label(root, text="設定執行按鍵").grid(row=13, column=0, padx=10, pady=10, sticky="w")
execution_key_var = ttk.StringVar(value=execution_key)
execution_key_combobox = ttk.Combobox(root, textvariable=execution_key_var, values=[f"F{i}" for i in range(1, 13)], width=5, state="readonly")
execution_key_combobox.grid(row=13, column=1, padx=10, pady=10)
execution_key_combobox.bind("<<ComboboxSelected>>", update_execution_key)

# 執行按鈕
execute_button = ttk.Button(root, text="執行", command=generate_and_run_ahk)
execute_button.grid(row=14, column=0, columnspan=2, pady=20)
execute_button.config(width=20, style="success.TButton")  # 加大按鈕寬度並設置樣式

# 啟動全局熱鍵監聽
if keyboard:
    keyboard.add_hotkey(execution_key, generate_and_run_ahk)
else:
    print("警告：無法啟用全局熱鍵，請安裝 'keyboard' 庫。")

# 啟動主循環
root.mainloop()
