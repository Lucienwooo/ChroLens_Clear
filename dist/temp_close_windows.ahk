#NoTrayIcon  ; 隱藏托盤圖示

; 嘗試執行窗口關閉操作，捕捉任何錯誤並忽略
try {
    WinClose("MuMu操作錄製")
    Sleep(200)  ; 等待 0.2 秒，執行下一輪關閉操作
    WinClose("MuMu操作錄製")
    Sleep(200)  ; 等待 0.2 秒，執行下一輪關閉操作
    WinClose("MuMu操作錄製")
    Sleep(200)  ; 等待 0.2 秒，執行下一輪關閉操作
    WinClose("MuMu操作錄製")
    Sleep(200)  ; 等待 0.2 秒，執行下一輪關閉操作
    WinClose("MuMu操作錄製")
    Sleep(200)  ; 等待 0.2 秒，執行下一輪關閉操作
    WinClose("MuMu操作錄製")
    Sleep(200)  ; 等待 0.2 秒，執行下一輪關閉操作
    WinClose("MuMu操作錄製")
    Sleep(200)  ; 等待 0.2 秒，執行下一輪關閉操作
    WinClose("MuMu操作錄製")
    Sleep(200)  ; 等待 0.2 秒，執行下一輪關閉操作
} catch {
    ; 如果出現錯誤，這裡捕捉錯誤並忽略，無視顯示
}
