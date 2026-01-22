import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
import subprocess
from PIL import Image, ImageTk
from pathlib import Path
import sys

# 添加 lib 目錄到路徑
sys.path.insert(0, str(Path(__file__).parent))
from lib import get_project_root, get_xlsx_path, get_image_path
"""
sys.exit()  # 在正式的程式中，建議使用 sys.exit()，因為它更明確且可控制退出碼。
sys.exit(0) # 表示正常結束
sys.exit(1) # 表示錯誤結束
exit() # 在正式的程式中，建議使用 sys.exit()，因為它更明確且可控制退出碼。
quit()

root.destroy() # 在 GUI 程式中（例如 tkinter）  
# 這會關閉視窗並結束 GUI 程式。
# root.quit() 也可以用來結束 mainloop()，但 destroy() 是更徹底的關閉方式。
"""
# 取得專案根目錄
base_path = get_project_root()

# 讀取 Excel（使用標準路徑函數）
menu_path = get_xlsx_path("menu.xlsx")

df = pd.read_excel(menu_path, engine='openpyxl')

# 建立選單字典
menu_dict = {}
for _, row in df.iterrows():
    system = row['System']
    code = row['Code']
    desc = row['Description']
    menu_dict.setdefault(system, []).append((code, desc))

# 執行程式

def execute_script(event):
    selected_item = tree.focus()
    item_text = tree.item(selected_item, 'text')
    if ' ' in item_text:
        code = item_text.split()[0]
        # 將 code 轉換為小寫
        code_lower = code.lower()
        
        # 智慧搜尋腳本位置
        # 1. 根據檔名前綴判斷系統類別（如 yaml015 → yaml, sql010 → sql, game005-guess → game）
        # 提取開頭的連續字母部分作為前綴
        prefix = ""
        for c in code:
            if c.isalpha():
                prefix += c
            else:
                break  # 遇到非字母就停止
        prefix = prefix.lower()
        
        # 2. 定義搜尋路徑（優先搜尋 .exe，找不到才找 .py）
        exe_name = f"{code_lower}.exe"
        py_name = f"{code_lower}.py"
        
        search_paths = [
            # 優先找 .exe（部署模式）
            ("py", base_path / "modules" / prefix / py_name),          # modules/[prefix]/xxx.py
            ("py", base_path / py_name),                               # 根目錄（向後兼容）
            ("exe", base_path / "dist" / exe_name),                    # dist/xxx.exe
            # 找不到才找 .py（開發模式）
        ]
        
        # 3. 搜尋並執行
        found = False
        for file_type, path in search_paths:
            if path.exists():
                found = True
                try:
                    if file_type == "exe":
                        # 執行 .exe（獨立運行，不需要 Python）
                        subprocess.Popen([str(path)])
                    else:
                        # 執行 .py（需要 Python 環境）
                        subprocess.run(['python', path], check=True)
                    break
                except Exception as e:
                    messagebox.showerror("執行錯誤", f"執行程式時發生錯誤：{e}")
                    break
        
        if not found:
            # 顯示錯誤訊息
            search_list = "\n".join([f"- {p}" for _, p in search_paths])
            messagebox.showinfo("錯誤", f"找不到程式：{code_lower}\n\n搜尋位置：\n{search_list}")
# 建立 GUI
root = tk.Tk()
root.title("建大ERP系統 FOR python")

# 讓視窗一開始就最大化
root.state('zoomed')

# 使用 Grid 佈局來管理左右兩部分
root.grid_columnconfigure(0, weight=0, minsize=318) # 左邊 TreeView，設置默認寬度318，且不能隨意拉伸


root.grid_columnconfigure(1, weight=1)              # 右邊 Frame，可以隨視窗大小拉伸

root.grid_rowconfigure(0, weight=1)


# 左邊的 TreeView Frame (可調整大小)
# 放置 Frame 的方式 你可以用三種方式放置 Frame：

# pack()：自動排列（上下或左右）
# grid()：格狀排列（像表格）
# place()：精確控制位置（x, y）

tree_frame = ttk.Frame(root)
tree_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=10)
# row=0, column=0  沒有準確的值, row(上下) column (左右) 各約到 >=100 就是最大的
# 在技術上，padx 和 pady 沒有「最大值」的限制，它們可以是任意的整數，只要不超出視窗的大小或造成排版問題。你可以設定為：
# 建議使用方式
# 小間距：padx=5, pady=5（常見於緊湊排版）
# 中間距：padx=10~20, pady=10~20（常見於一般 UI）
# 大間距：padx=30+, pady=30+（用於分隔區塊或強調）
# 如果你使用 sticky="nsew"，表示元件會延伸填滿格子，padding 就是額外的空間。
# 若你發現元件「看起來沒反應」，可能是因為父容器沒有設定 grid_rowconfigure 或 grid_columnconfigure 的 weight。

# 建立 TreeView
tree = ttk.Treeview(tree_frame)
tree.place(x=0, y=0, width=300, height=840)
# 如果你使用 .place()，就不需要 .pack() 或 .grid()，避免混用。
# tree.pack(fill=tk.BOTH, expand=True)

# 載入資料到 TreeView
for system, items in menu_dict.items():
    parent = tree.insert('', 'end', text=system, open=True)
    for code, desc in items:
        tree.insert(parent, 'end', text=f"{code} {desc}")

# 綁定雙擊事件
tree.bind('<Double-1>', execute_script)

# 右邊的資訊 Frame
info_frame = ttk.Frame(root)
info_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)

# 在 info_frame 中放入圖片、logo 和文字 (請確保 'logo.png' 存在於相同的目錄中)
try:
    
    # 載入圖片並調整大小（使用標準路徑函數）
    logo_path = get_image_path("logo----.png")
    original_image = Image.open(logo_path)
    # 將圖片調整為適合在 Frame 中顯示的大小
    resized_image = original_image.resize((300, 300), Image.LANCZOS)
    
    logo_img = ImageTk.PhotoImage(resized_image)
    
    # 建立一個 Label 來顯示圖片
    logo_label = ttk.Label(info_frame, image=logo_img)
    logo_label.pack(pady=20)
    # 保持對圖片的引用，防止被垃圾回收
    logo_label.image = logo_img
    
except FileNotFoundError:
    logo_label = ttk.Label(info_frame, text="找不到 logo.png 檔案", foreground="red")
    logo_label.pack(pady=20)

# 範例：在 info_frame 中加入文字
# info_label = ttk.Label(info_frame, text="這是一個示範應用程式\n\n歡迎使用", font=("Helvetica", 16))
# info_label.pack(pady=10)

# 啟動 GUI
# 左下角新增一個 100x100 的 Frame
bottom_left_frame = ttk.Frame(root, width=320, height=45)
bottom_left_frame.place(relx=0.0, rely=1.0, anchor='sw')  # 固定在左下角
bottom_left_frame.pack_propagate(False)  # 防止自動調整大小

# 顯示文字的按鈕
def show_text():
    label = ttk.Label(bottom_left_frame, text="")
    label.pack()

# 控制按鈕的位置與大小
# show_x = 10         # x 座標
# show_y = 10         # y 座標
# show_width = 80     # 寬度
# show_height = 30    # 高度

# 建立按鈕並使用 place 放置
show_button = ttk.Button(bottom_left_frame, text="備用文字", command=show_text)
show_button.place(x=5, y=8, width=70, height=30)
show_button.state(['disabled'])  # 設定為停用狀態

# show_button = ttk.Button(bottom_left_frame, text="顯示文字", command=show_text)
# show_button.pack(pady=5)
show_button = ttk.Button(bottom_left_frame, text="備用數字", command=show_text)
show_button.place(x=80, y=8, width=70, height=30)
show_button.state(['disabled'])  # 設定為停用狀態

show_button = ttk.Button(bottom_left_frame, text="顯示文字", command=show_text)
show_button.place(x=155, y=8, width=70, height=30)
# 離開程式的按鈕
exit_button = ttk.Button(bottom_left_frame, text="離開", command=lambda: sys.exit(0))
exit_button.place(x=230, y=8, width=70, height=30)
# exit_button.pack(pady=5)

root.mainloop()