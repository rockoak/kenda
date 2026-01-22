#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SQL010 - Excel 檔案郵件發送工具
讀取 Excel 檔案並透過 E-mail 發送
"""

import sys
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

# 添加 lib 目錄到路徑（從專案根目錄）
sys.path.insert(0, str(Path(__file__).parent.parent.parent / 'lib'))
import mainlib


class SQL010GUI:
    def __init__(self, root):
        self.root = root
        self.root.title("SQL010 - Excel 郵件發送工具")
        self.root.geometry("600x400")
        
        # 存儲讀取的資料
        self.df_result = None
        self.file_path = None
        
        self.setup_ui()
    
    def setup_ui(self):
        """設置UI介面"""
        # 主框架
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 標題
        title_label = tk.Label(
            main_frame, 
            text="Excel 郵件發送工具",
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # 說明文字
        info_label = tk.Label(
            main_frame,
            text="請選擇 Excel 檔案，然後發送郵件",
            font=("Arial", 11)
        )
        info_label.pack(pady=(0, 30))
        
        # 檔案路徑顯示框
        path_frame = tk.Frame(main_frame)
        path_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(path_frame, text="選擇的檔案：", font=("Arial", 10)).pack(anchor=tk.W)
        
        self.path_var = tk.StringVar(value="尚未選擇檔案")
        path_label = tk.Label(
            path_frame, 
            textvariable=self.path_var,
            font=("Arial", 9),
            fg="blue",
            wraplength=550,
            justify=tk.LEFT
        )
        path_label.pack(anchor=tk.W, pady=(5, 0))
        
        # 資料預覽框
        preview_frame = tk.Frame(main_frame)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        tk.Label(preview_frame, text="檔案資訊：", font=("Arial", 10)).pack(anchor=tk.W)
        
        self.info_text = tk.Text(
            preview_frame,
            height=8,
            width=70,
            font=("Consolas", 9),
            wrap=tk.WORD,
            state=tk.DISABLED
        )
        self.info_text.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        
        # 按鈕框架
        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=(10, 0))
        
        # 讀取按鈕
        self.read_button = tk.Button(
            button_frame,
            text="讀取",
            command=self.read_file,
            font=("Arial", 12),
            width=10,
            height=2,
            bg="#4CAF50",
            fg="white",
            cursor="hand2"
        )
        self.read_button.pack(side=tk.LEFT, padx=10)
        
        # 發送郵件按鈕
        self.mail_button = tk.Button(
            button_frame,
            text="發送E-mail",
            command=self.send_email,
            font=("Arial", 12),
            width=12,
            height=2,
            bg="#2196F3",
            fg="white",
            cursor="hand2",
            state=tk.DISABLED
        )
        self.mail_button.pack(side=tk.LEFT, padx=10)
        
        # 離開按鈕
        exit_button = tk.Button(
            button_frame,
            text="離開",
            command=self.exit_program,
            font=("Arial", 12),
            width=10,
            height=2,
            bg="#f44336",
            fg="white",
            cursor="hand2"
        )
        exit_button.pack(side=tk.LEFT, padx=10)
        
        # 狀態列
        self.status_var = tk.StringVar(value="準備就緒")
        status_label = tk.Label(
            self.root,
            textvariable=self.status_var,
            font=("Arial", 9),
            relief=tk.SUNKEN,
            anchor=tk.W,
            padx=5
        )
        status_label.pack(side=tk.BOTTOM, fill=tk.X)
    
    def read_file(self):
        """讀取 Excel 檔案"""
        # 開啟檔案選擇對話框
        file_path = filedialog.askopenfilename(
            title="選擇 Excel 檔案",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("Excel files", "*.xls"),
                ("All files", "*.*")
            ],
            initialdir=mainlib.get_project_root() / "xlsx"  # 預設開啟 xlsx 目錄
        )
        
        # 如果使用者取消選擇
        if not file_path:
            self.status_var.set("取消讀取檔案")
            return
        
        try:
            # 讀取 Excel 檔案
            self.df_result = pd.read_excel(file_path)
            self.file_path = file_path
            
            # 更新顯示
            self.path_var.set(file_path)
            
            # 顯示檔案資訊
            self.info_text.config(state=tk.NORMAL)
            self.info_text.delete(1.0, tk.END)
            
            info = f"檔案名稱：{Path(file_path).name}\n"
            info += f"資料行數：{len(self.df_result)} 行\n"
            info += f"資料欄數：{len(self.df_result.columns)} 欄\n\n"
            info += f"欄位名稱：\n"
            for i, col in enumerate(self.df_result.columns, 1):
                info += f"  {i}. {col}\n"
            
            self.info_text.insert(1.0, info)
            self.info_text.config(state=tk.DISABLED)
            
            # 啟用發送郵件按鈕
            self.mail_button.config(state=tk.NORMAL)
            
            self.status_var.set(f"成功讀取檔案：{Path(file_path).name}")
            messagebox.showinfo("成功", f"已成功讀取檔案！\n\n資料行數：{len(self.df_result)} 行\n資料欄數：{len(self.df_result.columns)} 欄")
            
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取檔案時發生錯誤：\n{str(e)}")
            self.status_var.set("讀取檔案失敗")
            self.df_result = None
            self.file_path = None
            self.mail_button.config(state=tk.DISABLED)
    
    def send_email(self):
        """發送郵件（調用 mainlib）"""
        if self.df_result is None:
            messagebox.showwarning("無資料", "請先讀取 Excel 檔案！")
            return
        
        # 定義進度更新回調函數
        def update_progress(message):
            self.status_var.set(message)
            self.root.update()
        
        try:
            # 調用 mainlib 的郵件發送功能
            success, message = mainlib.send_statistics_email(
                df_result=self.df_result,
                parent_window=self.root,
                progress_callback=update_progress
            )
            
            # 更新狀態訊息
            self.status_var.set(message)
            
        except Exception as e:
            messagebox.showerror("錯誤", f"發送郵件時發生錯誤：\n{str(e)}")
            self.status_var.set("郵件發送失敗")
    
    def exit_program(self):
        """離開程式"""
        self.root.quit()


def main():
    root = tk.Tk()
    app = SQL010GUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

