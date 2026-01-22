#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
YAML檔案 color-line 搜尋工具
搜尋所有.yaml檔案中的 color-line 資訊並去重顯示
"""

import os
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
from pathlib import Path
import pandas as pd
import getpass
from datetime import datetime
import sys
import re
import yaml

# 添加 lib 目錄到路徑（從專案根目錄）
sys.path.insert(0, str(Path(__file__).parent.parent.parent / 'lib'))
import mainlib


class ColorLineSearchGUI:
    
    def load_ignore_list(self):
        """讀取 .ignore 檔案中的忽略目錄列表"""
        ignore_file = Path('.ignore')
        ignored_dirs = set()
        
        if ignore_file.exists():
            try:
                with open(ignore_file, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        # 忽略空行和註解行（以 # 開頭）
                        if line and not line.startswith('#'):
                            # 移除結尾的 / 或 \ 符號
                            dir_name = line.rstrip('/\\')
                            if dir_name:
                                ignored_dirs.add(dir_name)
                
                if ignored_dirs:
                    print(f"已載入 .ignore 檔案，忽略以下目錄：{', '.join(ignored_dirs)}")
            except Exception as e:
                print(f"讀取 .ignore 檔案時發生錯誤：{e}")
        
        return ignored_dirs
    
    def __init__(self, root):
        self.root = root
        self.root.title("YAML Color-Line 搜尋工具")
        self.root.geometry("1200x800")
        
        # 添加結果變數
        self.df_result = None
        self.username = getpass.getuser()  # 獲取登入使用者名稱
        
        # 讀取忽略目錄列表
        self.ignored_dirs = self.load_ignore_list()
        
        # 儲存選擇的搜尋目錄
        self.selected_dir = None
        
        # 儲存找到的 color-line 資料
        self.color_line_data = []

        self.setup_ui()
        self.is_running = False

    def setup_ui(self):
        """設置UI介面"""
        # 獲取螢幕尺寸
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 配置grid權重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)

        # 標題
        title_label = ttk.Label(main_frame, text="YAML Color-Line 搜尋工具",
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 5))

        # 按鈕框架
        button_frame = tk.Frame(self.root, borderwidth=1, relief="flat")
        button_frame.place(x=0, y=40, width=screen_width, height=60)

        # 執行按鈕
        self.run_button = tk.Button(button_frame, text="執行", command=self.start_search, font=(12), width=8, height=2)
        self.run_button.place(x=10, y=10)

        # Excel存檔按鈕
        self.save_button = tk.Button(button_frame, text="Excel存檔", command=self.save_excel, font=(12), width=10, height=2, state="disabled")
        self.save_button.place(x=110, y=10)

        # 發送E-mail按鈕
        self.mail_button = tk.Button(button_frame, text="發送E-mail", command=self.send_email, font=(12), width=10, height=2, state="disabled")
        self.mail_button.place(x=240, y=10)

        # 離開按鈕
        exit_button = tk.Button(button_frame, text="離開", command=self.exit_program, font=(12), width=8, height=2)
        exit_button.place(x=370, y=10)

        # 進度條
        self.progress_var = tk.StringVar(value="準備就緒")
        self.progress_label = ttk.Label(main_frame, textvariable=self.progress_var)
        self.progress_label.grid(row=2, column=0, pady=(0, 5), sticky=tk.W)

        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.grid(row=3, column=0, pady=(0, 10), sticky=(tk.W, tk.E))

        # 結果顯示框架
        result_frame = ttk.Frame(main_frame)
        result_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)

        # 滾動文字框
        self.result_text = scrolledtext.ScrolledText(
            result_frame,
            wrap=tk.WORD,
            font=("Consolas", 10),
            bg='white'
        )
        self.result_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 設置標籤顏色
        self.result_text.tag_configure("header", foreground="blue", font=("Consolas", 10, "bold"))
        self.result_text.tag_configure("data_odd", background="#f0f0f0")
        self.result_text.tag_configure("data_even", background="#ffffff")
        self.result_text.tag_configure("color_line", foreground="red", font=("Consolas", 10, "bold"))

    def count_total_yaml_files(self, start_dir):
        """統計總共有多少個 .yaml 檔案"""
        count = 0
        for root, dirs, files in os.walk(start_dir):
            # 過濾忽略的目錄
            dirs[:] = [d for d in dirs if d not in self.ignored_dirs and not d.startswith('.') and d != '__pycache__']
            
            # 檢查路徑是否包含忽略目錄
            root_path = Path(root)
            should_skip = False
            for ignored in self.ignored_dirs:
                if ignored in root_path.parts:
                    should_skip = True
                    break
            
            if should_skip:
                continue
            
            for file in files:
                if file.lower().endswith('.yaml'):
                    count += 1
        
        return count

    def extract_color_line_from_yaml(self, yaml_file):
        """從 YAML 檔案中提取 color-line 資訊"""
        try:
            with open(yaml_file, 'r', encoding='utf-8') as f:
                content = f.read()
                
            # 使用正則表達式搜尋 "color-line: XXXX"
            # 匹配模式: "color-line: 任意字符直到遇到 "
            pattern = r'"color-line:\s*([^"]+)"'
            matches = re.findall(pattern, content)
            
            if matches:
                return matches[0].strip()  # 回傳第一個匹配結果
            
            return None
            
        except Exception as e:
            print(f"讀取檔案 {yaml_file} 時發生錯誤：{e}")
            return None

    def search_color_lines(self, start_dir=None):
        """搜尋所有 YAML 檔案中的 color-line"""
        self.progress_var.set("正在統計 .yaml 檔案總數...")
        self.result_text.delete(1.0, tk.END)
        
        # 使用選擇的目錄或當前目錄
        if start_dir is None:
            start_dir = Path('.')
        else:
            start_dir = Path(start_dir)
        
        # 顯示搜尋目錄
        self.result_text.insert(tk.END, f"📁 搜尋目錄：{start_dir.resolve()}\n", "header")
        
        # 顯示已加載的忽略目錄
        if self.ignored_dirs:
            self.result_text.insert(tk.END, f"⚠️ 已忽略以下目錄：{', '.join(sorted(self.ignored_dirs))}\n\n", "header")
        else:
            self.result_text.insert(tk.END, "\n", "header")
        
        # 先統計總 .yaml 檔案數
        total_yaml_files = self.count_total_yaml_files(start_dir)
        self.result_text.insert(tk.END, f"📊 找到 {total_yaml_files} 個 .yaml 檔案\n\n", "header")
        self.root.update()

        # 儲存找到的 color-line 資料
        color_line_dict = {}  # {color_line: [file_paths]}
        unique_colors = set()  # 用於去重
        total_files = 0
        current_count = 0
        
        # 遞歸搜尋所有 .yaml 檔案
        for root, dirs, files in os.walk(start_dir):
            # 過濾忽略的目錄
            dirs[:] = [d for d in dirs if d not in self.ignored_dirs and not d.startswith('.') and d != '__pycache__']
            
            # 檢查路徑是否包含忽略目錄
            root_path = Path(root)
            should_skip = False
            for ignored in self.ignored_dirs:
                if ignored in root_path.parts:
                    should_skip = True
                    break
            
            if should_skip:
                continue
            
            for file in files:
                if file.lower().endswith('.yaml'):
                    total_files += 1
                    current_count += 1
                    yaml_file = Path(root) / file
                    # 顯示帶計數器的進度訊息
                    self.progress_var.set(f"正在處理({current_count}/{total_yaml_files})：{yaml_file.name}...")
                    self.root.update()
                    
                    # 提取 color-line
                    color_line = self.extract_color_line_from_yaml(yaml_file)
                    
                    if color_line:
                        unique_colors.add(color_line)
                        if color_line not in color_line_dict:
                            color_line_dict[color_line] = []
                        color_line_dict[color_line].append(str(yaml_file))
        
        return color_line_dict, total_files, len(unique_colors)

    def display_results(self, color_line_dict, total_files, unique_count):
        """顯示搜尋結果並創建 DataFrame"""
        # 設定加大字體
        font_large = ("Consolas", 12)
        font_header = ("Consolas", 12, "bold")
        self.result_text.configure(font=font_large)
        
        # 設定標籤樣式
        self.result_text.tag_configure("header", font=font_header, foreground="blue")
        self.result_text.tag_configure("color_line", font=("Consolas", 12, "bold"), foreground="red")
        self.result_text.tag_configure("summary", background="#FFFF00", font=("Consolas", 12, "bold"))

        # 顯示統計摘要
        self.result_text.insert(tk.END, f"\n{'='*80}\n", "summary")
        self.result_text.insert(tk.END, f"搜尋完成統計\n", "summary")
        self.result_text.insert(tk.END, f"{'='*80}\n", "summary")
        self.result_text.insert(tk.END, f"總共掃描檔案數：{total_files} 個\n", "summary")
        self.result_text.insert(tk.END, f"找到 color-line 的唯一顏色數：{unique_count} 個\n", "summary")
        self.result_text.insert(tk.END, f"{'='*80}\n\n", "summary")

        # 創建 DataFrame 用於 Excel 匯出
        data_rows = []
        
        if color_line_dict:
            self.result_text.insert(tk.END, "找到的 Color-Line 列表（依字母排序）：\n\n", "header")
            
            # 按照 color-line 名稱排序
            sorted_colors = sorted(color_line_dict.keys())
            
            for i, color_line in enumerate(sorted_colors, 1):
                file_count = len(color_line_dict[color_line])
                
                # 顯示在文字框中
                tag = "data_even" if i % 2 == 0 else "data_odd"
                self.result_text.insert(tk.END, f"{i:3d}. ", tag)
                self.result_text.insert(tk.END, f"{color_line}", "color_line")
                self.result_text.insert(tk.END, f" (出現於 {file_count} 個檔案)\n", tag)
                
                # 收集資料到 DataFrame
                # 只顯示第一個檔案路徑作為範例
                example_file = color_line_dict[color_line][0]
                data_rows.append({
                    '序號': i,
                    'Color-Line': color_line,
                    '出現次數': file_count,
                    '範例檔案': example_file
                })
        else:
            self.result_text.insert(tk.END, "未找到任何 color-line 資訊\n", "header")

        self.result_text.see(tk.END)
        
        # 創建 DataFrame
        if data_rows:
            self.df_result = pd.DataFrame(data_rows)
            # 啟用按鈕
            self.save_button.config(state="normal")
            self.mail_button.config(state="normal")
        else:
            self.df_result = None
            self.save_button.config(state="disabled")
            self.mail_button.config(state="disabled")

    def start_search(self):
        """開始搜尋"""
        if self.is_running:
            messagebox.showwarning("警告", "搜尋正在進行中，請稍候...")
            return
        
        # 彈出目錄選擇對話框
        selected_dir = filedialog.askdirectory(
            title="選擇要搜尋的目錄",
            initialdir=os.getcwd()  # 默認為當前目錄
        )
        
        # 如果使用者取消選擇，則不執行
        if not selected_dir:
            self.progress_var.set("已取消搜尋")
            return
        
        self.selected_dir = selected_dir

        self.is_running = True
        self.run_button.config(state='disabled')
        self.progress_bar.start()

        # 在背景執行緒中進行搜尋
        def run_search():
            try:
                color_line_dict, total_files, unique_count = self.search_color_lines(self.selected_dir)
                self.root.after(0, lambda: self.display_results(color_line_dict, total_files, unique_count))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"搜尋過程中發生錯誤：{str(e)}"))
            finally:
                self.root.after(0, self.finish_search)

        thread = threading.Thread(target=run_search, daemon=True)
        thread.start()

    def finish_search(self):
        """完成搜尋"""
        self.is_running = False
        self.run_button.config(state='normal')
        self.progress_bar.stop()
        
        # 根據是否有結果決定按鈕狀態
        if self.df_result is not None and len(self.df_result) > 0:
            self.save_button.config(state='normal')
            self.mail_button.config(state='normal')
            self.progress_var.set(f"搜尋完成（找到 {len(self.df_result)} 個不同的 color-line）")
        else:
            self.save_button.config(state='disabled')
            self.mail_button.config(state='disabled')
            self.progress_var.set("搜尋完成（無資料）")

    def save_excel(self):
        """將搜尋結果儲存為 Excel 檔案"""
        if self.df_result is None or len(self.df_result) == 0:
            messagebox.showwarning("無資料", "目前沒有搜尋結果可以儲存！")
            return
        
        # 設定預設資料夾為 output 目錄
        output_dir = Path(__file__).parent.parent.parent / 'output'
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # 預設檔名為 yaml020.xlsx
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="儲存 Excel 檔案",
            initialdir=str(output_dir),
            initialfile="yaml020.xlsx"
        )
        
        if file_path:
            try:
                # 先儲存基本的 Excel 檔案
                self.df_result.to_excel(file_path, index=False, engine='openpyxl')
                # 格式化 Excel 檔案（使用 mainlib）
                mainlib.format_excel_file(file_path)
                self.progress_var.set(f"Excel 檔案已儲存至：{file_path}")
                messagebox.showinfo("成功", f"檔案已儲存至：\n{file_path}")
            except Exception as e:
                messagebox.showerror("錯誤", f"儲存檔案時發生錯誤：\n{e}")

    def send_email(self):
        """調用 mainlib 發送郵件"""
        if self.df_result is None or len(self.df_result) == 0:
            messagebox.showwarning("無資料", "目前沒有搜尋結果可以寄送！")
            return
        
        # 定義進度更新回調函數
        def update_progress(message):
            self.progress_var.set(message)
            self.root.update()
        
        try:
            # 調用 mainlib 的郵件發送功能
            success, message = mainlib.send_statistics_email(
                df_result=self.df_result,
                parent_window=self.root,
                progress_callback=update_progress
            )
            
            # 更新進度訊息
            self.progress_var.set(message)
            
            if not success:
                messagebox.showerror("郵件發送失敗", f"發生錯誤：{message}")
        except Exception as e:
            error_msg = f"發生錯誤：{str(e)}"
            self.progress_var.set("郵件發送失敗")
            messagebox.showerror("郵件發送失敗", f"發生錯誤：{str(e)}")

    def exit_program(self):
        """離開程式"""
        if self.is_running:
            if messagebox.askyesno("確認", "搜尋正在進行中，確定要離開嗎？"):
                self.root.quit()
        else:
            self.root.quit()


def main():
    root = tk.Tk()
    app = ColorLineSearchGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()