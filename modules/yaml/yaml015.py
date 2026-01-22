#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
YAML檔案統計GUI程式
統計各廠別的.yaml檔案數量
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

# 添加 lib 目錄到路徑（從專案根目錄）
sys.path.insert(0, str(Path(__file__).parent.parent.parent / 'lib'))
import mainlib


class YAMLStatisticsGUI:
    
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
        self.root.title("YAML檔案統計工具")
        self.root.geometry("1200x800")

        # 廠別列表（按指定順序）
        self.plants = ['KY', 'KU', 'KU1', 'KS', 'KC', 'KT', 'KV', 'KV1', 'KV2', 'KI']
        
        # 添加結果變數
        self.df_result = None
        self.username = getpass.getuser()  # 獲取登入使用者名稱
        
        # 讀取忽略目錄列表
        self.ignored_dirs = self.load_ignore_list()
        
        # 儲存選擇的統計目錄
        self.selected_dir = None

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
        title_label = ttk.Label(main_frame, text="YAML檔案統計工具",
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 5))

        # 按鈕框架
        button_frame = tk.Frame(self.root, borderwidth=1, relief="flat")
        button_frame.place(x=0, y=40, width=screen_width, height=60)

        # 執行按鈕
        self.run_button = tk.Button(button_frame, text="執行", command=self.start_statistics, font=(12), width=8, height=2)
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

        # 滾動文字框（設定 CHAR(10) 靠右，即換行時靠右對齊）
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

    def count_yaml_files(self, directory):
        """統計目錄下所有.yaml檔案數量（遞歸）"""
        count = 0
        try:
            for root, dirs, files in os.walk(directory):
                for file in files:
                    if file.lower().endswith('.yaml'):
                        count += 1
        except (PermissionError, OSError):
            pass
        return count

    def get_plant_statistics(self, base_dir):
        """獲取指定目錄下各廠別的統計資料"""
        stats = {}
        base_path = Path(base_dir)

        if not base_path.exists():
            return stats

        # 檢查各廠別目錄
        for plant in self.plants:
            plant_dir = base_path / plant
            if plant_dir.exists():
                count = self.count_yaml_files(str(plant_dir))
                if count > 0:
                    stats[plant] = count

        # 計算其他檔案（不在指定廠別目錄中的.yaml檔案）
        other_count = 0
        for item in base_path.iterdir():
            if item.is_file() and item.name.lower().endswith('.yaml'):
                other_count += 1
            elif item.is_dir() and item.name not in self.plants:
                other_count += self.count_yaml_files(str(item))

        if other_count > 0:
            stats['其他'] = other_count

        return stats

    def format_statistics_line(self, directory, stats):
        """格式化統計結果行"""
        # 廠別標題
        plant_headers = self.plants + ['其他', '總計']

        # 計算總計
        total = sum(stats.values())

        # 建立廠別數量列表
        counts = []
        for plant in self.plants:
            counts.append(str(stats.get(plant, 0)))
        counts.append(str(stats.get('其他', 0)))
        counts.append(str(total))

        # 使用固定寬度格式化 (每欄12個字符，廠別靠右，筆數靠右並加千位分隔符)
        header_line = "廠別 : " + "".join(f"{header:>12}" for header in plant_headers)
        count_line = "筆數 : " + "".join(f"{int(count):>12,}" for count in counts)

        return header_line, count_line

    def collect_statistics(self, start_dir=None):
        """收集所有統計資料"""
        self.progress_var.set("正在統計檔案...")
        self.result_text.delete(1.0, tk.END)
        
        # 使用選擇的目錄或當前目錄
        if start_dir is None:
            start_dir = Path('.')
        else:
            start_dir = Path(start_dir)
        
        # 儲存基礎目錄供相對路徑計算使用
        self.base_dir = start_dir
        
        # 顯示統計目錄
        self.result_text.insert(tk.END, f"📁 統計目錄：{start_dir.resolve()}\n", "header")
        
        # 顯示已加載的忽略目錄
        if self.ignored_dirs:
            self.result_text.insert(tk.END, f"⚠️ 已忽略以下目錄：{', '.join(sorted(self.ignored_dirs))}\n\n", "header")
        else:
            self.result_text.insert(tk.END, "\n", "header")

        results = []

        # 從選擇的目錄開始搜尋，實作廠別目錄停止規則
        processed_dirs = set()  # 記錄已經處理過的目錄，避免重複統計

        self._collect_from_directory(start_dir, results, processed_dirs)

        return results

    def _collect_from_directory(self, directory, results, processed_dirs, depth=0):
        """遞歸收集統計資料，遇到廠別目錄就停止"""
        if str(directory) in processed_dirs:
            return
        
        # 檢查路徑中是否包含被忽略的目錄
        # 將路徑轉換為字串，檢查是否包含忽略的目錄名稱
        dir_str = str(directory).replace('\\', '/')
        for ignored in self.ignored_dirs:
            # 檢查路徑是否以 "./{ignored}" 開頭，或包含 "/{ignored}/" 
            if dir_str == f'./{ignored}' or dir_str.startswith(f'./{ignored}/') or f'/{ignored}/' in dir_str:
                return

        # 檢查當前目錄是否包含廠別子目錄
        plant_subdirs = []
        for item in directory.iterdir():
            if item.is_dir() and item.name in self.plants:
                plant_subdirs.append(item)

        # 如果找到廠別子目錄，在當前目錄層級統計
        if plant_subdirs:
            self.progress_var.set(f"正在處理 {directory.name or '根目錄'}...")
            stats = self.get_plant_statistics(str(directory))
            if stats:  # 只有有資料時才顯示
                # 計算相對路徑（相對於基礎目錄）
                try:
                    rel_path = directory.relative_to(self.base_dir)
                    if str(rel_path) == '.':
                        display_path = str(self.base_dir.resolve())
                    else:
                        display_path = f"{self.base_dir.resolve()}\\{rel_path}"
                except ValueError:
                    display_path = str(directory.resolve())
                results.append((display_path, stats))

            # 標記已處理，避免重複統計
            processed_dirs.add(str(directory))
            return  # 遇到廠別目錄就停止搜尋這個分支

        # 如果沒有廠別子目錄，繼續遞歸搜尋子目錄
        for item in directory.iterdir():
            if (item.is_dir() and
                not item.name.startswith('.') and
                item.name != '__pycache__' and
                item.name not in self.ignored_dirs and  # 跳過 .ignore 檔案中指定的目錄
                str(item) not in processed_dirs):

                self._collect_from_directory(item, results, processed_dirs, depth + 1)

    def display_results(self, results):
        """顯示統計結果並創建 DataFrame"""
        # 設定加大字體 (使用等寬字體確保對齊)
        font_large = ("Consolas", 12)  # 使用等寬字體，適中大小
        font_header = ("Consolas", 12, "bold")  # 目錄與路徑再加粗
        self.result_text.configure(font=font_large)
        self.result_text.delete(1.0, tk.END)
        # 設定 "header" 標籤為加大加粗字體
        self.result_text.tag_configure("header", font=font_header)
        
        # 設定合計列的標籤樣式（黃色底、加粗）
        self.result_text.tag_configure("total_row", background="#FFFF00", font=("Consolas", 12, "bold"))

        # 創建 DataFrame 用於 Excel 匯出
        data_rows = []
        
        # 初始化各欄位的合計
        column_totals = {plant: 0 for plant in self.plants}
        column_totals['其他'] = 0
        column_totals['總計'] = 0

        for i, (directory, stats) in enumerate(results):
            # 選擇顏色標籤
            tag = "data_even" if i % 2 == 0 else "data_odd"

            # 目錄標題
            self.result_text.insert(tk.END, f"目錄 : {directory}\n", "header")

            # 統計行
            header_line, count_line = self.format_statistics_line(directory, stats)
            self.result_text.insert(tk.END, f"{header_line}\n", tag)
            self.result_text.insert(tk.END, f"{count_line}\n\n", tag)
            
            # 收集資料到 DataFrame
            row_data = {'目錄': directory}
            total = sum(stats.values())
            for plant in self.plants:
                value = stats.get(plant, 0)
                row_data[plant] = value
                column_totals[plant] += value
            row_data['其他'] = stats.get('其他', 0)
            column_totals['其他'] += stats.get('其他', 0)
            row_data['總計'] = total
            column_totals['總計'] += total
            data_rows.append(row_data)

        # 顯示 B到M欄的合計列（如果有資料）
        if data_rows:
            self.result_text.insert(tk.END, "=" * 150 + "\n", "total_row")
            self.result_text.insert(tk.END, f"目錄 : 【合計】\n", "total_row")
            
            # 建立合計的統計行
            plant_headers = self.plants + ['其他', '總計']
            header_line = "廠別 : " + "".join(f"{header:>12}" for header in plant_headers)
            count_line = "筆數 : " + "".join(f"{column_totals[header]:>12,}" for header in plant_headers)
            
            self.result_text.insert(tk.END, f"{header_line}\n", "total_row")
            self.result_text.insert(tk.END, f"{count_line}\n", "total_row")
            self.result_text.insert(tk.END, "=" * 150 + "\n\n", "total_row")
            
            # 將合計列也加入 DataFrame
            total_row = {'目錄': '【合計】'}
            for plant in self.plants:
                total_row[plant] = column_totals[plant]
            total_row['其他'] = column_totals['其他']
            total_row['總計'] = column_totals['總計']
            data_rows.append(total_row)

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

    def start_statistics(self):
        """開始統計"""
        if self.is_running:
            messagebox.showwarning("警告", "統計正在進行中，請稍候...")
            return
        
        # 彈出目錄選擇對話框
        selected_dir = filedialog.askdirectory(
            title="選擇要統計的目錄",
            initialdir=os.getcwd()  # 默認為當前目錄
        )
        
        # 如果使用者取消選擇，則不執行
        if not selected_dir:
            self.progress_var.set("已取消統計")
            return
        
        self.selected_dir = selected_dir

        self.is_running = True
        self.run_button.config(state='disabled')
        self.progress_bar.start()

        # 在背景執行緒中進行統計
        def run_stats():
            try:
                results = self.collect_statistics(self.selected_dir)
                self.root.after(0, lambda: self.display_results(results))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"統計過程中發生錯誤：{str(e)}"))
            finally:
                self.root.after(0, self.finish_statistics)

        thread = threading.Thread(target=run_stats, daemon=True)
        thread.start()

    def finish_statistics(self):
        """完成統計"""
        self.is_running = False
        self.run_button.config(state='normal')
        self.progress_bar.stop()
        
        # 根據是否有結果決定按鈕狀態
        if self.df_result is not None and len(self.df_result) > 0:
            self.save_button.config(state='normal')
            self.mail_button.config(state='normal')
            self.progress_var.set(f"統計完成（共 {len(self.df_result)} 個目錄）")
        else:
            self.save_button.config(state='disabled')
            self.mail_button.config(state='disabled')
            self.progress_var.set("統計完成（無資料）")

    def save_excel(self):
        """將統計結果儲存為 Excel 檔案"""
        if self.df_result is None or len(self.df_result) == 0:
            messagebox.showwarning("無資料", "目前沒有統計結果可以儲存！")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="儲存 Excel 檔案",
            initialfile=f"YAML統計結果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
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
            messagebox.showwarning("無資料", "目前沒有統計結果可以寄送！")
            return
        
        # 定義進度更新回調函數
        def update_progress(message):
            self.progress_var.set(message)
            self.root.update()
        
        # 調用 mainlib 的郵件發送功能
        success, message = mainlib.send_statistics_email(
            df_result=self.df_result,
            parent_window=self.root,
            progress_callback=update_progress
        )
        
        # 更新進度訊息
        self.progress_var.set(message)

    def exit_program(self):
        """離開程式"""
        if self.is_running:
            if messagebox.askyesno("確認", "統計正在進行中，確定要離開嗎？"):
                self.root.quit()
        else:
            self.root.quit()


def main():
    root = tk.Tk()
    app = YAMLStatisticsGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
