# yaml005.py

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import yaml
import threading
import datetime
import os

# --- 設定基本資訊 ---
current_path = os.getcwd()
date = datetime.datetime.now().strftime('%Y%m%d')

# --- 定義 YAML 轉 Excel 的核心功能 ---
def transform_yaml_to_excel(yaml_path):
    """將 YAML 檔案轉換為 Excel 檔案。"""
    
    # 讀取 YAML 檔案
    with open(yaml_path, 'r', encoding='utf-8') as file:
        data = yaml.safe_load(file)

    # 處理巢狀 YAML 資料
    flattened_data = []
    for item in data:
        source_data = item.get('_source', {})
        # 將 _source 內所有鍵值對展開
        flattened_item = {
            'ID': item.get('_id'),
            **source_data
        }
        flattened_data.append(flattened_item)

    # 建立 DataFrame
    df = pd.DataFrame(flattened_data)
    
    # 建立輸出檔案名稱
    # 從 YAML 檔名中提取部分作為輸出檔名
    base_name = os.path.splitext(os.path.basename(yaml_path))[0]
    output_filename = f"{base_name}.xlsx"
    output_path = os.path.join(os.path.dirname(yaml_path), output_filename)

    # 儲存為 Excel 檔案
    df.to_excel(output_path, index=False)
    print(f"檔案已成功轉換並儲存為：{output_path}")
    return output_path

# --- 主程式的視窗與功能 ---
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # 視窗設定
        self.title("YAML轉換Excel工具")
        self.geometry("700x500")
        
        # 儲存查詢結果的變數
        self.output_file = None
        
        self.create_widgets()

    def create_widgets(self):
        # 標題
        title_label = ttk.Label(self, text='YAML轉換Excel工具', font=('Arial', 18, 'bold'))
        title_label.pack(pady=10)

        # 按鈕框架
        button_frame = ttk.Frame(self)
        button_frame.pack(pady=10)

        # 查詢按鈕
        btn_query = ttk.Button(button_frame, text='查詢', command=self.query_file_thread)
        btn_query.pack(side=tk.LEFT, padx=10)

        # 狀態訊息標籤
        self.status_label = ttk.Label(self, text="準備就緒...", font=('Arial', 12))
        self.status_label.pack(pady=10)

    def query_file_thread(self):
        """用線程執行檔案處理，避免 UI 凍結。"""
        self.status_label.config(text="正在處理檔案，請稍候...")
        thread = threading.Thread(target=self.query_file)
        thread.daemon = True
        thread.start()

    def query_file(self):
        """開啟檔案對話框並執行 YAML 轉換。"""
        try:
            # 讓使用者選擇 YAML 檔案
            yaml_file_path = filedialog.askopenfilename(
                title="選擇要轉換的 YAML 檔案",
                filetypes=[("YAML Files", "*.yaml"), ("All Files", "*.*")]
            )

            if not yaml_file_path:
                self.status_label.config(text="操作已取消。")
                return

            # 執行轉換
            output_file_path = transform_yaml_to_excel(yaml_file_path)
            self.output_file = output_file_path
            
            # 更新狀態
            self.status_label.config(text=f"轉換完成！檔案路徑：{output_file_path}")
            messagebox.showinfo("完成", f"檔案已成功轉換！\n儲存路徑：\n{output_file_path}")
            
        except Exception as e:
            self.status_label.config(text="發生錯誤，請檢查檔案格式！")
            messagebox.showerror("錯誤", f"發生錯誤：{e}")

# --- 運行應用程式 ---
if __name__ == "__main__":
    app = App()
    app.mainloop()