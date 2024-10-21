import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd


# 打开并读取 Excel 文件
def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            # 使用 pandas 讀取 Excel 文件
            df = pd.read_excel(file_path, engine="openpyxl")

            for index, row in df.iterrows():
                cur_row = row.to_dict()  # 將每一行改為 dict
                print(cur_row["姓名"])

            messagebox.showinfo("Success", f"成功讀取 Excel 文件：{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"讀取 Excel 文件失敗：{str(e)}")


# 创建 Tkinter 界面
def create_gui():
    app = tk.Tk()
    app.title("Excel 工具")

    width = 600
    height = 800
    app.geometry(f"{width}x{height}")

    open_button = tk.Button(app, text="選擇 Excel 文件", command=open_file)
    open_button.pack(pady=20)
    app.mainloop()


if __name__ == "__main__":
    create_gui()
