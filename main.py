import os
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd
from PyPDF2 import PdfWriter, PdfReader
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph

FONTNAME = "NotoSansTC"


# 讀取 Excel 文件
def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            # 使用 pandas 讀取 Excel 文件
            df = pd.read_excel(file_path, engine="openpyxl")

            # 如果需要看裡面的內容 解開下面的註解
            # for index, row in df.iterrows():
            #     cur_row = row.to_dict()  # 將每一行改為 dict
            #     print(cur_row["姓名"])

            # 生成 PDF
            generate_pdfs(df)
        except Exception as e:
            messagebox.showerror("Error", f"讀取 Excel 文件失敗：{str(e)}")


# 註冊思源黑體
def register_fonts():
    font_path = os.path.join(
        os.path.dirname(__file__), "assets", "fonts", "NotoSansTC-Regular.ttf"
    )

    # 檢查字體文件是否存在
    if not os.path.exists(font_path):
        print(f"字體文件不存在：{font_path}")
    # 註冊字體
    pdfmetrics.registerFont(TTFont(FONTNAME, font_path))


def generate_single_pdf(row, pdf_file):
    # 創建 PDF 文件，A4 大小
    doc = SimpleDocTemplate(pdf_file, pagesize=A4)

    # 設定樣式
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="CustomStyle", fontName=FONTNAME, fontSize=12))

    elements = []

    # 確保資料存在 如果不存在就返回 N/A
    name = str(row.get("姓名", "N/A"))
    salary = str(row.get("薪資", "N/A"))

    # 建立表格資料，第一行為標題
    data = [["姓名", "薪資"], [name, salary]]

    # 建立表格
    table = Table(data)

    # 設定表格樣式 目前沒有設計稿 直接用網路上抄來的
    # (0, 0): 表示從第一列、第一行開始
    # (-1, 0): 表示到第一行的最後一列，-1 表示最後一列
    # (0, 1): 表示從第二行開始（第一行為表頭）
    # (-1, -1): 表示直到最後一行、最後一列，-1 表示最後一列或最後一行
    table.setStyle(
        TableStyle(
            [
                # 設定表頭的背景顏色和文字顏色
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.grey,
                ),  # 表頭背景，範圍：第一行的所有列
                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, 0),
                    colors.whitesmoke,
                ),  # 表頭文字顏色，範圍：第一行的所有列
                # 設定表格文字置中對齊
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),  # 範圍：整個表格
                # 設定表頭和內容字體
                ("FONTNAME", (0, 0), (-1, 0), FONTNAME),  # 表頭字體，範圍：第一行
                (
                    "FONTNAME",
                    (0, 1),
                    (-1, -1),
                    FONTNAME,
                ),  # 表格內容字體，範圍：第二行及之後
                # 設定表頭的底部填充
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, 0),
                    12,
                ),  # 為表頭增加 12px 的 padding
                # 12: 表示填充的距離為 12 像素
                # 設定表格內容的背景顏色
                (
                    "BACKGROUND",
                    (0, 1),
                    (-1, -1),
                    colors.beige,
                ),  # 設定表格內容的背景顏色為米色，範圍是表格的內容部分（即第一行以後的所有列）
                # 設定表格邊框
                ("GRID", (0, 0), (-1, -1), 1, colors.black),  # 表格邊框，範圍：整個表格
            ]
        )
    )

    elements.append(table)

    # 添加說明文字
    description = Paragraph("薪資表格的最後一段說明", styles["CustomStyle"])
    elements.append(description)

    # 生成 PDF 文件
    doc.build(elements)


# 直接處理來自 pandas 的資料
def generate_pdfs(df):
    output_dir = os.path.join(os.path.dirname(__file__), "output")

    # 創建 output 資料夾 (假如沒有的話)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 讀取每一行的資料
    for index, row in df.iterrows():
        # 生成 PDF 文件名
        pdf_filename = f"{row['姓名']}.pdf"
        pdf_file_path = os.path.join(output_dir, pdf_filename)

        # 生成 PDF 此時尚未加密
        generate_single_pdf(row, pdf_file_path)
        print(f"生成 PDF: {pdf_file_path}")

        # 取得身分證字號，若無則使用預設密碼 "12345678"
        id_number = row.get("身分證字號", "12345678")

        # 加密 PDF
        encrypt_pdf(pdf_file_path, id_number)


# 加密 PDF 文件的流程是：
# 1. 先使用 PdfReader 打開並讀取現有的 PDF 文件。
# 2. 然後將內容寫入 PdfWriter 對象，並設置加密。
# 3. 最後將加密後的內容寫回（可以覆蓋原文件或保存為新文件）。
def encrypt_pdf(pdf_file_path, password):
    # 開啟 PDF 文件
    reader = PdfReader(pdf_file_path)
    writer = PdfWriter()

    # 將 PDF 內容放到記憶體裡面
    for page_num in range(len(reader.pages)):
        writer.add_page(reader.pages[page_num])

    # 設置密碼
    writer.encrypt(password)

    # 保存加密後的文件，並覆蓋原本的文件
    with open(pdf_file_path, "wb") as f:
        writer.write(f)

    print(f"加密 PDF 文件: {pdf_file_path}，使用密碼: {password}")


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
    # 註冊字體
    register_fonts()
    create_gui()
