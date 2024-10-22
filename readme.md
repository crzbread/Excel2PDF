# Excel2PDF

A Python tool to read Excel files and convert them to encrypted PDFs using Tkinter.

## Versions

Python 3.10

## Installation
1. Clone the repository
    ```shell
    git clone git@github.com:crzbread/Excel2PDF.git
    cd excel2pdf
    ```
2. Create and activate a virtual environment:
    ```shell
    python -m venv venv
    source venv/bin/activate  # On Mac/Linux
    venv\Scripts\activate  # On Windows
    ```
3. Install dependencies:
    ```shell
    pip install -r requirements.txt
    ```
## Usage 
Run the application:

```shell
python main.py
```

A GUI will open, allowing you to select an Excel file.

The application will convert each row of the Excel file into an individual PDF file and encrypt it using the ID number from the Excel data (or a default password if no ID is found).

## How It Works

This section explains the step-by-step process of how the application works, from starting the GUI to generating and encrypting PDF files.

### 1. register_fonts (Initialization step):
- Registers fonts once during program initialization to handle Chinese characters.
### 2. create_gui:
- Establishes the GUI interface using Tkinter, allowing the user to select an Excel file.
### 3. open_file:
- Opens and reads the Excel file selected by the user.
- Calls generate_pdfs to process the data.
### 4. generate_pdfs:
- Iterates through each row of the Excel file.
- For each row, it generates a corresponding PDF file by calling generate_single_pdf.
- After generating each PDF, it encrypts the file using encrypt_pdf.
#### generate_single_pdf:
- Generates an individual PDF for each row of data.
- Formats the data into a table and applies the appropriate styles.
#### encrypt_pdf:
- Encrypts each generated PDF with the ID number found in the data.
- If the ID is missing, uses a default password (12345678).

## todo

關於 UI 的樣式完全沒有調整 需要等候設計稿

## Libraries Used

### 1. Tkinter
簡介：Tkinter 是 Python 的標準圖形用戶界面 (GUI) 套件，提供創建視窗、對話框、按鈕等圖形元件的功能。
本專案中的用途：Tkinter 用於建立圖形介面，讓用戶能夠選擇 Excel 文件，顯示訊息以及處理用戶操作。
官方文檔：[Tkinter 文檔](https://docs.python.org/3/library/tkinter.html)
### 2. Pandas
簡介：Pandas 是一個強大的數據處理和分析庫，提供類似 DataFrame 的數據結構，用於處理表格數據。
本專案中的用途：Pandas 用來讀取 Excel 文件，提取每一行的數據，並將數據處理後傳遞給 PDF 生成模組。
官方文檔：[Pandas 文檔](https://pandas.pydata.org/)
### 3. PyPDF2
簡介：PyPDF2 是一個用於處理已存在 PDF 文件的 Python 庫，提供合併、分割、讀取和加密 PDF 文件的功能。
本專案中的用途：PyPDF2 用於對生成的 PDF 文件進行加密，使用例如身分證字號作為密碼。
官方文檔：[PyPDF2 文檔](https://pypdf2.readthedocs.io/en/3.x/)
### 4. ReportLab
簡介：ReportLab 是一個用於生成 PDF 文件的強大庫，能夠創建自定義的 PDF 佈局，包括字體、表格、圖片等。
本專案中的用途：ReportLab 用來從 Excel 文件中提取數據並生成 PDF 文件，處理 PDF 內的表格、文字和字體格式。
官方文檔：[ReportLab 文檔](https://docs.reportlab.com/)
### 5. OS（操作系統模組）
簡介：os 模組提供與操作系統互動的功能，允許對文件和目錄進行操作，訪問環境變量等。
本專案中的用途：os 用來處理文件路徑，檢查目錄是否存在，以及創建輸出文件夾。
官方文檔：[os 文檔](https://docs.python.org/3/library/os.html)

