import openpyxl
from openpyxl import Workbook
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os


def get_excel_file_path():
    Tk().withdraw()  # ルートウィンドウを表示せずにファイル選択ダイアログを開く
    file_path = askopenfilename()  # ファイル選択ダイアログを表示
    if not file_path:
        print("ファイルが選択されませんでした。")
    return file_path


def create_new_excel_with_sheets():
    original_excel_file_path = get_excel_file_path()
    if not original_excel_file_path:
        return

    original_workbook = openpyxl.load_workbook(original_excel_file_path, data_only=True)
    original_sheet = original_workbook["試験項目"]

    new_workbook = Workbook()
    new_workbook.remove(new_workbook.active)  # デフォルトで作成されるシートを削除

    for row_num in range(8, original_sheet.max_row + 1):
        cell_value = original_sheet.cell(row=row_num, column=1).value
        if cell_value is None:
            break
        new_sheet_title = f"No.{cell_value}"
        new_workbook.create_sheet(title=new_sheet_title)

    original_file_name = os.path.basename(original_excel_file_path)
    new_file_name = f"エビ_{original_file_name}"
    new_workbook.save(new_file_name)
    print(f"新しいExcelファイル '{new_file_name}' を作成しました。")


if __name__ == "__main__":
    create_new_excel_with_sheets()
