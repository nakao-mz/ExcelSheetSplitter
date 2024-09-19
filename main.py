import tkinter.messagebox

import openpyxl
from openpyxl import Workbook
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os


def get_excel_file_path():
    """
    Excelファイルを選択するためのダイアログを表示し、選択されたファイルのパスを返す関数。
    """
    Tk().withdraw()  # ルートウィンドウを表示しないように設定
    file_path = askopenfilename()  # ファイル選択ダイアログを表示し、選択されたファイルのパスを取得
    if not file_path:  # ファイルが選択されなかった場合
        print("ファイルが選択されませんでした。")
    return file_path  # 選択されたファイルのパスを返す


def create_new_excel_with_sheets():
    """
    Excelファイルを読み込み、特定のシートの値を基に新しいExcelファイルを作成する関数。
    """
    original_excel_file_path = get_excel_file_path()  # 元のExcelファイルのパスを取得
    if not original_excel_file_path:  # ファイルが選択されなかった場合は処理を終了
        return

    original_workbook = openpyxl.load_workbook(original_excel_file_path, data_only=True)  # 元のExcelファイルを読み込む（数式ではなく値を取得）
    original_sheet = original_workbook["試験項目"]  # "試験項目"シートを取得

    new_workbook = Workbook()  # 新しいExcelブックを作成
    new_workbook.remove(new_workbook.active)  # デフォルトで作成されるシートを削除

    for row_num in range(8, original_sheet.max_row + 1):  # 8行目から最終行までループ
        cell_value = original_sheet.cell(row=row_num, column=1).value  # 1列目の値を取得
        if cell_value is None:  # 値が空の場合はループを抜ける
            break
        new_sheet_title = f"{cell_value}"  # 新しいシートのタイトルを設定
        new_workbook.create_sheet(title=new_sheet_title)  # 新しいシートを作成

    original_file_name = os.path.basename(original_excel_file_path)  # 元のファイル名を取得
    original_dir_path = os.path.dirname(original_excel_file_path)  # 元のファイルのディレクトリパスを取得
    new_file_name = f"エビ_{original_file_name}"  # 新しいファイル名を作成

    new_file_path = os.path.join(original_dir_path, new_file_name)  # 新しいファイルのパスを作成
    new_workbook.save(new_file_path)  # 新しいExcelファイルを保存
    tkinter.messagebox.showinfo("完了", f"新しいExcelファイルを作成しました。\n{new_file_path}")  # 完了メッセージを表示


if __name__ == "__main__":  # このスクリプトが直接実行された場合のみ以下の処理を実行
    create_new_excel_with_sheets()  # メインの処理を実行
