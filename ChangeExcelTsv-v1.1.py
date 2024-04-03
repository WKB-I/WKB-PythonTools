import os
from tkinter import Tk, filedialog, messagebox
import pandas as pd
import csv


def select_folder():
    root = Tk()
    root.withdraw()  # メインウィンドウを表示しないようにする
    folder_path = filedialog.askdirectory(title="フォルダを選択してください")

    if folder_path:
        convert_excel_to_tsv(folder_path)


def convert_excel_to_tsv(folder_path):
    try:
        files = os.listdir(folder_path)

        for file in files:
            if file.endswith(".xlsx"):
                excel_path = os.path.join(folder_path, file)
                tsv_name = os.path.splitext(file)[0] + ".tsv"
                tsv_path = os.path.join(folder_path, tsv_name)

                # ExcelファイルをDataFrameとして読み込む
                df = pd.read_excel(excel_path)

                # セル内の " をエスケープする
                #df = df.applymap(lambda x: str(x).replace('"', '""'))

                # DataFrameをタブ区切りのTSVファイルとして保存する
                df.to_csv(tsv_path, sep='\t', index=False, encoding='utf-8', quoting=csv.QUOTE_NONE)

                print(f'Converted Excel to TSV: {file} -> {tsv_name}')

        messagebox.showinfo('メッセージ', '処理が完了しました。ExcelファイルをTSVファイルに変換しました。')

    except Exception as e:
        print(str(e))
        messagebox.showinfo(f'エラーが発生しました: {str(e)}')


if __name__ == "__main__":
    select_folder()
