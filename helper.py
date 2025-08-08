### --- ここから実装 ---


import pandas as pd
import zipfile
import io
import re
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from copy import copy


def sanitize_sheet_name(sheet_name: str) -> str:
    """
    OSで禁止されている文字を除去し、255文字以内にする。
    """
    # Windows禁止文字: \\/:*?"<>|
    sanitized = re.sub(r'[\\/:*?"<>|]', '_', sheet_name)
    return sanitized[:255]


def split_excel_to_zip(file) -> bytes:
    """
    Excelファイルを受け取り、各シートを個別のExcelファイルに分割し、
    それらをZipバイト列として返す。元の書式設定を保持する。
    Args:
        file: BytesIOまたはアップロードファイルオブジェクト
    Returns:
        bytes: Zipファイルのバイト列
    Raises:
        ValueError: Excelファイルが不正な場合
    """
    try:
        # openpyxlでワークブックを読み込み
        if hasattr(file, 'read'):
            # ファイルオブジェクトの場合
            workbook = load_workbook(file)
        else:
            # バイトデータの場合
            workbook = load_workbook(io.BytesIO(file))
    except Exception as e:
        raise ValueError(f"Excelファイルの読み込みに失敗しました: {e}")

    sheet_names = workbook.sheetnames
    if not sheet_names:
        raise ValueError("シートが見つかりませんでした。")

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zipf:
        for sheet_name in sheet_names:
            try:
                # 元のシートを取得
                original_sheet = workbook[sheet_name]
                
                # 新しいワークブックを作成
                new_workbook = Workbook()
                new_workbook.remove(new_workbook.active)  # デフォルトシートを削除
                
                # 新しいシートを作成（元のシート名を使用）
                new_sheet = new_workbook.create_sheet(title=sheet_name)
                
                # データと書式をコピー
                for row in original_sheet.iter_rows():
                    for cell in row:
                        new_cell = new_sheet.cell(row=cell.row, column=cell.column)
                        new_cell.value = cell.value
                        
                        # 書式設定をコピー（StyleProxyエラーを回避）
                        if cell.has_style:
                            if cell.font:
                                new_cell.font = copy(cell.font)
                            if cell.border:
                                new_cell.border = copy(cell.border)
                            if cell.fill:
                                new_cell.fill = copy(cell.fill)
                            if cell.number_format:
                                new_cell.number_format = cell.number_format
                            if cell.protection:
                                new_cell.protection = copy(cell.protection)
                            if cell.alignment:
                                new_cell.alignment = copy(cell.alignment)
                
                # 列幅をコピー
                for column in original_sheet.column_dimensions:
                    if column in original_sheet.column_dimensions:
                        new_sheet.column_dimensions[column] = original_sheet.column_dimensions[column]
                
                # 行高をコピー
                for row in original_sheet.row_dimensions:
                    if row in original_sheet.row_dimensions:
                        new_sheet.row_dimensions[row] = original_sheet.row_dimensions[row]
                
                # マージされたセルをコピー
                for merged_range in original_sheet.merged_cells.ranges:
                    new_sheet.merge_cells(str(merged_range))
                
                # シートをバイトデータとして保存
                sheet_buffer = io.BytesIO()
                new_workbook.save(sheet_buffer)
                sheet_buffer.seek(0)
                
                # ZIPに追加
                sheet_filename = sanitize_sheet_name(sheet_name) + ".xlsx"
                zipf.writestr(sheet_filename, sheet_buffer.getvalue())
                
            except Exception as e:
                raise ValueError(f"シート '{sheet_name}' の処理中にエラー: {e}")
    
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

if __name__ == "__main__":
    # テスト用: コマンドラインから実行した場合の動作例
    with open("test.xlsx", "rb") as f:
        zip_bytes = split_excel_to_zip(f)
    with open("sheets.zip", "wb") as f:
        f.write(zip_bytes) 