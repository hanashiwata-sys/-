import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
import glob
import os
import win32com.client as win32
import gc

cols = [
    "出荷日",
    "向先",
    "注文番号",
    "受注番号",
    "客先品名",
    "品目名称",
    "数量",
    "単位",
    "ロットNo",
    "引当",
    "備考"
]

def read_csv_file():
    csv_files = glob.glob("CSVフォルダ/*.csv")

    if not csv_files:
        raise SystemExit("CSVフォルダ内にCSVファイルが見つかりません。")

    latest_csv = max(csv_files, key=os.path.getmtime)

    try:
        try:
            df = pd.read_csv(latest_csv, encoding="utf-8")
            print("文字コード UTF-8 で読み込みました。")
        
        except UnicodeDecodeError:
            df = pd.read_csv(latest_csv, encoding="cp932")
            print("文字コード cp932 で読み込みました。")
            
    except pd.errors.EmptyDataError:
        raise SystemExit("CSVファイルにデータがありません。")

    except pd.errors.ParserError:
        raise SystemExit("CSVファイルの形式が正しくありません。")

    # データ行が存在するか確認
    if df.empty:
        raise SystemExit("CSVファイルにデータ行がありません。")
    
    return df, latest_csv

# CSV読み込み関数を実行
df, latest_csv = read_csv_file()

# 出力ファイル名を取得
base_name = os.path.splitext(os.path.basename(latest_csv))[0]
output_file = f"{base_name}.xlsx"

# 列名の前後にある空白を削除
df.columns = df.columns.str.strip()

def check_required_columns(df, required_cols):
    
    # 必要な列が存在するか確認
    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        raise SystemExit(f"CSVに必要な列がありません: {missing_cols}")
    
    return df

df = check_required_columns(df, cols)

df_out = df[cols].copy()

# 数値チェック用の関数
def check_numeric(df, column_name):
    
    # 変換前のデータを保存
    original_values = df[column_name].copy()
    
    # 数値に変換
    df[column_name] = pd.to_numeric(
        df[column_name],
        errors="coerce"
    )
    
    # 数値を変換できなかった行を取得
    error_rows = df[df[column_name].isna()].copy()
    
    # 元の値を保存
    error_rows["元の値"] = original_values[error_rows.index]
    
    # CSVの行番号に合わせる
    error_rows.index = error_rows.index + 2
    
    # エラーがあれば表示して終了
    if not error_rows.empty:
        
        print(f"{column_name}の入力エラーが見つかりました。")
        
        for row_number, row in error_rows.iterrows():
                
                original_value = row["元の値"]
                
                if pd.isna(original_value):
                    print(f"{row_number}行目:{column_name}が未入力です。")
                    
                else:
                    print(
                        f"{row_number}行目:{column_name}が数値ではありません"
                        f"(入力値:{original_value})"
                    )
            
        raise SystemExit(f"{column_name}のデータを確認してください。")   
    
    return df   

def create_excel(df_out, output_file, base_name):
    
    try:
        df_out.to_excel(
            output_file,
            sheet_name=base_name[:31],
            index=False
        )
        
    except PermissionError:
        raise SystemExit(
            "Excelファイルを保存できません。\n"
            "出力先のExcelファイルが開いていないか確認してください。"
        )
        
    except OSError as e:
        raise SystemExit(
            f"Excelファイルの保存中にエラーが発生しました。\n"
            f"詳細: {e}"
        )
        
    wb = load_workbook(output_file)
    ws = wb.active

    # A4横向き
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

    # すべての列を1ページ
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    # ヘッダー
    ws.oddHeader.left.text = "&F\n&A"

    # フッター
    ws.oddFooter.left.text = "&D\n&T"

    # 罫線
    thin = Side(style="thin")

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        
        for cell in col:
            if cell.value is not None:
                cell_length = 0
                
                for char in str(cell.value):
                    if ord(char) > 255:
                        cell_length += 2
                    else:
                        cell_length += 1
                        
                max_length = max(max_length, cell_length)
                
        ws.column_dimensions[column].width = max_length * 1.2 +2 
        
    try:
        wb.save(output_file)
        print(f"Excelファイルを保存しました: {output_file}")
        
    except PermissionError:
        raise SystemExit(
            "Excelファイルを保存できません。\n"
            "出力先のExcelファイルが開いていないか確認してください。"
        )
        
    except OSError as e:
        raise SystemExit(
            f"Excelファイルの保存中にエラーが発生しました。\n"
            f"詳細: {e}"
        )
        
    finally:
        wb.close()            

# 数量をチェック
df_out = check_numeric(df_out, "数量")

# Excelファイルを作成
create_excel(df_out, output_file, base_name)

def create_pdf(output_file):
        
    excel = None
    wb_pdf = None

    try:
        excel = win32.DispatchEx("Excel.Application")
        
        wb_pdf = excel.Workbooks.Open(
            os.path.abspath(output_file)
        )
        
        pdf_file = output_file.replace(".xlsx", ".pdf")
        
        wb_pdf.ExportAsFixedFormat(
            0,
            os.path.abspath(pdf_file)
        )
        
    except Exception as e:
        raise SystemExit(
            "PDFファイルの作成中にエラーが発生しました。\n"
            f"詳細: {e}"
        )
        
    finally:
        if wb_pdf:
            wb_pdf.Close(False)
            
        if excel:
            excel.Quit()
            
        del wb_pdf
        del excel
        
        gc.collect()
        
    return pdf_file

pdf_file = create_pdf(output_file)

print(f"Excel作成完了: {output_file}")
print(f"PDF作成完了: {pdf_file}")