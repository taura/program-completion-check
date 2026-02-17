#!/usr/bin/env python
import sys
import gspread
import pandas as pd
import numpy as np

def get_client():
    """
    Google spreadsheetのデータを開く
    colab でもローカルでも実行できるように
    """
    try:
        # Colab
        from google.colab import auth
        auth.authenticate_user()
        from google.auth import default
        creds, _ = default()
        return gspread.authorize(creds)
    except ImportError:
        # Local
        from google_auth_oauthlib.flow import InstalledAppFlow
        SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
        flow = InstalledAppFlow.from_client_secrets_file(
            "credentials.json", SCOPES
        )
        creds = flow.run_local_server(port=0)
        return gspread.authorize(creds)

def open_sheet(client, url, sheet):
    # 科目表のGoogle spreadsheet
    doc = client.open_by_url(url)
    if type(sheet) is type(0):
        sh = doc.get_worksheet(sheet)
    else:
        sh = doc.worksheet(sheet)
    # 1枚目のシートをDFに
    subjects = pd.DataFrame(sh.get_all_records())
    return subjects

"""

プログラムに登録された学生 (STUDENTS_URL) と UTAS の成績 (UTAS_GRADE_URL) を 学籍番号で JOIN してプログラムに登録された学生の全ての成績を一覧にする

その中で 科目表に登録された科目だけに限定して集計する

"""

def check_columns(df, columns):
    """
    df に columns があることを確認
    それらのcolumn だけを取り出す(余分なcolumnを消す)
    """
    for c in columns:
        if c not in df.columns:
            print(f"プログラム登録学生に「{c}」列がありません ({df.columns})", file=sys.stderr)
            assert(c in df.columns), (c, df.columns)
    return df[columns]

def set_default(series, val, typ):
    """
    特定の列の空のセルの値をvalにする
    """
    return pd.to_numeric(series, errors="coerce").fillna(val).astype(typ)

def validate_program_students(gc, url, sheet=0):
    """
    プログラムの認定対象科目を記述した表の読み込み
    プログラムごとに作る
    以下の cols のカラムが必須
    """
    df = open_sheet(gc, url, sheet)
    cols = ["学籍番号", "学生氏名"]
    df = check_columns(df, cols)
    return df

def validate_utas_grade(gc, url, sheet=0):
    df = open_sheet(gc, url, sheet)
    cols = ["学籍番号", "学生氏名", "学生氏名カナ", "学生所属",
            "年度", "時間割所属", "時間割コード", "開講科目名",
            "科目コード", "科目名", "主担当教員名", "主担当教員共通ID",
            "開講区分名", "合否区分", "単位数"]
    df = check_columns(df, cols)
    return df

def validate_program_subjects(gc, url, sheet=0):
    df = open_sheet(gc, url, sheet)
    cols = ["科目コード", "開始年度", "終了年度"]
    df = check_columns(df, cols)
    df["開始年度"] = set_default(df["開始年度"], 0,    int)
    df["終了年度"] = set_default(df["終了年度"], 9999, int)
    return df

def do_check(program_students, utas_grade, program_courses):
    sg = pd.merge(program_students, utas_grade, how="left", left_on="学籍番号", right_on="学籍番号", suffixes=("_", ""))
    df = pd.merge(sg, program_courses, how="left", left_on="科目コード", right_on="科目コード")
    df["認定対象"] = np.where((df["開始年度"] <= df["年度"]) & (df["年度"] <= df["終了年度"]), 1, np.nan)
    df["認定"] = df["認定対象"].where(df["合否区分"] == "合格")
    df["認定単位"] = df["認定"] * df["単位数"]
    credit = df.groupby(["学籍番号", "学生氏名", "学生氏名カナ"])[["認定単位"]].sum()
    result_xlsx = "認定単位.xlsx"
    with pd.ExcelWriter(result_xlsx) as writer:
        credit.to_excel(writer, sheet_name="認定単位")
        df.to_excel(writer, sheet_name="詳細", index=False)
    return (credit, result_xlsx)
    
