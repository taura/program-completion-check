#!/usr/bin/env python
import io
import re
import sys
import pandas as pd
import numpy as np

GC = None

def get_gclient():
    """
    Google spreadsheetのデータを開く
    colab でもローカルでも実行できるように
    """
    import gspread
    in_colab = 0
    try:
        # Colab
        from google.colab import auth
        in_colab = 1
    except ImportError:
        pass
    if in_colab:
        auth.authenticate_user()
        from google.auth import default
        creds, _ = default()
        return gspread.authorize(creds)
    else:
        from google_auth_oauthlib.flow import InstalledAppFlow
        SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly",
                  "https://www.googleapis.com/auth/drive.readonly"]
        flow = InstalledAppFlow.from_client_secrets_file(
            "credentials.json", SCOPES
        )
        creds = flow.run_local_server(port=0)
        return gspread.authorize(creds)

def set_gclient():
    global GC
    GC = get_gclient()
    return GC

def extract_gdrive_file_id(url):
    """
    https://docs.google.com/spreadsheets/d/1HUi7QBmFGvM9QReHjTwFRLCH1-UYqYVG/...
    --> 1HUi7QBmFGvM9QReHjTwFRLCH1-UYqYVG
    """
    m = re.search(r"/file/d/([a-zA-Z0-9_-]+)", url)
    if m:
        return m.group(1)
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
    if m:
        return m.group(1)
    m = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", url)
    if m:
        return m.group(1)
    raise ValueError(f"Cannot extract file id from URL: {url}")

def get_creds_from_gclient(gclient):
    """
    gclient.http_client.auth
    """
    http = getattr(gclient, "http_client", None)
    if http is None:
        raise RuntimeError("gspread client has no http_client; cannot obtain credentials.")
    # よくある候補を順に探す
    for attr in ("auth", "credentials", "_auth", "_credentials"):
        creds = getattr(http, attr, None)
        if creds is not None:
            return creds
    # 念のため client 側も見る
    for attr in ("auth", "credentials"):
        creds = getattr(gclient, attr, None)
        if creds is not None:
            return creds
    raise RuntimeError("Could not find credentials in gspread client/http_client.")

def download_gdrive_file_as_bytes(file_id, creds):
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    fh = io.BytesIO()
    request = drive.files().get_media(fileId=file_id)
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return fh.getvalue()

def gdrive_get_metadata(file_id, creds):
    from googleapiclient.discovery import build
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    return drive.files().get(fileId=file_id, fields="id,name,mimeType").execute()

def is_gdrive_url(url):
    return url.startswith("https://")
    
def is_gspreadsheet(file_id, creds):
    meta = gdrive_get_metadata(file_id, creds)
    mime = meta["mimeType"]
    return mime == "application/vnd.google-apps.spreadsheet"

def open_sheet(url, sheet):
    """
    Google Sheets / Excel (even if opened via docs.google.com/spreadsheets/) を
    Drive の mimeType で判定して DataFrame として読む。

    sheet: int (0-based) or str (sheet name)
    """
    if not is_gdrive_url(url):
        return pd.read_excel(url, sheet_name=sheet)
    if GC is None:
        gclient = set_gclient()
    else:
        gclient = GC
    file_id = extract_gdrive_file_id(url)
    creds = get_creds_from_gclient(gclient)
    if is_gspreadsheet(file_id, creds):
        # Google Sheets
        doc = gclient.open_by_key(file_id)
        ws = doc.get_worksheet(sheet) if isinstance(sheet, int) else doc.worksheet(sheet)
        return pd.DataFrame(ws.get_all_records())
    else:
        # Excel (xlsx etc.)
        data = download_gdrive_file_as_bytes(file_id, creds)
        return pd.read_excel(io.BytesIO(data), sheet_name=sheet)

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
            print(f"\n\nエラー: プログラム登録学生に「{c}」列がありません ({df.columns})\n\n", file=sys.stderr)
            assert(c in df.columns), (c, df.columns)
    return df[columns]

def set_default(series, val, typ):
    """
    特定の列の空のセルの値をvalにする
    """
    return pd.to_numeric(series, errors="coerce").fillna(val).astype(typ)

def validate_program_students(url, sheet=0):
    """
    プログラムの認定対象科目を記述した表の読み込み
    プログラムごとに作る
    以下の cols のカラムが必須
    """
    df = open_sheet(url, sheet)
    cols = ["学籍番号", "学生氏名"]
    df = check_columns(df, cols)
    return df

def validate_utas_grade(url, sheet=0):
    df = open_sheet(url, sheet)
    cols = ["学籍番号", "学生氏名", "学生氏名カナ", "学生所属",
            "年度", "時間割所属", "時間割コード", "開講科目名",
            "科目コード", "科目名", "主担当教員名", "主担当教員共通ID",
            "開講区分名", "合否区分", "単位数"]
    df = check_columns(df, cols)
    return df

def validate_program_subjects(url, sheet=0):
    df = open_sheet(url, sheet)
    cols = ["科目コード", "開始年度", "終了年度"]
    df = check_columns(df, cols)
    df["開始年度"] = set_default(df["開始年度"], 0,    int)
    df["終了年度"] = set_default(df["終了年度"], 9999, int)
    return df

def do_check(program_students, utas_grade, program_courses):
    sg = pd.merge(program_students, utas_grade, how="left",
                  left_on="学籍番号", right_on="学籍番号", suffixes=("_", ""))
    df = pd.merge(sg, program_courses, how="left",
                  left_on="科目コード", right_on="科目コード", suffixes=("", "_"))
    df["認定対象"] = np.where((df["開始年度"] <= df["年度"]) & (df["年度"] <= df["終了年度"]), 1, np.nan)
    df["認定"] = df["認定対象"].where(df["合否区分"] == "合格")
    df["認定単位"] = df["認定"] * df["単位数"]
    credit = df.groupby(["学籍番号", "学生氏名", "学生氏名カナ"])[["認定単位"]].sum()
    result_xlsx = "認定単位.xlsx"
    with pd.ExcelWriter(result_xlsx) as writer:
        credit.to_excel(writer, sheet_name="認定単位")
        df.to_excel(writer, sheet_name="詳細", index=False)
    return (credit, result_xlsx)
    
