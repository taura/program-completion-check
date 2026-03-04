#!/usr/bin/env python
import math
import io
import re
import sys
import unicodedata
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
    return df

def set_default(series, val, typ):
    """
    特定の列の空のセルの値をvalにする
    """
    return pd.to_numeric(series, errors="coerce").fillna(val).astype(typ)

def ensure_column(df, c, val):
    """
    df[c] が存在していなければ加える
    """
    if c not in df.columns:
        df[c] = val
    return df

def safe_int(s):
    try:
        return int(s)
    except ValueError:
        return None

def parse_num_range(s):
    lo = 2015
    hi = 2035
    ab = s.split("-")
    if len(ab) == 1:
        ab = ab + ab            # [x] -> [x,x]
    if len(ab) == 2:
        [a, b] = ab
        if a == "" and b == "":
            return None
        a = safe_int(a) if a else lo
        b = safe_int(b) if b else hi
        if a is None or b is None:
            return None
        else:
            return set(range(a, b + 1))
    else:
        return None

def parse_num_set(s):
    """
    None, NaN -> NaN
    "" -> NaN
    "2023" -> [2023]
    "2023,2025" -> [2023, 2025]
    "2021,2023-2025" -> [2021, 2023, 2024, 2025]

    , が ; または 空白もOK とする
    - の前後の空白もOK
    """
    if pd.isnull(s):
        return s
    s = unicodedata.normalize("NFKC", s)
    s = s.strip()
    if s == "":
        return np.nan
    # commas -> spaces
    s = re.sub("[,;]", " ", s)
    # collapse spaces around '-' inside a range-like token
    #    e.g. "3 - 5" -> "3-5", "- 7" -> "-7", "7 -" -> "7-"
    s = re.sub(r"(?P<a>\d*)\s*-\s*(?P<b>\d*)", r"\g<a>-\g<b>", s)
    S = set()
    for rng in s.split():
        dS = parse_num_range(rng)
        if dS is None:
            return "ERROR"      # propagate error
        S.update(dS)
    return sorted(S)

def normalize_columns(df):
    """
    カラム名の全角は半角に
    e.g., 共通ＩＤ -> 共通ID 
    もともと 共通ID が存在していた場合はやらない
    """
    add_columns = []
    for c in df.columns:
        # 全角 -> 半角
        nc = unicodedata.normalize("NFKC", c)
        if nc != c and nc not in df.columns:
            # c -> nc
            add_columns.append((c, nc))
    for c, nc in add_columns:
        df[nc] = df[c]
    return df

def validate_program_students(url, sheet=0):
    """
    プログラムの認定対象科目を記述した表の読み込み
    プログラムごとに作る
    以下の cols のカラムが必須
    """
    df = open_sheet(url, sheet)
    df = normalize_columns(df)
    cols = ["共通ID", "学籍番号", "学生氏名"]
    df = check_columns(df, cols)
    return df

def validate_utas_grade(url, sheet=0):
    df = open_sheet(url, sheet)
    df = normalize_columns(df)
    cols = ["共通ID", "学籍番号", "学生氏名", "学生氏名カナ", "学生所属",
            "年度", "時間割所属", "時間割コード", "開講科目名",
            "科目コード", "科目名", "主担当教員名", "主担当教員共通ID",
            "開講区分名", "合否区分", "単位数"]
    df = check_columns(df, cols)
    return df

def validate_program_courses(url, sheet=0):
    df = open_sheet(url, sheet)
    df = normalize_columns(df)
    df = ensure_column(df, "開講年度", "")
    cols = ["科目コード", "科目名", "開講年度"]
    df = check_columns(df, cols)
    df["開講年度リスト"] = df["開講年度"].apply(parse_num_set)
    err_rows = df["開講年度リスト"] == "ERROR"
    if any(err_rows):
        for i, row in df[err_rows].iterrows():
            year = row["開講年度"]
            print(f'エラー: {i+1} 行目の開講年度 "{year}" が不正', file=sys.stderr)
        return None
    return df

def normalize_code(x, width=15):
    """
    学籍番号などのキー正規化

    - 全角 → 半角（NFKC）
    - 純数字なら0埋めして width 桁
    - width 超は警告
    - NaN はそのまま返す
    """
    if pd.isnull(x):
        return x
    # 文字列化 + 前後空白除去
    s = str(x).strip()
    # Unicode正規化（全角→半角）
    s = unicodedata.normalize("NFKC", s)
    # 純数字なら0埋め
    if s.isdigit():
        if len(s) >= width:
            print(f"コードが {width} 文字以上あります: {s}", file=sys.stderr)
            return s
        return s.zfill(width)
    return s

def normalize_utac(x):
    """
    NaN -> NaN
    全角 -> 半角
    前後空白除去
    10桁以下 -> 10桁
    10桁以下@utac.u-tokyo.ac.jp -> 10桁
    それ以外 -> error
    """
    if pd.isnull(x):
        return x
    # 文字列化 + 前後空白除去
    s = str(x)
    # Unicode正規化（全角→半角）
    s = unicodedata.normalize("NFKC", s)
    s = s.strip()
    # 純数字なら0埋め
    m = re.match(r"(?P<n>\d{1,10})(@utac\.u\-tokyo\.ac\.jp)?$", s)
    if m:
        n = m.group("n")
        n = "0" * (10 - len(n)) + n
        return n
    else:
        return None         # error

def year_in_year_list(r):
    years = r["開講年度リスト_科目一覧"]
    year = r["年度_UTAS"]
    if type(years) is type(0.0):
        assert(math.isnan(years)), years
        return False
    else:
        assert(type(years) is type([]))
        return (years == []) or (year in years)

def cat_unique(s):
    """
    s : Series
    """
    return "; ".join([str(x) for x in pd.unique(s.dropna())])
    
def do_check(program_students, utas_grade, program_courses,
             req_credits):
    """
    program_students : dataframe
    utas_grade : dataframe
    program_courses : dataframe
    req_credits : 認定に必要なクレジット
    """
    ps = program_students
    ug = utas_grade
    pc = program_courses
    ps["共通ID_正規化"] = ps["共通ID"].apply(normalize_utac)
    ug["共通ID_正規化"] = ug["共通ID"].apply(normalize_utac)
    ug["科目コード_正規化"] = ug["科目コード"].apply(normalize_code)
    pc["科目コード_正規化"] = pc["科目コード"].apply(normalize_code)
    ps = ps.add_suffix("_登録学生一覧")
    ug = ug.add_suffix("_UTAS")
    pc = pc.add_suffix("_科目一覧")
    # 登録学生一覧にUTAS成績をjoin
    ps_ug = pd.merge(ps, ug, how="left",
                     left_on="共通ID_正規化_登録学生一覧",
                     right_on="共通ID_正規化_UTAS")
    # それに認定科目一覧をjoin
    df = pd.merge(ps_ug, pc, how="left",
                  left_on="科目コード_正規化_UTAS",
                  right_on="科目コード_正規化_科目一覧")
    df["認定対象"] = np.where(df.apply(year_in_year_list, axis=1), 1, np.nan)
    df["認定"] = df["認定対象"].where(df["合否区分_UTAS"] == "合格")
    df["認定単位"] = df["認定"] * df["単位数_UTAS"]
    ps_columns = ps.columns
    ug_columns = ["学籍番号_UTAS", "学生氏名_UTAS", "学生氏名カナ_UTAS", "学生所属_UTAS"]
    agg_spec = { c : (c, cat_unique) for c in list(ps_columns) + ug_columns }
    agg_spec.update({"認定単位" : ("認定単位", "sum")})
    credit = df.groupby(["共通ID_正規化_登録学生一覧"], dropna=False).agg(**agg_spec)
    credit["認定結果"] = np.where(credit["認定単位"] >= req_credits, 1, np.nan)
    result_xlsx = "認定単位.xlsx"
    with pd.ExcelWriter(result_xlsx) as writer:
        credit.to_excel(writer, sheet_name="認定単位")
        df.to_excel(writer, sheet_name="詳細", index=False)
    return (credit, result_xlsx)
    
def main():
    program_students_url = "登録学生一覧.xlsx"
    program_students_sheet = 0
    program_students = validate_program_students(program_students_url, sheet=program_students_sheet)
    assert(program_students is not None)
    utas_grade_url = "utas_grade.xlsx"
    utas_grade_sheet = 0
    utas_grade = validate_utas_grade(utas_grade_url, sheet=utas_grade_sheet)
    assert(utas_grade is not None)
    program_courses_url = "科目一覧.xlsx"
    program_courses_sheet = 0
    program_courses = validate_program_courses(program_courses_url, sheet=program_courses_sheet)
    assert(program_courses is not None)
    required_credits = 3
    do_check(program_students, utas_grade, program_courses, required_credits)

if 0 or __name__ == "__main__":
    main()
    

