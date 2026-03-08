#!/usr/bin/env python
"""
main.py --- main file for program-completion-check

プログラムに登録された学生 (STUDENTS_URL) と UTAS の成績 (UTAS_GRADE_URL) を 学籍番号で JOIN してプログラムに登録された学生の全ての成績を一覧にする
その中で 科目表に登録された科目だけに限定して集計する
"""

import math
import io
import re
import sys
import unicodedata
import pandas as pd
import numpy as np

GC = None

def make_gclient():
    """
    Google spreadsheetのデータを開く
    colab でもローカルでも実行できるように
    """
    # pylint: disable=import-outside-toplevel
    import gspread
    in_colab = 0
    try:
        # Colab
        from google.colab import auth
        in_colab = 1
    except ImportError:
        pass
    if in_colab:                # pylint: disable=no-else-return
        auth.authenticate_user()
        from google.auth import default
        creds, _ = default()
        return gspread.authorize(creds)
    else:
        from google_auth_oauthlib.flow import InstalledAppFlow
        scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly",
                  "https://www.googleapis.com/auth/drive.readonly"]
        flow = InstalledAppFlow.from_client_secrets_file(
            "credentials.json", scopes
        )
        creds = flow.run_local_server(port=0)
        return gspread.authorize(creds)

def ensure_gclient():
    """
    GC に google client がセットされている状態を作る
    """
    # pylint: disable=global-statement
    global GC
    if GC is None:
        GC = make_gclient()
    return GC

def extract_gdrive_file_id(url):
    """
    gdrive の URL から file id 部分取り出し
    
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
    gclient から credential を取り出す怪しい方法
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
    """
    gdrive の file (Excel とか) の中身を取り出す
    """
    # pylint: disable=import-outside-toplevel
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
    """
    gdrive の file の meta data (ファイルのタイプ) を取り出す
    """
    # pylint: disable=import-outside-toplevel
    from googleapiclient.discovery import build
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    return drive.files().get(fileId=file_id, fields="id,name,mimeType").execute()

def is_gdrive_url(url):
    """
    URL か普通のファイル名か
    """
    return url.startswith("https://")

def is_gspreadsheet(file_id, creds):
    """
    Google spreadsheet か否 (Excel) か
    """
    meta = gdrive_get_metadata(file_id, creds)
    mime = meta["mimeType"]
    return mime == "application/vnd.google-apps.spreadsheet"

def open_sheet(url, sheet):
    """
    ローカルファイル (Excel)
    Google Drive の spreadsheet
    Google Drive 上のExcel
    どれでも読める関数
    
    Google Sheets / Excel (even if opened via docs.google.com/spreadsheets/) を
    Drive の mimeType で判定して DataFrame として読む。

    sheet: int (0-based) or str (sheet name)
    """
    if not is_gdrive_url(url):
        # ローカルファイル
        return pd.read_excel(url, sheet_name=sheet)
    # URL -> Excel or spreadsheet
    gclient = ensure_gclient()
    file_id = extract_gdrive_file_id(url)
    creds = get_creds_from_gclient(gclient)
    if is_gspreadsheet(file_id, creds): # pylint: disable=no-else-return
        # Google Sheets
        doc = gclient.open_by_key(file_id)
        ws = doc.get_worksheet(sheet) if isinstance(sheet, int) else doc.worksheet(sheet)
        return pd.DataFrame(ws.get_all_records())
    else:
        # Excel (xlsx etc.)
        data = download_gdrive_file_as_bytes(file_id, creds)
        return pd.read_excel(io.BytesIO(data), sheet_name=sheet)

def check_columns(df, columns):
    """
    df に columns があることを確認
    """
    for c in columns:
        if c not in df.columns:
            print(f"""

エラー: 必須カラム "{c}" がありません ({df.columns})

""", file=sys.stderr)
            assert(c in df.columns), (c, df.columns)
    return df

def set_default(series, val, typ):
    """
    特定の列の空のセルの値をvalにする
    """
    return pd.to_numeric(series, errors="coerce").fillna(val).astype(typ)

def ensure_column(df, c, val):
    """
    カラム c (df[c]) が存在していなければ加える. 値は val
    """
    if c not in df.columns:
        df[c] = val
    return df

def safe_int(s):
    """
    文字列 -> int; 失敗したら None
    """
    try:
        return int(s)
    except ValueError:
        return None

def is_real_nan(x):
    """
    本当に NaN か?
    (pd.isnull(x) は x がsequenceの時、各値が nan かどうかを並べたsequenceを
    返してしまう. ここではそれはきっぱり False になる)
    """
    try:
        return math.isnan(x)
    except TypeError:
        return False

def parse_num_range(s):
    """
    "2023-2025" -> [2023,2024,2025] みたいなこと
    """
    lo = 2015
    hi = 2035
    ab = s.split("-")
    if len(ab) == 1:
        ab = ab + ab            # [x] -> [x,x]
    if len(ab) == 2:            # pylint: disable=no-else-return
        [a, b] = ab
        if a == "" and b == "":
            return None
        a = safe_int(a) if a else lo
        b = safe_int(b) if b else hi
        if a is None or b is None:
            return None
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
    if is_real_nan(s):
        return s
    s = unicodedata.normalize("NFKC", str(s))
    s = s.strip()
    if s == "":
        return np.nan
    # commas -> spaces
    s = re.sub("[,;]", " ", s)
    # collapse spaces around '-' inside a range-like token
    #    e.g. "3 - 5" -> "3-5", "- 7" -> "-7", "7 -" -> "7-"
    s = re.sub(r"(?P<a>\d*)\s*-\s*(?P<b>\d*)", r"\g<a>-\g<b>", s)
    num_set = set()
    for rng in s.split():
        ds = parse_num_range(rng)
        if ds is None:
            return "ERROR"      # propagate error
        num_set.update(ds)
    return sorted(num_set)

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

def normalize_code(x, width=15):
    """
    学籍番号などのキー正規化

    - 全角 → 半角（NFKC）
    - 純数字なら0埋めして width 桁
    - width 超は警告
    - NaN はそのまま返す
    """
    if is_real_nan(x):
        return x
    # 文字列化 + 前後空白除去
    s = str(x).strip()
    # Unicode正規化（全角→半角）
    s = unicodedata.normalize("NFKC", s)
    # 純数字なら0埋め
    if s.isdigit():             # pylint: disable=no-else-return
        if len(s) >= width:     # pylint: disable=no-else-return
            print(f"警告: コードが {width} 文字以上あります: {s}", file=sys.stderr)
            return s
        else:
            return s.zfill(width)
    else:
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
    if is_real_nan(x):
        return x
    # 文字列化 + 前後空白除去
    s = str(x)
    # Unicode正規化（全角→半角）
    s = unicodedata.normalize("NFKC", s)
    s = s.strip()
    # 純数字なら0埋め
    m = re.match(r"(?P<n>\d{1,10})(@utac\.u\-tokyo\.ac\.jp)?$", s)
    if m:                       # pylint: disable=no-else-return
        n = m.group("n")
        n = "0" * (10 - len(n)) + n
        return n
    else:
        return None             # error

def validate_program_students(url, sheet=0):
    """
    プログラムの認定対象科目を記述した表の読み込み
    共通ID カラムが必須
    """
    df = open_sheet(url, sheet)
    df = normalize_columns(df)
    cols = ["共通ID"] # "学籍番号", "学生氏名"
    check_columns(df, cols)
    df["共通ID_正規化"] = df["共通ID"].apply(normalize_utac)
    df = df.add_suffix("_登録学生一覧")
    return df

def validate_utas_grade(url, sheet=0):
    """
    UTASの成績を読み込む
    """
    df = open_sheet(url, sheet)
    df = normalize_columns(df)
    cols = ["共通ID", "年度", "科目コード", "合否区分", "単位数"]
    opt_cols = ["学籍番号", "学生氏名", "学生氏名カナ", "学生所属"]
    for c in opt_cols:
        ensure_column(df, c, "")
    if 0:                       # pylint: disable=using-constant-test
        cols = ["共通ID", "学籍番号", "学生氏名", "学生氏名カナ", "学生所属",
                "年度", "時間割所属", "時間割コード", "開講科目名",
                "科目コード", "科目名", "主担当教員名", "主担当教員共通ID",
                "開講区分名", "合否区分", "単位数"]
    df = check_columns(df, cols + opt_cols)
    df["共通ID_正規化"] = df["共通ID"].apply(normalize_utac)
    df["科目コード_正規化"] = df["科目コード"].apply(normalize_code)
    df = df.add_suffix("_UTAS")
    return df

def make_year_list(df):
    """
    df["対象年度"]の表記から, 対象年度リスト (整数のリスト) にする
    """
    # 開講年度の文字列をリスト化
    years = df["対象年度"].apply(parse_num_set)
    err_rows = years == "ERROR"
    # エラーがあればエラーを表示してNone
    if any(err_rows):
        for i, row in df[err_rows].iterrows():
            year = row["対象年度"]
            print(f'エラー: {i+1} 行目の対象年度 "{year}" が不正', file=sys.stderr)
        return None
    return years

def validate_program_courses(url, sheet=0):
    """
    科目一覧の読み込み
    """
    df = open_sheet(url, sheet)
    df = normalize_columns(df)
    # 「対象年度」列自体がなかったら全部空 (=全部毎年) とみなす
    cols = ["科目コード", "対象年度"] # "科目名"
    opt_cols = ["対象年度", "組み合わせ科目群", "排他科目群"]
    for c in opt_cols:
        ensure_column(df, c, "")
    df = check_columns(df, cols + opt_cols)
    df["科目コード_正規化"] = df["科目コード"].apply(normalize_code)
    years = make_year_list(df)
    if years is None:
        return None
    df["対象年度リスト"] = years
    df = df.add_suffix("_科目一覧")
    return df

def year_in_year_list(row):
    """
    row : 行
    年度_UTAS が 対象年度リスト_科目一覧 に入っているかどうか?
    """
    years = row["対象年度リスト_科目一覧"]
    year = row["年度_UTAS"]
    if is_real_nan(years):
        return True
    assert(isinstance(years, type([]))), years
    return year in years

def cat_unique(s):
    """
    s : Series.
    NaNは無視. それ以外は異なるものだけ取り出してつなげる
    """
    return "\n".join([str(x) for x in pd.unique(s.dropna())])

def merge_sets(s):
    """
    s : Series.
    各要素はリストもしくは NaN (全集合)
    """
    a = set()
    for x in s:
        if is_real_nan(x):
            return x
        a.update(x)
    return sorted(list(a))

def group_and_concat_uniques(df, group_cols):
    """
    df を C でグループ化し, 残りのカラムは重複を削除し, 残ったものを \n で連結
    """
    cols = [c for c in df.columns if c not in group_cols]
    agg_spec = {c : (c, cat_unique) for c in cols}
    df = df.groupby(group_cols, as_index=False, dropna=False).agg(**agg_spec)
    return df

def group_and_merge_years(df, group_cols, merge_cols):
    """
    df を C でグループ化し, 残りのカラムは重複を削除し, 残ったものを \n で連結
    """
    cat_cols = [c for c in df.columns if c not in (merge_cols + group_cols)]
    cat_spec   = {c : (c, cat_unique) for c in cat_cols}
    merge_spec = {c : (c, merge_sets) for c in merge_cols}
    agg_spec = cat_spec | merge_spec
    df = df.groupby(group_cols, as_index=False, dropna=False).agg(**agg_spec)
    return df

def do_check(program_students, utas_grade, program_courses,
             req_credits):
    """
    program_students : dataframe
    utas_grade : dataframe
    program_courses : dataframe
    req_credits : 認定に必要なクレジット
    """
    # pylint: disable=too-many-locals
    ps = program_students
    ug = utas_grade
    pc = program_courses
    # 同一人物(共通ID)が異なる行に存在している(修士と博士両方で登録など)ケースをマージ
    ps = group_and_concat_uniques(ps, ["共通ID_正規化_登録学生一覧"])
    # 同一科目コードで対象年度が複数行で書かれているケースをマージ
    pc = group_and_merge_years(pc,
                               ["科目コード_正規化_科目一覧"],
                               ["対象年度リスト_科目一覧"])
    # 登録学生一覧にUTAS成績をjoin
    ps_ug = pd.merge(ps, ug, how="left",
                     left_on="共通ID_正規化_登録学生一覧",
                     right_on="共通ID_正規化_UTAS")
    # それに認定科目一覧をjoin
    df = pd.merge(ps_ug, pc, how="left",
                  left_on="科目コード_正規化_UTAS",
                  right_on="科目コード_正規化_科目一覧")
    df["認定対象"] = np.where(df.apply(year_in_year_list, axis=1), 1, np.nan)
    df["認定対象かつ合格"] = df["認定対象"].where(df["合否区分_UTAS"] == "合格")
    # 複数の科目を組み合わせると初めて認定されるケースの処理
    # 組み合わせ科目群_科目一覧 に同じ値が入っている科目群すべてが
    # 認定対象かつ合格=1 だったらそれらすべてが認定される
    # ただし 組み合わせ科目群_科目一覧 が空 ("") だったものはそれぞれ単独で
    # 科目群とみなされる
    #
    # 組み合わせ科目群_科目一覧 が空でないものとそうでないものに分ける
    df["組み合わせ科目群_KEY"] = np.where((df["組み合わせ科目群_科目一覧"] == "") | (df["組み合わせ科目群_科目一覧"].isnull()),
                                          "i" + df.index.astype("string"),
                                          "k" + df["組み合わせ科目群_科目一覧"].fillna("").astype("string"))
    # 学生, 組み合わせ_given, 組み合わせでグループ化
    # 認定対象かつ合格にひとつでも na があるかを見る
    combined = (df.groupby(["共通ID_正規化_登録学生一覧",
                            "組み合わせ科目群_KEY"])["認定対象かつ合格"]
                .transform(lambda x: x.isna().any()))
    df["組み合わせ認定"] = df["認定対象かつ合格"].mask(combined)
    # 排他的科目 (それらからいくつとっても一つと数えられる)ものの処理
    # 排他科目群_科目一覧 が同じ値の科目群はその中から一つしか認定されない
    df["排他科目群_KEY"] = np.where((df["排他科目群_科目一覧"] == "") | (df["排他科目群_科目一覧"].isnull()),
                                    "i" + df.index.astype("string"),
                                    "k" + df["排他科目群_科目一覧"].fillna("").astype("string"))
    mutexed = (df.groupby(["共通ID_正規化_登録学生一覧",
                           "排他科目群_KEY"])["組み合わせ認定"]
               .idxmax().dropna())
    df["認定"] = np.nan
    df.loc[mutexed, "認定"] = 1
    df["認定単位"] = df["認定"] * df["単位数_UTAS"]
    # join key
    jkeys = ["共通ID_正規化_登録学生一覧"]
    ps_columns = [c for c in ps.columns if c not in jkeys]
    ug_columns = ["学籍番号_UTAS", "学生氏名_UTAS", "学生氏名カナ_UTAS", "学生所属_UTAS"]
    agg_spec = { c : (c, cat_unique) for c in ps_columns + ug_columns }
    agg_spec.update({"認定単位" : ("認定単位", "sum")})
    credit = df.groupby(jkeys, dropna=False).agg(**agg_spec)
    credit["認定結果"] = np.where(credit["認定単位"] >= req_credits, 1, np.nan)
    result_xlsx = "認定単位.xlsx"
    with pd.ExcelWriter(result_xlsx) as writer:
        credit.to_excel(writer, sheet_name="認定単位")
        df[["共通ID_正規化_登録学生一覧","学生氏名_UTAS","年度_UTAS","科目コード_UTAS","科目名_UTAS","合否区分_UTAS","単位数_UTAS","対象年度_科目一覧","認定対象","認定対象かつ合格","組み合わせ科目群_KEY","組み合わせ認定","排他科目群_KEY","認定","認定単位"]].to_excel(writer, sheet_name="詳細", index=False)
        df.to_excel(writer, sheet_name="全て", index=False)
    return (credit, result_xlsx)

def main():
    """
    main function
    """
    program_students_url = "登録学生一覧.xlsx"
    program_students_sheet = 0
    program_students = validate_program_students(program_students_url, sheet=program_students_sheet)
    assert program_students is not None
    utas_grade_url = "utas_grade.xlsx"
    utas_grade_sheet = 0
    utas_grade = validate_utas_grade(utas_grade_url, sheet=utas_grade_sheet)
    assert utas_grade is not None
    program_courses_url = "科目一覧.xlsx"
    program_courses_sheet = 0
    program_courses = validate_program_courses(program_courses_url, sheet=program_courses_sheet)
    assert program_courses is not None
    required_credits = 3
    do_check(program_students, utas_grade, program_courses, required_credits)

def test():
    """
    test corner cases
    """
    url = "test.xlsx"
    ps = validate_program_students(url, sheet=0)
    assert ps is not None
    ug = validate_utas_grade(url, sheet=1)
    assert ug is not None
    pc = validate_program_courses(url, sheet=2)
    assert pc is not None
    do_check(ps, ug, pc, 3)

if 0 or __name__ == "__main__": # pylint: disable=simplifiable-condition
    main()
    # test()
