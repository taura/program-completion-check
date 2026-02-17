"""
# プログラム修了認定アプリ

* 以下を実行するとブラウザに, Googleでのサインインを求めるページが現れるので, g.ecc アカウントでサインインしてください
* その後アプリケーション wings-credit-count が許可を求めてくるので許可してください
"""
import pandas as pd
import program_completion_check as pcc
pd.set_option("display.max_rows", None)

try:
    GC
except:
    GC = pcc.get_client()

"""
(1) 当該プログラムに登録している学生のリスト
* プログラム担当者が作る (おそらく始めからどこかに存在している)
* 必須フィールド
  * 学籍番号 : 形式は UTAS の学籍番号の形式に合わせる (48196459 みたいな)
* URL は Google Drive の Google Spreadsheet の URL
  * 注: **アップロードした Excel そのままでは現状開けない (修正したい・その予定)**
  * 開けるようにするには Excel をGoogle Drive上で開き, File -> Google Spreadsheetとして保存する
* SHEET は 番号 (0から始まる何枚目かを表す番号) もしくはシートの名前
* 以下をSHIFT + ENTERで実行するとシートの読み込みと必須フィールドのチェックが行われる

以下同様

"""
PROGRAM_STUDENTS_URL = "https://docs.google.com/spreadsheets/d/1fxk_V9YZMvLAWk9N-g3NU5ygMfWZ4WcveZ3yH9jh-PU/"
PROGRAM_STUDENTS_SHEET = 0

PROGRAM_STUDENTS = pcc.validate_program_students(GC, PROGRAM_STUDENTS_URL, sheet=PROGRAM_STUDENTS_SHEET)
print(f"{PROGRAM_STUDENTS.size} 行あります (先頭5行は以下です)")
PROGRAM_STUDENTS.head(5)

"""
(2) UTAS から抽出した成績表 (本部学務課から送られてくる想定)

必須フィールド:
* 共通ＩＤ
* 学籍番号
* 学生氏名
* 学生氏名カナ
* 学生所属
* 年度
* 時間割所属
* 時間割コード
* 開講科目名
* 科目コード
* 科目名
* 主担当教員名
* 主担当教員共通ID
* 開講区分名
* 合否区分
* 単位数
"""
UTAS_GRADE_URL = "https://docs.google.com/spreadsheets/d/1HploSNLzKuvhJcIev6WCzwGzjZOikvADFLzzmP2kSJU/"
UTAS_GRADE_SHEET = 0

UTAS_GRADE = pcc.validate_utas_grade(GC, UTAS_GRADE_URL, sheet=UTAS_GRADE_SHEET)
print(f"{UTAS_GRADE.size} 行あります (先頭5行は以下です)")
UTAS_GRADE.head(5)

"""
(3) 科目表 = 当該プログラムの科目表

必須フィールド:
* 科目コード (形式は UTAS の科目コードと合わせる)
* 開始年度
* 終了年度
"""
PROGRAM_COURSES_URL = "https://docs.google.com/spreadsheets/d/1zIJGOZ6Lfxki4YoaDGG_UcBYhlmSxs2QVK5h0VmMpNw/"
PROGRAM_COURSES_SHEET = 0

PROGRAM_COURSES = pcc.validate_program_subjects(GC, PROGRAM_COURSES_URL, sheet=PROGRAM_COURSES_SHEET)
print(f"{PROGRAM_COURSES.size} 行あります (先頭5行は以下です)")
PROGRAM_COURSES.head(5)

"""
# 実際の修了認定チェック!

* 詳しい結果は「認定単位.xlsx」というファイルに出力されるのでダウンロードしてください

"""

CREDIT, RESULT_XLSX = pcc.do_check(PROGRAM_STUDENTS, UTAS_GRADE, PROGRAM_COURSES)
print(f"{CREDIT.size} 行あります. 結果を {RESULT_XLSX} に保存しました")
CREDIT

