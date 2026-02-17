""" md
#* プログラム修了認定アプリ

* 基本はファイル名の設定を3箇所行って SHIFT + ENTER で実行していくだけです
* 指示を読みながら, SHIFT + ENTER で実行を続けてください
"""

""" md
# 実行前の事前準備

* 以下を初めて実行 (SHIFT + ENTER) すると
  * Googleに login していなければ求められるので "Sign in" を押して  g.ecc アカウントでサインインしてください
  * Sign inをしたらもう一度SHIFT + ENTERで実行してください
* 実行すると「This notebook was not authored by Google」みたいな警告がでてきます. "Run Anyway" を押して先へ進んでください
* ソフトのインストールに数十秒かかります

"""

""" code kernel=python """
!pip install -q git+https://github.com/taura/program-completion-check
""" """

""" md
* 以下を初めて実行すると, 小さな窓が出てきます
* 窓が出てきたらスクロールして最後の部分を表示して「許可」のボタンをおして先へ進んでください
* その後, wings-credit-count というアプリケーションへ Google Drive へのアクセスを許してよいかを聞いてくるので全て許可を選んで許可してください
"""

""" code kernel=python """
import pandas as pd
import program_completion_check as pcc
pd.set_option("display.max_rows", None)

try:
    GC
except:
    GC = pcc.get_client()
""" """

""" md
# 当該プログラムに登録している学生のリスト (Google spreadsheet) を設定

* プログラム担当者が作る (おそらく始めからどこかに存在している) 想定

* **以下の PROGRAM_STUDENTS_URL = "..." を自分が使いたいものに変更してください**
* 以下はこちらで用意した誰でも読めるものになっているので試しにそのまま実行してみてもOK

* **<font color=red>注意 (重要)</font>**
* PROGRAM_STUDENTS_URL は Google spreadsheet の URL です
  * 注: **アップロードした Excel そのままでは, 開けません (そのうち修正したい・その予定)**
  * アップロードしたExcel をGoogle spreadsheet で開けるようにするにはその Excel を一旦Google Drive上で開き, File -> 「Google spreadsheetとして保存する」を実行
  * Google Drive上で拡張子 (.xlsx) がなくなって表示されるのが Google spreadsheet です
  * それが並んでいるフォルダ上でそのファイルを右クリックして「共有」 -> 「リンクをコピー」するかファイルを開いた状態で「共有」の右側の▼ -> 「リンクをコピー」でURLを取得します
  * URL末尾の `edit?usp=sharing` みたいな部分はなくてもOK (多分)

* PROGRAM_STUDENTS_SHEET はGoogle spreadsheet中のシート番号 (0から始まる, 何枚目かを表す番号) もしくはシートの名前
  * シートが一枚しか無ければ0のままでOK
* 以下をSHIFT + ENTERで実行するとシートの読み込みと必須フィールドのチェックが行われる
* 以上の注意は以下も同様

* 必須フィールド
  * **学籍番号** : 形式は UTAS の学籍番号の形式に合わせる (48196459 みたいな)

"""

""" code kernel=python """
# 本番実行時は "..." を書き換えて設定してください
PROGRAM_STUDENTS_URL = "https://docs.google.com/spreadsheets/d/1fxk_V9YZMvLAWk9N-g3NU5ygMfWZ4WcveZ3yH9jh-PU/"
PROGRAM_STUDENTS_SHEET = 0

PROGRAM_STUDENTS = pcc.validate_program_students(GC, PROGRAM_STUDENTS_URL, sheet=PROGRAM_STUDENTS_SHEET)
print(f"{PROGRAM_STUDENTS.size} 行あります (先頭5行は以下です)")
PROGRAM_STUDENTS.head(5)
""" """

""" md
# UTAS から抽出した成績表 (Google spreadsheet) を設定

* これは本部学務課から送られてくる想定
* UTAS_GRADE_URL に関する注意は上記を参照

* 必須フィールド:
  * **共通ＩＤ**
  * **学籍番号**
  * **学生氏名**
  * **学生氏名カナ**
  * **学生所属**
  * **年度**
  * **時間割所属**
  * **時間割コード**
  * **開講科目名**
  * **科目コード**
  * **科目名**
  * **主担当教員名**
  * **主担当教員共通ID**
  * **開講区分名**
  * **合否区分**
  * **単位数**
"""

""" code kernel=python """
# 本番実行時は "..." を書き換えて設定してください
UTAS_GRADE_URL = "https://docs.google.com/spreadsheets/d/1HploSNLzKuvhJcIev6WCzwGzjZOikvADFLzzmP2kSJU/"
UTAS_GRADE_SHEET = 0

UTAS_GRADE = pcc.validate_utas_grade(GC, UTAS_GRADE_URL, sheet=UTAS_GRADE_SHEET)
print(f"{UTAS_GRADE.size} 行あります (先頭5行は以下です)")
UTAS_GRADE.head(5)
""" """

""" md
# 当該プログラムの科目表 (Google spreadsheet) を設定

* プログラム担当者が作る想定 (おそらく始めからそのようなファイルがある)

* 必須フィールド:
  * **科目コード (形式は UTAS の科目コードと合わせる)**
  * **開始年度**
  * **終了年度**
"""

""" code kernel=python """
# 本番実行時は "..." を書き換えて設定してください
PROGRAM_COURSES_URL = "https://docs.google.com/spreadsheets/d/1zIJGOZ6Lfxki4YoaDGG_UcBYhlmSxs2QVK5h0VmMpNw/"
PROGRAM_COURSES_SHEET = 0

PROGRAM_COURSES = pcc.validate_program_subjects(GC, PROGRAM_COURSES_URL, sheet=PROGRAM_COURSES_SHEET)
print(f"{PROGRAM_COURSES.size} 行あります (先頭5行は以下です)")
PROGRAM_COURSES.head(5)
""" """

""" md
# 実際の修了認定チェック!

* 詳しい結果は「認定単位.xlsx」というファイルに出力されるのでダウンロードしてください

"""

""" code kernel=python """
CREDIT, RESULT_XLSX = pcc.do_check(PROGRAM_STUDENTS, UTAS_GRADE, PROGRAM_COURSES)
print(f"{CREDIT.size} 行あります. 結果を {RESULT_XLSX} に保存しました")
CREDIT
""" """
