""" md
#* プログラム修了認定アプリ

* 基本はファイル名の設定を3箇所行って SHIFT + ENTER で実行していくだけです
* 指示を読みながら, SHIFT + ENTER もしくは各セルの ▶ をクリックして実行を続けてください
* 慣れてきたら  ▷ Run all ですべて一撃で実行してもOK
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
""" """

""" md
# 登録学生一覧の設定

* 当該プログラムに登録している学生のリスト (GoogleドライブのExcel, GoogleドライブのGoogleスプレッドシート, colabにアップしたExcelのいずれか) を設定
* プログラム担当者が作る (おそらく始めからどこかに存在している) 想定

* **以下の `PROGRAM_STUDENTS_URL = "..."` を自分が使いたいものに変更してください**
* 以下はこちらで用意した誰でも読めるものになっているので試しにそのまま実行してみてもOK
* 初めて実行する時に窓が出てきて "Allow this notebook to access your Google credentials?" と聞かれるので許可 "Allow" して下さい
* 別の小窓が出てきて「Third-party authored notebook code に再ログインしようとしています」「Third-party authored notebook code が Google アカウントへのアクセスを求めています」みたいなことを聞かれるので「次へ」や「続行」を押して先へ進んで下さい (ボタンを押すためにスクロールする必要があるかも知れません)

* **<font color=red>注意 (重要)</font>**
* `PROGRAM_STUDENTS_URL` は以下のいずれかです
  * Googleドライブ上のExcelのURL
  * Googleドライブ上のGoogleスプレッドシートのURL
  * colabにアップロードされたファイル名
* Googleドライブ上のファイルURLの取得方法: フォルダ上でそのファイルを右クリックして「共有」 -> 「リンクをコピー」するか, ファイルを開いた状態で「共有」の右側の▼ -> 「リンクをコピー」でURLを取得します
  * URL末尾の `edit?usp=sharing` みたいな部分はあってもなくてもOK (多分)
* colabにファイルをアップロードする方法
  * 左側のアイコンからフォルダのアイコンを選んでフォルダ内容を表示し, 上にあるアップロードボタンを押す

* `PROGRAM_STUDENTS_SHEET` はシート番号もしくはシートの名前
  * 先頭のシートであれば0
  * 特にシートが一枚しか無ければ0でOK
* 以下をSHIFT + ENTERで実行するとシートの読み込みと必須フィールドのチェックが行われる
* 以上の注意は以下も同様

* 必須フィールド (以下以外に余分なフィールドがあっても構いません)
  * **学籍番号** : 形式は UTAS の学籍番号の形式に合わせる (48196459 みたいな)
  * **学生氏名** : 処理には使われないが結果の目視を容易にするため結果には含まれる

* TODO
  * 学籍番号の表記ゆれに対応したい (そうする予定)
  * 目視のため結果に含めたいフィールドを追加で指定できるようにしたい (そうする予定)

"""

""" code kernel=python """
# 本番実行時は "..." を書き換えて設定してください
PROGRAM_STUDENTS_URL = "https://docs.google.com/spreadsheets/d/1Abeszk5-iJokwxGV6W115uUX5Bq2lziB/"
# colabに直接アップした場合はこちら
#PROGRAM_STUDENTS_URL = "登録学生一覧.xlsx"
PROGRAM_STUDENTS_SHEET = 0

PROGRAM_STUDENTS = pcc.validate_program_students(PROGRAM_STUDENTS_URL, sheet=PROGRAM_STUDENTS_SHEET)
print(f"{PROGRAM_STUDENTS.size} 行あります (先頭5行は以下です)")
PROGRAM_STUDENTS.head(5)
""" """

""" md
# UTAS から抽出した成績表の設定

* これは本部学務課から送られてくる, なにもいじらなくてよい想定
* `UTAS_GRADE_URL` に関する注意は上記を参照

* 必須フィールド (以下以外に余分なフィールドがあっても構いません)
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
UTAS_GRADE_URL = "https://docs.google.com/spreadsheets/d/1Qq31omlr3QT2kjPzzxEjcIfbWLTTDZzx/"
# colabに直接アップした場合はこちら
#UTAS_GRADE_URL = "utas_grade.xlsx"
UTAS_GRADE_SHEET = 0

UTAS_GRADE = pcc.validate_utas_grade(UTAS_GRADE_URL, sheet=UTAS_GRADE_SHEET)
print(f"{UTAS_GRADE.size} 行あります (先頭5行は以下です)")
UTAS_GRADE.head(5)
""" """

""" md
# 科目一覧の設定

* 当該プログラムの科目一覧 (Googleドライブ上のExcelまたはスプレッドシート) を設定
* プログラム担当者が作る想定 (おそらく始めからそのようなファイルがある)

* 必須フィールド (以下以外に余分なフィールドがあっても構いません)
  * **科目コード (形式は UTAS の科目コードと合わせる)**
  * **科目名**
  * **開始年度**
  * **終了年度**

* TODO
  * 科目コードの表記ゆれに対応したい (そうする予定)
  * 開始から終了だけでなく年度ごとに指定する(2023, 2025のような)パターンに対応したい (そうする予定)

"""

""" code kernel=python """
# 本番実行時は "..." を書き換えて設定してください
PROGRAM_COURSES_URL = "https://docs.google.com/spreadsheets/d/1NIOPkmqLLK6MgkvL3NMcGL1u2KDK8Ykv/"
# colabに直接アップした場合はこちら
#PROGRAM_COURSES_URL = "科目一覧.xlsx"
PROGRAM_COURSES_SHEET = 0

PROGRAM_COURSES = pcc.validate_program_courses(PROGRAM_COURSES_URL, sheet=PROGRAM_COURSES_SHEET)
print(f"{PROGRAM_COURSES.size} 行あります (先頭5行は以下です)")
PROGRAM_COURSES.head(5)
""" """

""" md
# 実際の修了認定チェック!

* 詳しい結果は「認定単位.xlsx」というファイルに出力されます
* 左にあるメニューからフォルダアイコンを選んでファイルを表示してダウンロードしてください

"""

""" code kernel=python """
CREDIT, RESULT_XLSX = pcc.do_check(PROGRAM_STUDENTS, UTAS_GRADE, PROGRAM_COURSES)
print(f"{CREDIT.size} 行あります. 結果を {RESULT_XLSX} に保存しました")
CREDIT
""" """
