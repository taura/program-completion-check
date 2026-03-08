# program-completion-check

## これは何?

* WINGSなどのプログラムの修了を判定するツール
* [![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/taura/program-completion-check/blob/main/src/check.ipynb) https://colab.research.google.com/github/taura/program-completion-check/blob/main/src/check.ipynb を踏むと Google Colab という作業環境のページが開かれる
* あとは指示のとおりに実行すれば良いはず
* 個人ごとに別の環境になるので他の人に影響するなどの心配は不要
* 説明動画: https://youtu.be/6ajm5mGJM8Y

## 準備すべきデータ

1. 本部学務課から共有される, UTASから抽出した全学生の成績ファイル
1. プログラム担当の方で用意する, プログラム登録者一覧
1. プログラム担当の方で用意する, プログラムの提供科目一覧

* なお形式は, ExcelもしくはGoogleのスプレッドシート
  * Excelはcolab環境に直接アップしたものでもGoogleドライブにアップしたものでも可

## データを用意せずにダミー実行

* データを用意しなくても, ダミーデータを使って上記をそのまま実行できる
* 参考までにダミーデータはこちら
  * UTASデータ https://drive.google.com/drive/folders/16vqDsIk_WNCPDMuUNtoT2nkCiEbIQyYU?usp=drive_link
  * プログラム登録者一覧および提供科目一覧 https://drive.google.com/drive/folders/1slHCnVXp7NOhNSPQs2BRossUBBSV9ZqH?usp=drive_link
* 自分のデータを用意する際もこれをサンプルとして見ていただけると良いです

## 変更記録

* 2026/03/04
  * 開講年度指定の柔軟化 
	* 「開始年度」, 「終了年度」の指定を廃止して, 「開講年度」とし, そこに, "2022", "2021-2025", "2021, 2023", "2021, 2023-2025" などの記法を許す. 何も書かなければ無限過去から無限未来開講の意味. 「開講年度」がなければすべてその意味
  * 登録学生一覧とUTASを照合する際のキーとして「共通ID」というカラムを用いる(IDは全角でも可). 同じ人物(共通ID)でも修士と博士で学籍番号が違うことに対応
  * 登録学生一覧にあるカラムを全て出力に含める (カラム名にはすべて "_登録学生一覧" が最後につく)
  * 科目一覧にある余分なカラムを全て「詳細」出力に含める (カラム名には全て "_科目一覧" が最後につく)
* 2026/03/05
  * 「開講年度」という言葉が誤解を招く可能性があるため「対象年度」に変更 (プログラムで認定する対象の年度という意味)

* 2026/03/07
  * 組み合わせ科目群 (同じ群に所属する科目を全て取ると初めて認定される科目群) をサポート
  * 排他科目群 (同じ群に所属する科目をいくつとっても一つしか認定されない科目群) をサポート
