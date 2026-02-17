# program-completion-check

## これは何?

* WINGSなどのプログラムの修了を判定するツール
* [![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/taura/program-completion-check/blob/main/src/check.ipynb) https://colab.research.google.com/github/taura/program-completion-check/blob/main/src/check.ipynb を踏むと Google Colab という作業環境のページが開かれる
* あとは指示のとおりに実行すれば良いはず
* 個人ごとに別の環境になるので他の人に影響するなどの心配は不要

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

