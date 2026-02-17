# program-completion-check

## これは何?

* WINGSなどのプログラムの修了を判定するツール
* https://colab.research.google.com/github/taura/program-completion-check/blob/main/src/check.ipynb を踏むと Google Colab という作業環境のページが開かれる
* あとは指示のとおりに実行すれば良いはず
* 個人ごとに別の環境になるので他の人に影響するなどの心配は不要

## 準備すべきデータ

1. 本部学務課から共有される, UTASから抽出した全学生の成績ファイル
1. プログラム担当の方で用意する, プログラム登録者のリスト
1. プログラム担当の方で用意する, プログラムに登録された科目のリスト

* なお形式は全てGoogleのスプレッドシート
* Excelで渡された場合はGoogle Driveにアップした上で開いて, File -> 「Google Spreadsheetとして保存」する
  * 今後, 直接Excelを開けるように改善予定
  * (独り言) どうせアップするならcolabにアップすれば良い説

## データを用意せずにダミー実行

* データを用意しなくても, ダミーデータを使って上記をそのまま実行できる
* 参考までにダミーデータはこちら https://drive.google.com/drive/folders/1slHCnVXp7NOhNSPQs2BRossUBBSV9ZqH?usp=drive_link
* 自分のデータを用意する際もサンプルとして利用可能

