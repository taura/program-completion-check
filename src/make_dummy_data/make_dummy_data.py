#!/usr/bin/env python

import pandas as pd
import numpy as np
import random

COURSE_NAMES = [
    "ほげほげ概論", 
    "ガチャガチャ特論", 
    "うっかり工学入門", 
    "だいたい理解学", 
    "なんとなく統計学", 
    "眠気制御理論", 
    "締切直前行動科学", 
    "コーヒー依存症研究", 
    "雑談力応用演習", 
    "先延ばし最適化論", 
    "気合と根性の数理", 
    "やる気探索アルゴリズム", 
    "昼寝計画学", 
    "レポート生成錬金術", 
    "空気読み力学", 
    "微妙な沈黙論", 
    "既読スルー情報学", 
    "雰囲気でわかる量子力学", 
    "とりあえず保存学", 
    "バズり現象解析", 
    "深夜テンション工学", 
    "無限言い訳生成法", 
    "会議短縮失敗学", 
    "目的を見失う設計論", 
    "適当モデリング入門", 
    "思いつき実験学", 
    "忘却曲線活用術", 
    "雑務自動増殖論", 
    "永久仮置きファイル学", 
    "なんとかなる確率論", 
    "直感だけで解く微積分", 
    "ゆるふわプログラミング", 
    "謎仕様解読演習", 
    "エラー文鑑賞学", 
    "バグとの対話実践", 
    "再起動信仰論", 
    "環境構築永久入門", 
    "依存関係地獄学", 
    "コピペ高度応用", 
    "コメントアウト文化史", 
    "未完成美学特講", 
    "締切延長交渉術", 
    "それっぽい図表作成法", 
    "雰囲気プレゼン実習", 
    "拍手誘導デザイン", 
    "質疑応答回避技法", 
    "長い前置き研究", 
    "結論後回し学", 
    "タイトル詐欺概論", 
    "体裁だけ整える技術", 
    "根拠薄弱推論学", 
    "多分大丈夫工学", 
    "見なかったことにする法", 
    "後でやる管理論", 
    "優先順位混乱学", 
    "時間が溶ける情報科学", 
    "気づいたら夜明け学", 
    "謎の自信形成論", 
    "突然のやる気論", 
    "三日坊主持続可能性", 
    "計画倒れデザイン", 
    "積読消滅学", 
    "買って満足経済学", 
    "無料だけ使う戦略論", 
    "パスワード忘却史", 
    "二度手間最適化", 
    "ショートカット過信学", 
    "設定迷子論", 
    "説明書未読文化", 
    "なんでも再利用工学", 
    "仮フォルダ恒久保存論", 
    "最新版追跡不能学", 
    "互換性崩壊史", 
    "突然動かなくなる学", 
    "昨日まで動いてたのに論", 
    "原因不明現象学", 
    "とりあえず様子見学", 
    "誰かがやるだろう理論", 
    "責任所在拡散学", 
    "雰囲気合意形成論", 
    "多数決万能説批判", 
    "微妙な空気維持学", 
    "雑なまとめ方特講", 
    "それ後で送ります論", 
    "CC入れすぎ問題研究", 
    "返信タイミング最適化", 
    "了解です学概論", 
    "お世話になっております論", 
    "よろしくお願いします学", 
    "差し戻し芸術論", 
    "永久ドラフト保存学", 
    "最終版（仮）研究", 
    "本当に最終版研究", 
    "最終版ver2研究", 
    "最終版fix最終研究", 
    "もう触らない宣言学", 
    "静かにフェードアウト論", 
    "なかったことにする歴史", 
    "気づかなかったふり実践", 
    "いい感じに終わらせる学", 
]

STUDENT_NAMES = [
    "いしはり さとみん",
    "きたぐわ けいこちゃ",
    "わおたに しょうたいむ",
    "あらがち ゆいりん",
    "ながさま まさみん",
    "あやせる はるかるん",
    "はまなべ みなみゅ",
    "ひろしぇ すずたん",
    "かわぐてぃ はるにゃ",
    "もりなな ぴょん",
    "きむたくん たくやん",
    "すだまちゃ まさきゅ",
    "よしざわ りょうち",
    "やまじゃき けんとん",
    "にしじま ひでとん",
    "たけうてぃ りょうまる",
    "さかぐてぃ けんたろぴ",
    "ふくだん みらきゅ",
    "おおいずみ ひろしぇ",
    "ほしのげ げんち",
    "だるびしゅ ゆうにゃ",
    "すずきせ いやたん",
    "むらかみ むねきゅん",
    "よしだま さたろ",
    "さかもとは やとりん",
    "やまだて つとん",
    "ほんだけ いすけちゃ",
    "ながとも ゆうぴょ",
    "みうらか ずたん",
    "まついひ できゅ",
    "おおさか なみりん",
    "にしこり けいたん",
    "はにゅう ゆづりゅ",
    "たかなし さらら",
    "いけえ りかこぴ",
    "とむるん くるーちゃ",
    "でぃかぷり おれおん",
    "ぶらぴと ぴっとん",
    "えまま そんりん",
    "ていすう ふとちゃ",
    "めしりお ねるん",
    "ろなくり どるん",
    "ねいまる るぴょ",
    "べかむで びちゃ",
    "もどりち るかりん",
    "ぼるうさ いんたん",
    "ふぇでら ろじゃりん",
    "なだる らふぁみゅ",
    "かりすて ふりん",
    "じょだま いけるん",
    "あいしゅ たいんるん",
    "にゅとーん いさりん",
    "だーうぃん ちゃるぴょ",
    "きゅりまり りんぬ",
    "てすらん にこぴょ",
    "えじそと とまち",
    "がりれお るんた",
    "こぺるに くすみゅ",
    "ぱすてる るんぴ",
    "めんでれ ふぃん",
    "もーつぁ るとぴ",
    "べーとー べんちゃ",
    "ばっはん るんみ",
    "しょぱん ぴょり",
    "りすとん るみゃ",
    "ぶらーむ すりん",
    "わーぐな るんちゃ",
    "まーらー るぴょ",
    "どびゅっし ゅりん",
    "らべるん みゃ",
    "えりっく さてぃん",
    "びよんせ るんちゃ",
    "あでるる ぴょみ",
    "ていらー すうぃりん",
    "りあーな ぴょんぬ",
    "ががれで ぃみゅ",
    "けいしー ましゅりん",
    "ありあな ぐらぴょ",
    "ぶるーの まるりん",
    "じゃすてぃ んびりゅ",
    "しょーん めんりん",
    "えどしー らんぴ",
    "こーるど ぷれりん",
    "はりすた いるんぬ",
    "えま すとんり",
    "じぇにふぁ ろれみゅ",
    "すかーれ っとりん",
    "なただり りんぴょ",
    "あんじぇり なりゅ",
    "じゅりあ ろばりん",
    "さんどら ぶるぴ",
    "めりる すとりゃ",
    "にこーる きどりん",
    "けいと うぃんり",
    "あん ねはさりん",
    "える ふぁにんぬ",
    "だこた じょんり",
    "みりー ぼびりん",
    "しどに すうぇり",
    "じぇな おてりん",
    "ふろーれ んすぴ",
    "おりびあ ころりん",
    "まーご っとりん",
    "じぇし かちゃり",
    "えみりー ぶらりん",
    "はりー すたぴょ",
    "ぜんで いやりん",
    "とむ ほらりん",
    "てぃも しゃらりん",
    "ろばー とぱりん",
    "くりす へむりん",
    "べねでぃ くとりん",
    "えでぃ れどりん",
    "らいあん ごすりん",
    "じぇいく ぎれりん",
    "ひゅー じゃくりん",
    "だにえる くらりん",
    "きあぬ りぶりん",
    "うぃる すみりん",
    "でんぜる わしりん",
    "まっと でいりん",
    "べん あふりん",
    "じょーじ くるりん",
    "ぶらっど ぴょりん",
    "れおなる どりん",
    "まーく らふりん",
    "とびー まぐりん",
    "あんどりゅ りんふ",
    "へんりー かびりん",
    "ろばーと ぱてりん",
    "でびっど べかりん",
    "くりろな うどりん",
    "めしりお ぴょりん",
    "ねいまー るりん",
    "むばっぺ きゅりん",
    "はーらん どりん",
    "もどりち るりん",
    "くろーす りんぬ",
    "いにえす たりん",
    "しゃび りんみゅ",
    "ぺどり りんぴ",
    "がび りんちゃ",
    "すあれす りんぴょ",
    "べんぜま りんる",
    "ればんど ふすりん",
    "まね りんぬ",
    "さらー りんみゅ",
    "そんふん みんりん",
    "みとま りんぴょ",
    "くぼ りんた",
    "えんどう りんみゅ",
    "とみやす りんぬ",
    "よしだ りんちゃ",
    "たき りんぴ",
    "かま りんみ",
    "おかざき りんぬ",
]

FACULTY_NAMES = [
    "とよとみ ひでおす",
    "おだのぶにゃが",
    "とくがわ おいでやす",
    "みなもとの よしつねりん",
    "たいらの きよもりん",
    "しょうとく たいしん",
    "さいごう たかもら",
    "さかもと りょうまん",
    "いたがき いたいすけ",
    "ふくざわ ゆきぽよ",
    "なぽれおん ぼなぱるとん",
    "しーざー じゅりあな",
    "くれおぱとろん",
    "あきれさんだろす だいおう",
    "はんにばるん",
    "ちんぎす ふーん",
    "まるこ ぽろろん",
    "こぺるんにくすん",
    "がーりがりがりれいくん",
    "にゅーとんとん",
    "だーうぃん ちゃー",
    "あいんしゅたいん あるべるとん",
    "まっくす ぷらんくとん",
    "きゅーりー まりこ",
    "ぱすてぅーるん",
    "めんでれーふん",
    "えじそん とますん",
    "てすらん にこらん",
    "らいとん ぶらざーず",
    "べるん あれきさんだー",
    "あぶらうりのりんかーん",
    "じょーじわしやねん",
    "るーずべるとん",
    "ちゃーちるん",
    "びすまるくん",
    "まるくす かーるん",
    "えんげるすん",
    "どつきえふすきー",
    "かみゅかみゅ",
    "ごるばちょふん",
    "こうしくん",
    "もうしくん",
    "もうしばせん",
    "そうそうそう",
    "りゅうびん",
    "そんけんじ",
    "しょかつりょうこ",
    "りーしん",
    "もっつぁると",
    "べんとーべん",
]    

SCHOOL_NAMES = [
    "なんくるないさ研究科",
    "名称未定（仮）研究科",
    "思てたのとちゃう研究科",
    "会議増える系研究科",
    "だれが責任取るの研究科",
    "誰もきーひん研究科",
    "まだ決まってない研究科",
    "そのうち考える研究科",
    "いつからこうなった研究科",
    "けっきょく元通り研究科",
    "資料どこ行った研究科",
    "前も同じことやった研究科",
    "今さら引き返せない研究科",
    "予算だけ先に消える研究科",
    "とくに目的はない研究科",
    "いい感じにお願いします研究科",
    "深い意味はない研究科",
    "気づいたら設立されてた研究科",
    "だいたい合ってる研究科",
    "最終的に誰も覚えてない研究科",
]

HOBBIES = [
    "家の周りを一周する",
    "コンビニの新商品チェック",
    "冷蔵庫の中を定期巡回する",
    "天気予報の答え合わせ",
    "エスカレーターで立ち位置を考える",
    "信号が変わるタイミング予想",
    "電車のドア位置研究",
    "コンビニコーヒーの濃さ比較",
    "机の上の整理（30分だけ）",
    "ペン回し（成功率30%）",
    "カフェの席取り戦略研究",
    "YouTubeおすすめの深掘り",
    "Googleマップで遠くの街を散歩",
    "Wi‑Fiの電波強度チェック",
    "エレベーターの待ち時間観測",
    "新しいボールペンの試し書き",
    "スーパーの値引きタイミング観察",
    "電車の発車メロディ研究",
    "お菓子のカロリー計算（食べた後）",
    "ノートの1ページ目をきれいに書く",
    "メールの件名を考えすぎる",
    "スマホのバッテリー節約術研究",
    "カップ麺の待ち時間アレンジ",
    "レシートの長さランキング",
    "自販機のラインナップ分析",
    "近所の猫の行動観察",
    "イヤホンの絡まり解消",
    "天井の模様を眺める",
    "電車の吊り広告を全部読む",
    "新しいショートカットキー探し",
    "机のガタつき調整",
    "雲の形の解釈",
    "早歩きと普通歩きの境界研究",
    "消しゴムの角を守る",
    "おにぎりの具ランキング更新",
    "Google検索の予測変換観察",
    "カレンダーの祝日配置チェック",
    "コーヒーを冷ましすぎる",
    "ペットボトルのラベル研究",
    "コンセントの空きを探す",
    "歩道のタイル模様を追う",
    "キーボードの掃除（気が向いたとき）",
    "書店で背表紙を読む",
    "電車の中で立ち位置最適化",
    "雨が降りそうかどうかの体感予測",
    "靴ひもを結び直すタイミング研究",
    "エアコンの最適温度探し",
    "目覚ましのスヌーズ回数管理",
    "スマホのホーム画面整理",
    "次にやることを考える",
]

OBJECTIVES = [
    "今年こそ朝型人間になる（予定）",
    "締切の前日に焦らない生活を目指す",
    "講義に5分前に着く人になる",
    "ノートをちゃんと取る（たぶん）",
    "シラバスをちゃんと読む",
    "レポートを提出期限より前に出す経験をする",
    "寝坊の理由を天候以外にする",
    "学食の新メニューを全部試す",
    "研究室のドアをノックする勇気を持つ",
    "積読の本を一冊は読む",
    "USBメモリをなくさない一年にする",
    "スマホを見ない時間を一日5分作る",
    "締切という概念と仲良くなる",
    "パスワードを覚えられるものにする",
    "机の上を週に一度は片付ける",
    "課題を提出してから安心して寝る",
    "講義中に眠くならない方法を研究する",
    "メールの返信を24時間以内にする",
    "システムのログインパスワードを忘れない",
    "研究室の椅子を壊さない",
    "プリンタの紙詰まりを自分で直せるようになる",
    "Wi-Fiが遅いときにまず自分を疑う",
    "発表スライドを前日に完成させる",
    "研究の進捗を月に一度は作る",
    "昼夜逆転を月に一度は直す",
    "「あとでやる」を減らす",
    "図書館の静けさに慣れる",
    "エレベーターを使わず階段を使う日を作る",
    "講義の質問を一度はしてみる",
    "パソコンのバックアップを取る",
    "机の引き出しの化石を整理する",
    "研究テーマを説明できるようになる",
    "Gitでやらかさない",
    "データを消してから後悔しない",
    "ToDoリストをToDoneリストにする",
    "朝ごはんを食べる日を増やす",
    "課題のファイル名をわかりやすくする",
    "机の周りのケーブルを整理する",
    "深夜のネットサーフィンを控える（理想）",
    "レポートの誤字を減らす",
    "プレゼンで「えー」を減らす",
    "コーヒーを飲みすぎない",
    "図書館で寝ない",
    "講義に出る回数を増やす",
    "勉強と休憩のバランスを取る",
    "友達の課題の進み具合を見て焦らない",
    "パソコンのアップデートを放置しない",
    "締切の存在を忘れない",
    "今年の抱負を来年も思い出す",
    "なんとか無事に一年を乗り切る",
]

MOTTOS = [
    "板書はゆっくり、消すのはもっとゆっくり",
    "わからない顔を見逃さない（たぶん）",
    "例え話でだいたい乗り切る",
    "難しい話ほど雑談から入る",
    "学生より少しだけ早起きする",
    "スライドは増えるが内容は削る",
    "一回くらいは笑いを取る",
    "沈黙が続いたら自分で答える",
    "質問がなくても気にしないふりをする",
    "脱線しても最後は戻る",
    "チョークを落としても動じない",
    "マイクのスイッチは先に確認",
    "時間通りに終わる努力はする",
    "毎年少しずつスライドが増える",
    "学生の表情から進度を推測する",
    "難しい式ほど「気持ち」を説明する",
    "だいたい一回は黒板を間違える",
    "教室の時計を信じすぎない",
    "静かな教室にも慣れる",
    "スライド番号を頼りに生きる",
    "質問されたらうれしい",
    "質問されなくても気にしない",
    "チョークの粉とは共存する",
    "例え話はその場で考える",
    "パソコンは授業前に再起動",
    "話がそれても授業の一部",
    "わからないと言われてもめげない",
    "学生の眠気とは長い付き合い",
    "最後の5分はだいたい足りない",
    "とりあえず今日もなんとかする"
]

def hiragana_to_katakana(s):
    return ''.join(
        chr(ord(c) + 0x60) if 'ぁ' <= c <= 'ゖ' else c
        for c in s
    )

def make_number(m, rg):
    return "".join([chr(ord('0') + rg.randint(0, 9)) for _ in range(m)])

def make_random_students(n, rg):
    """
    ID, 学生証番号, 名前, カナ, 所属研究科
    """
    # 重複チェック
    assert(len(STUDENT_NAMES) == len(set(STUDENT_NAMES))), (len(STUDENT_NAMES), len(set(STUDENT_NAMES)))
    student_names = STUDENT_NAMES[:]
    rg.shuffle(student_names)
    names = student_names[:n]
    katakana_names = [hiragana_to_katakana(s) for s in names]
    utacs = [ make_number(10, rg) for _ in range(n) ]
    sids0 = [ "8" + make_number(7, rg) for _ in range(n) ]
    sids1 = [ "9" + make_number(7, rg) for _ in range(n) ]
    sids = list(zip(sids0, sids1))
    schools = rg.choices(SCHOOL_NAMES, k=n)
    hobbies = rg.choices(HOBBIES, k=n)
    objectives = rg.choices(OBJECTIVES, k=n)
    assert(len(set(names)) == n)
    assert(len(set(sids0 + sids1)) == 2 * n)
    assert(len(set(utacs)) == n)
    l = sorted(list(zip(utacs, sids, names, katakana_names,
                        schools, hobbies, objectives)))
    return l
    
def make_random_faculties(n, rg):
    """
    ID, 名前 のリスト
    """
    assert(len(FACULTY_NAMES) == len(set(FACULTY_NAMES))), (len(FACULTY_NAMES), len(set(FACULTY_NAMES)))
    faculty_names = FACULTY_NAMES[:]
    rg.shuffle(faculty_names)
    names = faculty_names[:n]
    utacs = [ make_number(10, rg) for _ in range(n) ]
    assert(len(set(names)) == n)
    assert(len(set(utacs)) == n)
    return list(zip(utacs, names))

def make_random_subject_code(rg):
    """
    科目コード
    """
    d = "0123456789"
    if rg.random() < 0.6:
        chars = rg.choices(d, k=4) + ["-"] + rg.choices(d, k=4)
    else:
        chars = rg.choices(d, k=6)
    return "".join(chars)

def make_random_courses(n, faculties, rg):
    """
    n個のコースをランダムに生成
    時間割所属, 時間割コード, 開講科目名, 科目コード,
      科目名, 主担当教員名, 主担当教員共通ID
    """
    assert(len(COURSE_NAMES) == len(set(COURSE_NAMES))), (len(COURSE_NAMES), len(set(COURSE_NAMES)))
    assert(len(SCHOOL_NAMES) == len(set(SCHOOL_NAMES))), (len(SCHOOL_NAMES), len(set(SCHOOL_NAMES)))
    course_names = COURSE_NAMES[:]
    school_names = SCHOOL_NAMES[:]
    rg.shuffle(school_names)
    rg.shuffle(course_names)
    schools = school_names
    courses = course_names[:n]
    assert(len(set(courses)) == n)
    codes = [make_random_subject_code(rg) for _ in range(n)]
    faculty_names = [name for utac, name in faculties]
    faculty_utacs = [utac for utac, name in faculties]
    credit_candidates = [ 2, 2, 1, 0.5 ]
    credit = rg.choices(credit_candidates, k=n)
    term_candidates = ["通年", "S1S2", "A1A2"] * 2 + ["S1", "S2", "A1", "A2"]
    terms = rg.choices(term_candidates, k=n)
    mottos = rg.choices(MOTTOS, k=n)
    return list(zip(schools, codes, courses, codes,
                    courses, faculty_names, faculty_utacs,
                    terms, credit, mottos))
    
def make_random_utas_grade(students, courses, rg):
    """共通ＩＤ	学籍番号	学生氏名	学生氏名カナ	学生所属	年度	時間割所属	時間割コード	開講科目名	科目コード	科目名	主担当教員名	主担当教員共通ID	開講区分名	合否区分	単位数"""
    # students: ID, 学生証番号, 名前, カナ, 所属研究科
    statuses = ["合格"] * 2 + ["不合格", "履修中", ""]
    utas = []
    for s_utac, s_ids, s_name, s_kana, s_school, s_hobby, s_objective in students:
        p = rg.random()
        for c_school, code, course, code_, course_, f_name, f_id, term, cred, motto in courses:
            if rg.random() < p:
                # this student took this course
                # which year?
                year = rg.randint(2020, 2025)
                st = rg.choice(statuses)
                s_id = rg.choice(s_ids)
                utas.append((s_utac, s_id, s_name, s_kana, s_school,
                             year, c_school, code, course,
                             code_, course_, f_name, f_id,
                             term, st, cred))
    columns = ["共通ＩＤ", "学籍番号", "学生氏名", "学生氏名カナ", "学生所属", "年度",
               "時間割所属", "時間割コード", "開講科目名", "科目コード", "科目名",
               "主担当教員名", "主担当教員共通ID", "開講区分名", "合否区分", "単位数"]
    df = pd.DataFrame(utas, columns=columns)
    return df

def make_num_range_string(start, end, rg):
    if start == end:
        return f"{start}"
    else:
        spcs = rg.choices([" ", ""], k=2)
        return f"{start}{spcs[0]}-{spcs[1]}{end}"

def make_num_set_string(years, rg):
    start = None
    last = None
    S = []
    for y in years:
        if last is None:
            start = y
            last = y
        elif last + 1 < y:
            S.append(make_num_range_string(start, last, rg))
            start = y
            last = y
        else:
            last = y
    if last is not None:
        S.append(make_num_range_string(start, last, rg))
    sep = rg.choice([";", ",", " "])
    spcs = rg.choices([" ", ""], k=2)
    sep = f"{spcs[0]}{sep}{spcs[1]}"
    return sep.join(S)
    
def make_random_program_courses(courses, m, rg):
    """
    修了プログラムが対象とするコースを, courses の中からランダム(rg)に
    m個選び, DataFrame化
    """
    courses = courses[:]
    rg.shuffle(courses)
    courses = courses[:m]
    year_choices = list(range(2020, 2026))
    p_courses = []
    for c_school, code, course, code_, course_, f_name, f_id, term, cred, motto in courses:
        p = rg.random()
        if p < 0.2:
            years = np.nan      # every year
        else:
            years = make_num_set_string([y for y in year_choices if rg.random() < 0.8], rg)
        p_courses.append((code, course, years, motto))
    columns = ["科目コード", "科目名", "対象年度", "モットー的なもの"]
    df = pd.DataFrame(p_courses, columns=columns)
    return df

def make_random_program_students(students, m, rg):
    """
    修了プログラムに登録した学生を, students の中からランダム(rg)に
    m個選び, DataFrame化
    """
    students = students[:]
    rg.shuffle(students)
    students = students[:m]
    # s_id, s_no, s_name, s_kana, s_school, s_hobby, s_objective
    p_students = [(s_utac, s_id, s_name, s_hobby, s_objective)
                  for s_utac, s_ids, s_name, s_kana, s_school,
                      s_hobby, s_objective in students
                      for s_id in s_ids]
    columns = ["共通ID", "学籍番号", "学生氏名", "ごしゅみ", "もくひょー"]
    df = pd.DataFrame(p_students, columns=columns)
    return df

def main():
    rg = random.Random()
    rg.seed(12345)
    n_courses = len(COURSE_NAMES)
    n_students = len(STUDENT_NAMES)
    n_faculties = len(FACULTY_NAMES)
    n_program_courses = n_courses // 2
    n_program_students = n_students // 2
    students = make_random_students(n_students, rg)
    faculties = make_random_faculties(n_faculties, rg)
    courses = make_random_courses(n_courses, faculties, rg)
    utas_grade = make_random_utas_grade(students, courses, rg)
    program_courses = make_random_program_courses(courses, n_program_courses, rg)
    program_students = make_random_program_students(students, n_program_students, rg)
    utas_grade.to_excel("utas_grade.xlsx", index=False)
    program_courses.to_excel("科目一覧.xlsx", index=False)
    program_students.to_excel("登録学生一覧.xlsx", index=False)

main()
