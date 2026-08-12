三幸町自治会 プログラム配布リスト生成ツール
Jichikai Program Distribution Tool

三幸町自治会 祭りプログラム・金券・粗品 配布リスト 自動生成ツール（現在バージョン: v1.7.0）

フォルダ構成
jichikai_program_distribution_tool/
  program_distribution_tool.py   ← メインスクリプト
  installer.iss                  ← Inno Setup インストーラースクリプト
  version_info.txt               ← EXEのバージョン情報（PyInstaller用）
  app_icon_final.ico             ← アプリアイコン
  README.md                      ← このファイル
一般ユーザー向け：インストール方法

開発環境を用意しなくても、以下の手順でそのまま使用できます。

Releases ページから最新版の jichikai_program_distribution_tool_setup.zip をダウンロードし、右クリック→「すべて展開」で解凍する
解凍してできた jichikai_program_distribution_tool_setup.exe を実行（管理者権限は不要です）
インストール後、スタートメニューまたはデスクトップのショートカットから起動
⚠ インストール時に警告が出た場合

このインストーラーは個人・自治会内での配布用のため、Microsoftへの発行元登録（コード署名）を行っていません。 そのため、インストーラーを実行すると Windows から次のような警告が表示されることがあります。

ブラウザのダウンロードバーに「このファイルは危険な場合があります」「保持しますか？」等の表示が出た場合 → 「保持」または「詳細情報」→「実行」 を選んでください
実行時に「WindowsによってPCが保護されました」という青い画面（SmartScreen）が出た場合 → 「詳細情報」をクリック→表示された「実行」ボタン を押してください

いずれも、既知のソフトであれば安全に進めて問題ありません。表示が不安な場合は、開発者（三幸町自治会）にご確認ください。

概要

区部ごとに担当の部長へプログラム冊子・金券（100円）・粗品を渡し、部長から各寄付者へ配布してもらうための一覧表を生成します。

入力ファイル
ファイル	内容	使用列
プログラム広告順番案 Excel (.xlsx)	町外・町内・特別の寄付者リスト	A列=ID, B列=区部, C列=氏名, G列=寄付金額
積立者リスト Excel (.xlsx)	積立者リスト	A列=ID, B列=区部, C列=氏名
出力テンプレート (.xltx)	出力フォーマットの定義ファイル	—
出力ファイル
形式：Excel (.xlsx)
出力先：プログラム広告順番案 Excel と同フォルダの output サブフォルダ
出力列構成
列	内容	ルール
A	ID	入力から取得
B	ID漢字	IDのプレフィックス部分（例：外、内、積立）
C	ID数字	IDの番号部分・先頭ゼロ保持（例：01、001）
D	区部	入力から取得（例：1-1）
E	名前	入力C列から取得
F	100円券枚数	種別ごとのルールで計算
G	No.つきプログラム	積立=1、それ以外=0
H	No.なしプログラム	町外/町内=1、それ以外=0
I	粗品	町外/町内=1、それ以外=0
配布物ルール
種別	ID例	100円券	No.つき	No.なし	粗品
積立	積立001	5枚固定	1	0	0
町外	外01	寄付金額÷1000枚	0	1	1
町内	内01	寄付金額÷1000枚	0	1	1
特別	特001	寄付金額÷1000枚	0	0	0
並び順

区部順（1-1 → 1-2 → 2-1 …）→ 種別順（積立→外→内→特）→ ID番号順

使い方
ツールを起動すると、ファイル選択画面が開きます
以下の3つのファイルを選択してください：
項目	内容
プログラム広告順番 Excel	寄付者データのExcelファイル
積立者リスト Excel	積立者データのExcelファイル
出力テンプレートファイル	出力先のExcelテンプレート（.xltx）
各ファイルに対してシートを選択します
出力ファイル名を入力します
「生成」ボタンを押すと処理が実行されます
入力ファイルと同じフォルダの output/ フォルダに結果ファイルが保存されます
シート構成・印刷設定
区部ごとに別シートを生成（シートタブ名：「1区1部」「1区2部」など）
各シートの表示：ページレイアウト表示
印刷ヘッダー左：「X区X部　部長様」（区部ごとに変わる）
印刷ヘッダー中央：「プログラム・金券・記念品配布一覧」
印刷フッター中央：ページ番号
印刷範囲：A〜I列（データ最終行まで）
タイトル行：2行目（全ページに繰り返し印刷）
枠線：印刷あり
開発環境（コードを修正する場合）
Windows 11
Python 3.x
必要ライブラリ（以下でインストール）
pip install openpyxl
スクリプトを実行する
python program_distribution_tool.py
EXE化・インストーラー作成手順（開発者向け）
Step 1: バージョン番号を更新

以下3箇所のバージョン番号を揃えて更新する：

program_distribution_tool.py の APP_VERSION
version_info.txt の filevers / prodvers / FileVersion / ProductVersion
installer.iss の MyAppVersion
Step 2: EXEファイルを作成（PyInstaller）
powershell
python -m PyInstaller `
  --onefile `
  --windowed `
  --icon="app_icon_final.ico" `
  --add-data="app_icon_final.ico;." `
  --version-file="version_info.txt" `
  --noupx `
  --name="jichikai_program_distribution_tool" `
  program_distribution_tool.py

→ dist/jichikai_program_distribution_tool.exe が生成されます

Step 3: インストーラーを作成（Inno Setup）

installer.iss をInno Setup Compilerで開き、「Build」→「Compile」を実行する。

→ Output/jichikai_program_distribution_tool_setup.exe が生成されます

Step 4: GitHubへ公開

公開する際は、①ソースコードの更新と②インストーラー（実行ファイル）の公開の、 性質の異なる2つの作業を行います。①はコードの保管、②はダウンロード用ファイルの公開で、 保存される場所がまったく別なので、両方とも行う必要があります。

① ソースコード一式をリポジトリへpush

.gitignore で dist/・Output/・build/ を除外しているため、 git add . を実行しても以下のようなソースファイルだけが対象になります：

program_distribution_tool.py
installer.iss
version_info.txt
README.md
app_icon_final.ico
powershell
git add .
git commit -m "vX.X.X: 変更内容"
git push
② インストーラーをzip化してReleasesページで公開

準備：exeをzip化

エクスプローラーで Output\jichikai_program_distribution_tool_setup.exe を右クリック
「送る」→「圧縮(zip形式)フォルダー」を選択
同じ場所に jichikai_program_distribution_tool_setup.zip が作成される （zipにする理由：ブラウザがexeを未確認の実行ファイルとして警告・ブロックし、 ダウンロード中にファイル名が変わってしまう現象を防ぐため）

GitHub側の操作

ブラウザで Releases ページを開く
「Draft a new release」ボタンをクリック
「Choose a tag」の入力欄にバージョン番号（例：v1.7.0）を入力し、 「Create new tag: v1.7.0 on publish」をクリックして選択
「Release title」にタイトルを入力 （例：三幸町自治会 プログラム配布リスト生成ツール v1.7.0）
本文欄に変更点を記入
画面下の点線の枠「Attach binaries by dropping them here or selecting them」に jichikai_program_distribution_tool_setup.zip をドラッグ＆ドロップ
アップロード完了を確認後、「Publish release」ボタンをクリックして公開
変更履歴
v1.7.0
アイコンファイル名を app_icon_final.ico に統一
iconbitmap(default=...) でウィンドウ左上・タスクバー両方にアイコンを反映
v1.6.0
生成される全シートの表示をページレイアウト表示に設定
v1.5.0
印刷フッター中央にページ番号を設定
印刷範囲を A1:I（データ最終行）に設定
タイトル行（2行目）を全ページ繰り返し印刷に設定
枠線印刷を有効化
print_area の設定順序バグを修正（last_data_row 確定後に設定）
v1.4.0
印刷ヘッダー中央に「プログラム・金券・記念品配布一覧」を追加
v1.3.0
出力列構成を変更：G列「プログラム」を「No.つき」「No.なし」の2列に分割、I列「粗品」を追加（計9列構成）
データ書き込み後の空白行を自動削除
SUBTOTAL式をG・H・I列に対応
v1.2.0
3行目のデータ行ヘッダーを廃止
各シートの印刷ヘッダー（ページレイアウト→ヘッダー/フッター左側）に「X区X部　部長様」を設定
GUIバージョン表示を ttk.Label に変更（表示が確実になった）
GitHubリポジトリ公開
v1.1.0
B列・C列の書き込みを数式から値に変更（先頭ゼロの欠落を修正）
例：外01 → B=「外」、C=「01」（修正前は B=「外0」、C=「1」）
集計結果に種別ごとの金額合計を追加
バージョン表示を APP_VERSION 定数から参照するよう統一
v1.0.0
初版リリース
基本機能実装：プログラム広告順番 Excel・積立者リスト Excel を読み込み、区部ごとにシートを分けて配布リストを生成
区部ごとに別シート出力（シートタブ名：「X区X部」）
並び順：区部→積立→外→内→特→ID番号順
リポジトリ

https://github.com/miyuki-jichikai/jichikai_program_distribution_tool

お問い合わせ

三幸町自治会

コンテンツ
installer.iss

ISS

README.md

179行

MD

version_info.txt

28行

TXT

README.md

192行

MD