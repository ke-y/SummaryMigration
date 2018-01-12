=============================
  SummaryMigration  Ver 1.2
=============================
[Name]
SummaryMigration.exe

[Version]
1.2

[Overview]
文書データ移行ツール

[AUTHORS]     
ESM)K.Y

[Requirement]
Microsoft Windows 7 以降
Microsoft .NET Framework 4.5 以降

[Install]
Cドライブ直下に「SummaryMigration」フォルダごと格納

[UnInstall]   
「SummaryMigration」フォルダごと削除。
（レジストリ等は変更していない）

[Restriction]
CSVファイルとPDFが同一フォルダ内に存在しないと、移行対象として
認識されない。

[CHANGE LOG]
1.2：2018/01/12
メニュー機能 追加
PDFコピー処理を変更。

1.1：2017/12/10
抽出範囲指定、一覧出力機能 追加

1.0：2017/11/30
初版作成



[Future]
・途中Commit処理
・同階層にCSVがない場合でも対象とできるよう、DB登録時の処理を見直し
・件数が大量にある場合に、getFilesでは処理しきれない。
　ロジック変更してDB登録化を検討
・移行中の画面フリーズ対応
・一覧の絞込み表示機能
