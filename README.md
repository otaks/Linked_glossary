# Linked_glossary
キーワードリンクを持つHTML単語集を作成するExcel VBAアプリ

### 用途
エクセルで作成した用語集の各単語説明に素早くアクセスする

### 動作確認環境
Windows 10 Home, Microsoft Excel 2010

### 環境設定
- Alt+F11で開いたVBE - ツール - 参照設定から「Microsoft Scripting Runtime」
　「Microsoft Visual Basic for Applications Extensibility 5.3」をチェック

### 使用手順
1. Linked_glossary.xlsm　単語集シートに単語情報を入力
1. HTML出力ボタン押下
1. 同ディレクトリに出力される、単語集.htmlをブラウザで開く

### 補足
- PUSH時は、ExportAllマクロ（モジュール書き出し）を実行すること

### 使用イメージ
#### <エクセルで作成した単語集>
<img src="https://github.com/otaks/img/blob/master/LinkWords_cap1.png" width="500px">

#### <上記から作成したHTML単語集>
<img src="https://github.com/otaks/img/blob/master/LinkWords_cap2.png" width="500px">
