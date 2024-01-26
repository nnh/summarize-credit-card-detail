# クレジットカード明細集計

## 事前設定

### 表示名マスタの設定

1. **A列：表示名:** PDFに出力する項目の文字列を設定してください。
2. **B列：表示順:** PDFに出力する順番を設定してください。
3. **C列：開始年度:** その項目をPDFに出力する最初の年度を設定してください。
4. **D列：終了年度:** その項目をPDFで非表示にする最初の年度を設定してください。指定なしの場合は9999を設定してください。

### 表示名と項目名の対応

1. **A列：表示名:** PDFに出力する項目の文字列を選択してください。
2. **B列：項目名:** CSVファイルの項目名を設定してください。部分一致で表示名に変換します。

### サービス一覧の設定

1. **A列：サービス名:** PDFに出力する項目の文字列を選択してください。
2. **B列：使用者:** 使用者を記載してください。
3. **C列：カテゴリー、D列サービス内容:** PDFに出力する項目の文字列を選択してください。
4. **E列：使用目的:** 必要な情報を記載してください.

## 処理概要

1. **クレジットカード明細の保存**

   - メールの添付ファイルからCSVファイルを取得し保存します。
   - 毎月一日、五日に自動で実行されます。
   - 取り込まれたCSVファイルはスクリプトプロパティの`csvSaveFolderId`で指定したIDのフォルダに格納されます。
   - 手動で取り込みを行う場合は、スプレッドシートのメニューの「クレジットカード明細集計」から「CSV保存」を実行してください。

2. **クレジットカード明細の取り込み**

   - 保存されたCSVファイルの情報をスプレッドシートのCSVシートに取り込みます。
   - 毎月一日、五日に自動で実行されます。
   - 手動で取り込みを行う場合は、スプレッドシートのメニューの「クレジットカード明細集計」から「CSV取り込み」を実行してください。

3. **表示名の設定**

   - CSVシートの表示名の情報をマスタから再取り込みします。CSVファイル内の項目名が変更になった場合等、集計結果が不正になった際は、表示名マスタを修正し、スプレッドシートのメニューの「クレジットカード明細集計」から「表示名設定」を実行してください。

4. **集計用シート作成**

   - 各年度の集計用シートを作成します。毎年一月に次年度分が自動で作成されます。
   - 手動で作成を行う場合は、スプレッドシートのメニューの「クレジットカード明細集計」から「集計用シート作成」を実行してください。

5. **PDF出力**
   - 今年度、前年度の集計用シートから、理事会用資料を作成します。スプレッドシートのメニューの「クレジットカード明細集計」から「PDF作成」を実行してください。
