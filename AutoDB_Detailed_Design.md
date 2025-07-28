# 詳細設計書：端末マスタ管理 半自動化ツール

---

## 1. システム概要

本システムは、ユーザー端末へのアプリケーションインストールに関する依頼情報をもとに、  
Oracleデータベースの「端末マスタ」テーブルを管理・更新するための半自動化支援ツールです。

主に Excel（VBA）と Windowsバッチファイル（.bat）を用い、以下の処理を支援します：

- ヒアリング情報の取込
- DBマスタ情報の取得・整形
- SQL文の出力
- 更新前後データの差分比較

---

## 2. モジュール構成

```
AutoDB_Processor/
├── ユーザー入力        ：ヒアリングシート(.txt)
├── バッチ処理         ：export_terminal.bat, export_terminal.sql
├── マクロExcel        ：import_master.xlsm
├── 出力フォルダ       ：AutoDB_output/
└── オプションテンプレ ：templates/sql_template.txt
```

---

## 3. 処理フロー詳細

### 3.1 ヒアリングシートの取り込み（VBA）

- 入力：`*.txt`（「項目名：値」形式）
- 処理：行単位で読み込み、区切り文字「：」で分割
- 出力：ExcelシートのA列に項目名、B列に値を格納
- 備考：チェック機能（未入力・形式ミス）を搭載予定

---

### 3.2 Oracleマスタ情報の取得（BAT＋SQL）

#### バッチファイル：`export_terminal.bat`
- Oracle接続文字列とSQL*Plusを使用
- SQL実行スクリプト：`export_terminal.sql` を呼び出し

#### SQLスクリプト：`export_terminal.sql`
- `SELECT ... FROM MST_Terminal` を実行し、CSVとしてSPOOL出力
- 出力先：`AutoDB_output/mst_terminal.csv`

---

### 3.3 CSVからExcel変換（VBA）

- ファイル：`mst_terminal.csv`
- 処理：
  - `Workbooks.OpenText` にてCSVを開く
  - `.SaveAs` にて `.xlsx` として保存（`terminal_master_before.xlsx` 等）

---

### 3.4 SQL出力支援（VBAオプション）

- テンプレート：`templates/sql_template.txt`
- 処理：
  - プレースホルダ（{terminal_id} など）を置換
  - 出力形式：Excelシート または テキストファイル

---

### 3.5 差分比較処理（VBA）

- 入力：
  - `terminal_master_before.xlsx`
  - `terminal_master_after.xlsx`
- 処理：
  - 主キーを軸に同一行を突き合わせ
  - セル単位で内容を比較し、差異がある場合は着色表示
  - 差分一覧を `terminal_diff.xlsx` に保存

---

## 4. UI設計（import_master.xlsm）

- ヒアリング取込ボタン
- CSV変換ボタン
- SQL出力ボタン（オプション）
- 差分比較ボタン
- チェックボックス（例：固有番号か共通番号か）

---

## 5. データ設計

### 5.1 MST_Terminalテーブル（想定）

| カラム名          | 型         | 説明             |
|------------------|------------|------------------|
| terminal_id      | VARCHAR2   | 主キー           |
| terminal_name    | VARCHAR2   | 端末名           |
| app_code         | VARCHAR2   | アプリコード     |
| area             | VARCHAR2   | 利用エリア       |
| team_name        | VARCHAR2   | 利用チーム名     |
| update_timestamp | DATE       | 更新日時         |

※カラム構成は実DBに合わせて調整

---

## 6. 例外・エラーハンドリング

- ファイル未存在チェック
- 読込中エラー／アクセス拒否時のメッセージ表示
- マスタファイルの形式崩れ（列数不一致）時に警告表示

---

## 7. 拡張性・保守性への配慮

- SQLテンプレートや比較対象列はExcel上から変更可能
- マスタ保存ファイルは日付付きファイル名を推奨（差分履歴を追いやすく）
- 複数条件（チーム別・エリア別）の柔軟な対応を見据えたロジック分離

---

## 8. セキュリティ対応

- DB更新はあくまで手動で実施
- VBA上のID/PASS情報は保持しない（BAT側で管理）
- 不要ファイル（中間CSVなど）は自動削除処理も検討

---

（以上）
