# 端末マスタ管理半自動化ツール

## 概要
このツールは、端末アプリケーションインストール依頼に伴う端末マスタ更新業務を半自動化するためのExcel+BATファイルベースのシステムです。

## ファイル構成
```
AutoDB_Processor/
├── export_terminal.bat              # Oracle DBからCSVエクスポート用バッチ
├── export_terminal.sql              # SQL*Plus用SQLスクリプト
├── VBA_Code_for_import_master.vbs   # VBAマクロコード（手動インポート用）
├── templates/
│   └── sql_template.txt             # SQLテンプレートファイル
├── AutoDB_output/                   # 出力フォルダ
│   ├── mst_terminal.csv             # Oracle出力CSV（自動生成）
│   ├── terminal_master_before.xlsx  # 更新前マスタ（変換後）
│   ├── terminal_master_after.xlsx   # 更新後マスタ（変換後）
│   ├── terminal_diff.xlsx           # 差分比較結果
│   └── generated_sql.sql            # 生成されたSQL
└── README.md                        # このファイル
```

## セットアップ手順

### 1. Oracleデータベース接続設定
`export_terminal.bat` を編集し、以下の接続情報を実際の環境に合わせて設定してください：

```batch
set DB_USER=your_username          # データベースユーザー名
set DB_PASS=your_password          # データベースパスワード  
set DB_CONNECT=your_server:1521/your_service  # 接続文字列
```

### 2. Excelマクロブック作成
1. 新しいExcelブックを作成し、「import_master.xlsm」として保存
2. Alt+F11でVBAエディタを開く
3. `VBA_Code_for_import_master.vbs` の内容をコピーして新しいモジュールに貼り付け
4. 必要に応じて以下のワークシートを作成：
   - 「ヒアリング情報」シート
   - 「制御パネル」シート

### 3. SQLスクリプト調整
`export_terminal.sql` 内のテーブル名・カラム名を実際のデータベース構造に合わせて調整してください。

## 使用手順

### 基本フロー
1. **初回マスタ取得**: `export_terminal.bat` を実行
2. **Excel変換**: `import_master.xlsm` を開き、「CSV→Excel変換」マクロで更新前データを保存
3. **DB更新作業**: 手動でデータベース更新を実施
4. **更新後マスタ取得**: 再度 `export_terminal.bat` を実行
5. **差分比較**: Excel内の「差分比較実行」マクロで変更内容を確認

### 詳細操作

#### A. ヒアリング情報の取り込み
1. ユーザーから提供された `.txt` ファイル（項目名：値形式）を準備
2. Excelの「ヒアリングシート取り込み」マクロを実行
3. ファイル選択ダイアログで対象ファイルを選択

#### B. CSV→Excel変換
1. `export_terminal.bat` 実行後に生成される `mst_terminal.csv` を変換
2. 更新前/更新後を選択して適切なファイル名で保存

#### C. SQL自動生成（オプション）
1. ヒアリング情報が入力済みの状態で「SQL生成」マクロを実行
2. テンプレートのプレースホルダが置換されたSQLファイルが出力

#### D. 差分比較
1. 更新前後の `.xlsx` ファイルが揃った状態で実行
2. 差分はハイライト表示され、`terminal_diff.xlsx` に保存

## 注意事項

### セキュリティ
- データベースの接続情報は `.bat` ファイルで管理
- 実際のDB更新は手動で実施（自動更新はしない）

### カスタマイズ
- テーブル構造に応じて `export_terminal.sql` のSELECT文を調整
- `templates/sql_template.txt` でSQL生成テンプレートを変更可能
- VBAコード内の比較対象カラム数は環境に応じて調整

### トラブルシューティング
- Oracle SQL*Plusが正しくインストールされていることを確認
- ファイルパスに日本語や特殊文字が含まれていないことを確認
- Excelのマクロセキュリティ設定でVBA実行が許可されていることを確認

## 拡張予定機能
- 入力値チェック機能の強化
- ログ出力機能
- スケジュール実行対応
- エリア・チーム別の固有ルール対応

## 問い合わせ
システムに関する問い合わせや改善要望は、開発担当者までご連絡ください。 