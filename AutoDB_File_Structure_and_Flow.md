# Excel+BATによる端末マスタ更新 半自動化構成

## 📁 ファイル構成案

```
AutoDB_Processor/
├── export_terminal.bat             # Oracle DBからCSVをエクスポート
├── export_terminal.sql             # SQL*Plusで使用するSQLファイル（SPOOLで出力）
├── import_master.xlsm              # VBAマクロ付きブック（CSV → xlsx変換、差分比較）
├── templates/
│   └── sql_template.txt            # SQLテンプレート（任意）
├── AutoDB_output/
│   ├── mst_terminal.csv            # Oracleから出力されたCSV
│   ├── terminal_master_before.xlsx # 更新前のマスタ情報
│   ├── terminal_master_after.xlsx  # 更新後のマスタ情報
│   └── terminal_diff.xlsx          # 差分結果（差分比較後に生成）
```

---

## 🔁 実行順フロー図

```
1. [ユーザー] export_terminal.bat を実行（DB → CSV）
       ↓
2. [システム] mst_terminal.csv が AutoDB_output に生成される
       ↓
3. [ユーザー] import_master.xlsm を開き、VBAマクロ [CSV → xlsx 変換] 実行
       ↓
4. [ユーザー] terminal_master_before.xlsx を更新作業前に保存
       ↓
5. [ユーザー] DB更新作業（手動でSQL実行）
       ↓
6. [ユーザー] 再度 export_terminal.bat を実行（DB → 最新CSV）
       ↓
7. [ユーザー] terminal_master_after.xlsx に変換保存
       ↓
8. [ユーザー] import_master.xlsm の [差分比較マクロ] を実行
       ↓
9. [システム] 差分結果を terminal_diff.xlsx に出力
```

---

## 📌 補足

- `.xlsm` は VBA付きテンプレートとして活用可能
- 共有フォルダに `AutoDB_output/` を置くことで共通アクセスを確保
- 実行順は変更可（例：日次バッチで `before.xlsx` 自動保存 など）
