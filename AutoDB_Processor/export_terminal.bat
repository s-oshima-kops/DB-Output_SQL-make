@echo off
rem ====================================================================
rem 端末マスタ情報エクスポートバッチ
rem Oracle DBからCSVファイルを出力します
rem ====================================================================

echo [%date% %time%] 端末マスタ情報の取得を開始します...

rem 出力フォルダの存在確認・作成
if not exist "AutoDB_output" (
    echo 出力フォルダを作成します...
    mkdir AutoDB_output
)

rem Oracle環境変数の設定（必要に応じて調整）
rem set ORACLE_HOME=C:\app\oracle\product\11.2.0\client_1
rem set PATH=%ORACLE_HOME%\bin;%PATH%

rem SQL*Plusでスクリプト実行
echo SQL*Plusでマスタ情報を取得中...

rem 接続情報（実際の環境に合わせて変更）
rem set DB_USER=your_username
rem set DB_PASS=your_password  
rem set DB_CONNECT=your_server:1521/your_service

rem SQL*Plus実行（接続情報は実際の環境に合わせて設定）
sqlplus -s %DB_USER%/%DB_PASS%@%DB_CONNECT% @export_terminal.sql

if %ERRORLEVEL% == 0 (
    echo [%date% %time%] マスタ情報の取得が完了しました
    echo 出力ファイル: AutoDB_output\mst_terminal.csv
) else (
    echo [%date% %time%] エラー: マスタ情報の取得に失敗しました
    echo エラーレベル: %ERRORLEVEL%
)

rem 一時ファイルのクリーンアップ
if exist "AutoDB_output\*.log" del /q "AutoDB_output\*.log"

echo 処理を完了しました。Enterキーを押してください。
pause 