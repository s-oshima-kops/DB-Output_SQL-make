@echo off
rem ====================================================================
rem 【端末マスタ情報エクスポートバッチ】
rem 機能: Oracle データベースから端末マスタ情報をCSV形式で出力
rem 作成者: システム管理部
rem 更新日: 2024-01-01
rem ====================================================================

echo.
echo ========================================
echo 端末マスタ情報エクスポート処理開始
echo 開始時刻: [%date% %time%]
echo ========================================

rem ====================================================================
rem ★★★ ここを現場の環境に合わせて修正してください ★★★
rem 初心者の方でも安全に変更できるよう、修正箇所を <<<>>> で囲んでいます
rem ====================================================================

rem ┌─────────────────────────────────────┐
rem │ 🔐 データベース接続設定【必須修正】          │
rem └─────────────────────────────────────┘
rem ※重要※ システム管理者から提供された情報に置き換えてください

set DB_USER=<<<your_username>>>
rem ↑ データベースのユーザー名を入力
rem   例: set DB_USER=terminal_user

set DB_PASS=<<<your_password>>>
rem ↑ データベースのパスワードを入力  
rem   例: set DB_PASS=mypassword123

set DB_CONNECT=<<<your_server:1521/your_service>>>
rem ↑ データベースの接続文字列を入力
rem   例: set DB_CONNECT=dbserver01:1521/ORCL
rem   形式: サーバー名:ポート番号/サービス名

rem ┌─────────────────────────────────────┐
rem │ 📁 出力フォルダ設定（通常は変更不要）        │  
rem └─────────────────────────────────────┘
set OUTPUT_FOLDER=AutoDB_output
rem ↑ CSVファイルの出力先フォルダ名
rem   変更する場合: AutoDB_output → <<<新しいフォルダ名>>>

rem ┌─────────────────────────────────────┐
rem │ ⚙️ Oracle環境設定（必要な場合のみ）         │
rem └─────────────────────────────────────┘
rem Oracle クライアントが標準的な場所にない場合は以下を有効化
rem （行頭の「rem」を削除すると有効になります）

rem set ORACLE_HOME=<<<C:\app\oracle\product\11.2.0\client_1>>>
rem ↑ Oracle クライアントのインストールパス
rem   システム管理者に確認してください

rem set PATH=%ORACLE_HOME%\bin;%PATH%

rem ====================================================================
rem 【前処理】出力フォルダの準備
rem ====================================================================

echo 出力フォルダの確認中...
if not exist "%OUTPUT_FOLDER%" (
    echo 出力フォルダが存在しないため作成します: %OUTPUT_FOLDER%
    mkdir "%OUTPUT_FOLDER%"
    if %ERRORLEVEL% neq 0 (
        echo エラー: 出力フォルダの作成に失敗しました
        goto ERROR_EXIT
    )
) else (
    echo 出力フォルダが確認できました: %OUTPUT_FOLDER%
)

rem ====================================================================
rem 【メイン処理】Oracle データベースからCSVエクスポート
rem ====================================================================

echo データベースへの接続準備中...
echo 接続先: %DB_CONNECT%
echo ユーザー: %DB_USER%

rem 接続情報の妥当性チェック
if "%DB_USER%"=="your_username" (
    echo.
    echo 警告: データベース接続情報が初期値のままです
    echo このバッチファイル内の「設定部分」を実際の環境に合わせて修正してください
    echo.
    pause
    goto ERROR_EXIT
)

echo SQL*Plus でマスタ情報を取得中...
echo 実行するSQLファイル: export_terminal.sql

rem SQL*Plus実行（-s オプションで画面出力を抑制）
sqlplus -s %DB_USER%/%DB_PASS%@%DB_CONNECT% @export_terminal.sql

rem ====================================================================
rem 【後処理】実行結果の確認とクリーンアップ
rem ====================================================================

rem SQL*Plus の実行結果をチェック
if %ERRORLEVEL% == 0 (
    echo.
    echo ========================================
    echo 処理完了
    echo 完了時刻: [%date% %time%]
    echo ========================================
    echo.
    echo マスタ情報の取得が正常に完了しました
    echo 出力ファイル: %OUTPUT_FOLDER%\mst_terminal.csv
    echo.
    
    rem 出力ファイルの存在確認
    if exist "%OUTPUT_FOLDER%\mst_terminal.csv" (
        echo ✓ CSVファイルが正常に作成されました
        for %%F in ("%OUTPUT_FOLDER%\mst_terminal.csv") do echo   ファイルサイズ: %%~zF bytes
    ) else (
        echo ⚠ 警告: CSVファイルが見つかりません
        echo SQLの実行は成功しましたが、出力ファイルが作成されていない可能性があります
    )
    
) else (
    echo.
    echo ========================================
    echo エラー発生
    echo エラー発生時刻: [%date% %time%]
    echo ========================================
    echo.
    echo エラー: マスタ情報の取得に失敗しました
    echo エラーレベル: %ERRORLEVEL%
    echo.
    echo 【考えられる原因】
    echo 1. データベース接続情報が間違っている
    echo 2. Oracle クライアントが正しくインストールされていない
    echo 3. ネットワークの問題でデータベースに接続できない
    echo 4. export_terminal.sql ファイルに問題がある
    echo.
    echo 【対処方法】
    echo 1. このファイル内の接続設定を確認してください
    echo 2. SQL*Plus が正しく動作するかテストしてください
    echo 3. システム管理者に連絡してください
    echo.
    goto ERROR_EXIT
)

rem 一時ファイルのクリーンアップ（SQL*Plus が作成するログファイルを削除）
echo 一時ファイルをクリーンアップ中...
if exist "%OUTPUT_FOLDER%\*.log" (
    del /q "%OUTPUT_FOLDER%\*.log"
    echo ✓ ログファイルを削除しました
)

rem 正常終了
echo.
echo 次のステップ: import_master.xlsm を開いて「CSV→Excel変換」を実行してください
echo.
echo 処理を完了しました。Enterキーを押してください。
pause
exit /b 0

rem ====================================================================
rem 【エラー終了処理】
rem ====================================================================
:ERROR_EXIT
echo.
echo 処理が異常終了しました。
echo システム管理者に連絡し、設定を確認してください。
echo.
pause
exit /b 1 