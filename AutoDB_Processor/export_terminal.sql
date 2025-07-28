-- ====================================================================
-- 端末マスタ情報エクスポート用SQLスクリプト
-- Oracle DBから端末マスタ情報をCSV形式で出力
-- ====================================================================

-- 出力設定
SET PAGESIZE 0
SET LINESIZE 1000
SET COLSEP ','
SET TRIMSPOOL ON
SET HEADSEP OFF
SET FEEDBACK OFF
SET ECHO OFF
SET VERIFY OFF

-- CSV出力先の指定
SPOOL AutoDB_output/mst_terminal.csv

-- ヘッダー行の出力
SELECT 'TERMINAL_ID,TERMINAL_NAME,APP_CODE,AREA,TEAM_NAME,UPDATE_TIMESTAMP' FROM DUAL;

-- 端末マスタ情報の取得
-- ※実際のテーブル名・カラム名は環境に合わせて調整してください
SELECT 
    terminal_id || ',' ||
    terminal_name || ',' ||
    app_code || ',' ||
    area || ',' ||
    team_name || ',' ||
    TO_CHAR(update_timestamp, 'YYYY-MM-DD HH24:MI:SS')
FROM MST_Terminal
ORDER BY terminal_id;

-- 出力終了
SPOOL OFF

-- 処理完了メッセージ
SELECT 'エクスポート処理が完了しました。' AS MESSAGE FROM DUAL;

-- 接続終了
EXIT; 