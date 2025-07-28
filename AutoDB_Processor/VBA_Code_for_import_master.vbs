' ====================================================================
' 【端末マスタ管理半自動化ツール VBAコード】
' ファイル名: import_master.xlsm 用のマクロコード
' 作成者: システム管理部
' 更新日: 2024-01-01
' 
' 【初心者向け修正ガイド】
' 1. このファイル内の「設定部分」を現場に合わせて変更してください
' 2. エラーが発生した場合は、まず設定値を確認してください
' 3. 分からない場合は、システム管理者に連絡してください
' ====================================================================

' ====================================================================
' 【★★★ ここを現場に合わせて修正してください ★★★】
' VBA初心者の方でも安全に変更できるよう、修正箇所を明確に示しています
' ====================================================================

' ┌─────────────────────────────────────┐
' │ 📁 ファイルパス設定（通常は変更不要）        │
' └─────────────────────────────────────┘
Const OUTPUT_FOLDER As String = "AutoDB_output"
' ↑ CSVファイルが保存されるフォルダ名
' 　 変更する場合: "AutoDB_output" → "<<<新しいフォルダ名>>>"

Const MASTER_BEFORE_FILE As String = "terminal_master_before.xlsx"
' ↑ 更新前のマスタファイル名
' 　 変更する場合: "terminal_master_before.xlsx" → "<<<新しいファイル名.xlsx>>>"

Const MASTER_AFTER_FILE As String = "terminal_master_after.xlsx"
' ↑ 更新後のマスタファイル名
' 　 変更する場合: "terminal_master_after.xlsx" → "<<<新しいファイル名.xlsx>>>"

Const DIFF_RESULT_FILE As String = "terminal_diff.xlsx"
' ↑ 差分比較結果ファイル名
' 　 変更する場合: "terminal_diff.xlsx" → "<<<新しいファイル名.xlsx>>>"

Const GENERATED_SQL_FILE As String = "generated_sql.sql"
' ↑ 自動生成されるSQLファイル名
' 　 変更する場合: "generated_sql.sql" → "<<<新しいファイル名.sql>>>"

' ┌─────────────────────────────────────┐
' │ 🔢 データ設定（現場のテーブル構造に合わせて変更） │
' └─────────────────────────────────────┘
Const COMPARE_COLUMN_COUNT As Integer = 6
' ↑ 【重要】差分比較で比較するカラム数
' 　 現場のテーブルのカラム数に合わせて変更してください
' 　 例: 端末ID, 端末名, アプリコード, エリア, チーム名, 更新日時 = 6個
' 　 変更する場合: 6 → <<<実際のカラム数>>>
' 　 ※export_terminal.sql のSELECT文のカラム数と一致させる必要があります

Const HEARING_SEPARATOR As String = "："
' ↑ ヒアリングシートの区切り文字
' 　 「項目名：値」の「：」部分
' 　 変更する場合: "：" → "<<<新しい区切り文字>>>"
' 　 例: 半角コロンの場合 "："→ ":"

' ┌─────────────────────────────────────┐
' │ 🎨 表示色設定（差分比較結果の見た目）         │
' └─────────────────────────────────────┘
Const HIGHLIGHT_COLOR_R As Integer = 255
' ↑ ハイライト色の赤成分 (0-255)

Const HIGHLIGHT_COLOR_G As Integer = 255
' ↑ ハイライト色の緑成分 (0-255)  

Const HIGHLIGHT_COLOR_B As Integer = 0
' ↑ ハイライト色の青成分 (0-255)
' 　 現在の設定: 赤255 + 緑255 + 青0 = 黄色
' 　 色の変更例:
' 　   赤色: R=255, G=0, B=0
' 　   青色: R=0, G=0, B=255
' 　   緑色: R=0, G=255, B=0

' ┌─────────────────────────────────────┐
' │ 🐛 デバッグ設定（トラブル時に使用）           │
' └─────────────────────────────────────┘
Const DEBUG_MODE As Boolean = False
' ↑ デバッグ情報の表示ON/OFF
' 　 エラーが発生して詳細を確認したい場合: False → True
' 　 通常運用時: True → False

' ====================================================================
' 【1. ヒアリングシート取り込み処理】
' 機能: ユーザー提供のテキストファイルをExcelに取り込む
' 形式: 「項目名：値」の形式で記載されたファイルを読み込み
' ====================================================================
Sub ImportHearingSheet()
    ' エラーが発生した場合はErrorHandler に飛ぶ
    On Error GoTo ErrorHandler
    
    ' --- 変数宣言 ---
    Dim filePath As String          ' 選択されたファイルのパス
    Dim fileContent As String       ' ファイルの内容全体
    Dim lines As Variant           ' ファイルを行ごとに分割した配列
    Dim i As Long                  ' ループ用カウンタ
    Dim colonPos As Integer        ' 区切り文字「：」の位置
    Dim itemName As String         ' 項目名（：の左側）
    Dim itemValue As String        ' 値（：の右側）
    Dim importCount As Integer     ' 取り込み件数カウンタ
    
    ' デバッグログ表示
    If DEBUG_MODE Then
        Debug.Print "ヒアリングシート取り込み開始: " & Now
    End If
    
    ' --- ステップ1: ファイル選択ダイアログを表示 ---
    filePath = Application.GetOpenFilename( _
        "テキストファイル (*.txt),*.txt", , _
        "ヒアリングシートを選択してください")
    
    ' キャンセルボタンが押された場合は処理終了
    If filePath = "False" Then
        If DEBUG_MODE Then Debug.Print "ユーザーによりキャンセルされました"
        Exit Sub
    End If
    
    ' ファイル存在確認
    If Dir(filePath) = "" Then
        MsgBox "指定されたファイルが見つかりません。" & vbCrLf & _
               "ファイルパス: " & filePath, vbCritical, "ファイルエラー"
        Exit Sub
    End If
    
    ' --- ステップ2: ファイル読み込み ---
    If DEBUG_MODE Then Debug.Print "ファイル読み込み開始: " & filePath
    
    ' ファイルを開いて内容を全て読み込む
    Open filePath For Input As #1
    fileContent = Input$(LOF(1), 1)  ' LOF(1)でファイルサイズを取得し、全て読み込み
    Close #1
    
    ' 読み込み内容が空の場合はエラー
    If Len(Trim(fileContent)) = 0 Then
        MsgBox "ファイルの内容が空です。" & vbCrLf & _
               "正しいヒアリングシートファイルを選択してください。", _
               vbCritical, "ファイル内容エラー"
        Exit Sub
    End If
    
    ' --- ステップ3: ヒアリング情報シートの準備 ---
    ' シート存在確認（無ければ作成）
    Call EnsureWorksheetExists("ヒアリング情報")
    
    ' ヒアリング情報シートをクリア
    Worksheets("ヒアリング情報").Range("A:B").Clear
    
    ' ヘッダー行を設定
    Worksheets("ヒアリング情報").Range("A1").Value = "項目名"
    Worksheets("ヒアリング情報").Range("B1").Value = "値"
    
    ' ヘッダー行を太字にして見やすくする
    Worksheets("ヒアリング情報").Range("A1:B1").Font.Bold = True
    
    ' --- ステップ4: データ取り込みメイン処理 ---
    ' 改行コードで行ごとに分割
    lines = Split(fileContent, vbCrLf)
    importCount = 0  ' 取り込み件数をカウント
    
    ' 各行を処理
    For i = 0 To UBound(lines)
        ' 空行をスキップ
        If Len(Trim(lines(i))) > 0 Then
            ' 区切り文字「：」が含まれているかチェック
            If InStr(lines(i), HEARING_SEPARATOR) > 0 Then
                ' 区切り文字の位置を取得
                colonPos = InStr(lines(i), HEARING_SEPARATOR)
                
                ' 項目名と値に分割
                itemName = Trim(Left(lines(i), colonPos - 1))      ' 左側（項目名）
                itemValue = Trim(Mid(lines(i), colonPos + 1))      ' 右側（値）
                
                ' 項目名が空でない場合のみ取り込み
                If Len(itemName) > 0 Then
                    ' Excelシートに書き込み（i+2は、1行目がヘッダーなので+1、さらに配列が0から始まるので+1）
                    Worksheets("ヒアリング情報").Range("A" & (importCount + 2)).Value = itemName
                    Worksheets("ヒアリング情報").Range("B" & (importCount + 2)).Value = itemValue
                    
                    importCount = importCount + 1  ' 取り込み件数を加算
                    
                    If DEBUG_MODE Then
                        Debug.Print "取り込み: " & itemName & " = " & itemValue
                    End If
                End If
            Else
                ' 区切り文字がない行はスキップ（デバッグ情報として出力）
                If DEBUG_MODE Then
                    Debug.Print "スキップ（区切り文字なし）: " & lines(i)
                End If
            End If
        End If
    Next i
    
    ' --- ステップ5: 結果表示 ---
    ' 列幅を自動調整
    Worksheets("ヒアリング情報").Columns("A:B").AutoFit
    
    ' 完了メッセージ
    MsgBox "ヒアリングシートの取り込みが完了しました。" & vbCrLf & _
           "取り込み件数: " & importCount & "件" & vbCrLf & _
           "ファイル: " & filePath, vbInformation, "取り込み完了"
    
    If DEBUG_MODE Then
        Debug.Print "ヒアリングシート取り込み完了: " & Now & " (" & importCount & "件)"
    End If
    
    Exit Sub
        
ErrorHandler:
    ' エラー発生時の処理
    MsgBox "ヒアリングシート取り込み中にエラーが発生しました。" & vbCrLf & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & vbCrLf & _
           "システム管理者に連絡してください。", _
           vbCritical, "エラー発生"
    
    If DEBUG_MODE Then
        Debug.Print "エラー発生: " & Err.Description & " (番号: " & Err.Number & ")"
    End If
End Sub

' ====================================================================
' 2. CSV → Excel変換処理
' ====================================================================
Sub ConvertCSVToExcel()
    On Error GoTo ErrorHandler
    
    Dim csvPath As String
    Dim xlsxPath As String
    Dim wb As Workbook
    Dim saveType As String
    
    csvPath = ThisWorkbook.Path & "\AutoDB_output\mst_terminal.csv"
    
    ' CSVファイル存在確認
    If Dir(csvPath) = "" Then
        MsgBox "CSVファイルが見つかりません: " & csvPath, vbCritical
        Exit Sub
    End If
    
    ' 保存タイプ選択
    saveType = InputBox("保存タイプを選択してください:" & vbCrLf & _
                       "1: 更新前 (before)" & vbCrLf & _
                       "2: 更新後 (after)", "保存タイプ選択", "1")
    
    If saveType = "1" Then
        xlsxPath = ThisWorkbook.Path & "\AutoDB_output\terminal_master_before.xlsx"
    ElseIf saveType = "2" Then
        xlsxPath = ThisWorkbook.Path & "\AutoDB_output\terminal_master_after.xlsx"
    Else
        MsgBox "無効な選択です。", vbExclamation
        Exit Sub
    End If
    
    ' CSVファイルを開く
    Set wb = Workbooks.Open(csvPath, Local:=True)
    
    ' Excel形式で保存
    wb.SaveAs xlsxPath, xlWorkbookDefault
    wb.Close
    
    MsgBox "変換が完了しました: " & xlsxPath, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    If Not wb Is Nothing Then wb.Close False
End Sub

' ====================================================================
' 3. 差分比較処理
' ====================================================================
Sub CompareMasterFiles()
    On Error GoTo ErrorHandler
    
    Dim beforePath As String
    Dim afterPath As String
    Dim diffPath As String
    Dim wbBefore As Workbook, wbAfter As Workbook, wbDiff As Workbook
    Dim wsBefore As Worksheet, wsAfter As Worksheet, wsDiff As Worksheet
    Dim lastRowBefore As Long, lastRowAfter As Long
    Dim i As Long, j As Long
    Dim diffCount As Long
    
    beforePath = ThisWorkbook.Path & "\AutoDB_output\terminal_master_before.xlsx"
    afterPath = ThisWorkbook.Path & "\AutoDB_output\terminal_master_after.xlsx"
    diffPath = ThisWorkbook.Path & "\AutoDB_output\terminal_diff.xlsx"
    
    ' ファイル存在確認
    If Dir(beforePath) = "" Then
        MsgBox "更新前ファイルが見つかりません: " & beforePath, vbCritical
        Exit Sub
    End If
    
    If Dir(afterPath) = "" Then
        MsgBox "更新後ファイルが見つかりません: " & afterPath, vbCritical
        Exit Sub
    End If
    
    ' ファイルを開く
    Set wbBefore = Workbooks.Open(beforePath)
    Set wbAfter = Workbooks.Open(afterPath)
    Set wbDiff = Workbooks.Add
    
    Set wsBefore = wbBefore.Worksheets(1)
    Set wsAfter = wbAfter.Worksheets(1)
    Set wsDiff = wbDiff.Worksheets(1)
    
    ' 差分結果シートの初期化
    wsDiff.Name = "差分結果"
    wsDiff.Range("A1:F1").Value = Array("端末ID", "項目", "更新前", "更新後", "変更種別", "行番号")
    
    lastRowBefore = wsBefore.Cells(wsBefore.Rows.Count, 1).End(xlUp).Row
    lastRowAfter = wsAfter.Cells(wsAfter.Rows.Count, 1).End(xlUp).Row
    
    diffCount = 1
    
    ' 差分比較処理
    For i = 2 To Application.WorksheetFunction.Max(lastRowBefore, lastRowAfter)
        Dim terminalIdBefore As String, terminalIdAfter As String
        
        If i <= lastRowBefore Then terminalIdBefore = wsBefore.Cells(i, 1).Value
        If i <= lastRowAfter Then terminalIdAfter = wsAfter.Cells(i, 1).Value
        
        ' 各セルを比較
        For j = 1 To 6 ' カラム数に応じて調整
            Dim valueBefore As String, valueAfter As String
            
            If i <= lastRowBefore Then valueBefore = wsBefore.Cells(i, j).Value
            If i <= lastRowAfter Then valueAfter = wsAfter.Cells(i, j).Value
            
            If valueBefore <> valueAfter Then
                diffCount = diffCount + 1
                wsDiff.Cells(diffCount, 1).Value = IIf(terminalIdBefore <> "", terminalIdBefore, terminalIdAfter)
                wsDiff.Cells(diffCount, 2).Value = wsBefore.Cells(1, j).Value ' ヘッダー
                wsDiff.Cells(diffCount, 3).Value = valueBefore
                wsDiff.Cells(diffCount, 4).Value = valueAfter
                wsDiff.Cells(diffCount, 5).Value = IIf(valueBefore = "", "追加", IIf(valueAfter = "", "削除", "変更"))
                wsDiff.Cells(diffCount, 6).Value = i
                
                ' ハイライト表示
                wsDiff.Range("A" & diffCount & ":F" & diffCount).Interior.Color = RGB(255, 255, 0)
            End If
        Next j
    Next i
    
    ' 差分結果を保存
    wbDiff.SaveAs diffPath
    
    ' ファイルを閉じる
    wbBefore.Close False
    wbAfter.Close False
    
    MsgBox "差分比較が完了しました。差分件数: " & (diffCount - 1) & "件" & vbCrLf & _
           "結果ファイル: " & diffPath, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    If Not wbBefore Is Nothing Then wbBefore.Close False
    If Not wbAfter Is Nothing Then wbAfter.Close False
    If Not wbDiff Is Nothing Then wbDiff.Close False
End Sub

' ====================================================================
' 4. SQL出力支援処理
' ====================================================================
Sub GenerateSQL()
    On Error GoTo ErrorHandler
    
    Dim templatePath As String
    Dim templateContent As String
    Dim sqlContent As String
    Dim outputPath As String
    Dim terminalId As String, terminalName As String
    Dim appCode As String, area As String, teamName As String
    
    templatePath = ThisWorkbook.Path & "\templates\sql_template.txt"
    outputPath = ThisWorkbook.Path & "\AutoDB_output\generated_sql.sql"
    
    ' テンプレートファイル存在確認
    If Dir(templatePath) = "" Then
        MsgBox "SQLテンプレートファイルが見つかりません: " & templatePath, vbCritical
        Exit Sub
    End If
    
    ' ヒアリング情報から値を取得
    Dim ws As Worksheet
    Set ws = Worksheets("ヒアリング情報")
    
    terminalId = GetHearingValue(ws, "端末ID")
    terminalName = GetHearingValue(ws, "端末名")
    appCode = GetHearingValue(ws, "アプリコード")
    area = GetHearingValue(ws, "エリア")
    teamName = GetHearingValue(ws, "チーム名")
    
    ' テンプレート読み込み
    Open templatePath For Input As #1
    templateContent = Input$(LOF(1), 1)
    Close #1
    
    ' プレースホルダ置換
    sqlContent = templateContent
    sqlContent = Replace(sqlContent, "{terminal_id}", terminalId)
    sqlContent = Replace(sqlContent, "{terminal_name}", terminalName)
    sqlContent = Replace(sqlContent, "{app_code}", appCode)
    sqlContent = Replace(sqlContent, "{area}", area)
    sqlContent = Replace(sqlContent, "{team_name}", teamName)
    
    ' SQL出力
    Open outputPath For Output As #2
    Print #2, sqlContent
    Close #2
    
    MsgBox "SQLファイルを生成しました: " & outputPath, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

' ====================================================================
' 【5. ヘルパー関数】
' 共通で使用される便利な関数群
' ====================================================================

' --- ワークシート存在確認・作成関数 ---
' 指定した名前のシートが存在しない場合は作成する
Sub EnsureWorksheetExists(sheetName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Worksheets(sheetName)
    
    ' シートが存在しない場合は作成
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = sheetName
        
        If DEBUG_MODE Then
            Debug.Print "ワークシート作成: " & sheetName
        End If
    End If
    
    On Error GoTo 0
End Sub

' --- ヒアリング情報取得関数 ---
' ヒアリング情報シートから指定した項目名の値を取得する
' 引数: ws - ワークシート, itemName - 項目名
' 戻り値: 該当する値（見つからない場合は空文字）
Function GetHearingValue(ws As Worksheet, itemName As String) As String
    Dim lastRow As Long        ' 最終行番号
    Dim i As Long             ' ループ用カウンタ
    
    If DEBUG_MODE Then
        Debug.Print "ヒアリング情報検索: " & itemName
    End If
    
    ' 最終行を取得（A列基準）
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 2行目から最終行まで検索（1行目はヘッダー）
    For i = 2 To lastRow
        ' A列の項目名が一致するかチェック
        If ws.Cells(i, 1).Value = itemName Then
            ' 一致した場合、B列の値を返す
            GetHearingValue = ws.Cells(i, 2).Value
            
            If DEBUG_MODE Then
                Debug.Print "見つかりました: " & itemName & " = " & GetHearingValue
            End If
            
            Exit Function
        End If
    Next i
    
    ' 見つからない場合は空文字を返す
    GetHearingValue = ""
    
    If DEBUG_MODE Then
        Debug.Print "見つかりませんでした: " & itemName
    End If
End Function

' --- ファイルパス作成関数 ---
' 現在のワークブックのパスに相対パスを結合する
Function BuildFilePath(relativePath As String) As String
    BuildFilePath = ThisWorkbook.Path & "\" & relativePath
End Function

' --- 安全なファイル削除関数 ---
' ファイルが存在する場合のみ削除する
Sub SafeDeleteFile(filePath As String)
    On Error Resume Next
    
    If Dir(filePath) <> "" Then
        Kill filePath
        
        If DEBUG_MODE Then
            Debug.Print "ファイル削除: " & filePath
        End If
    End If
    
    On Error GoTo 0
End Sub

' --- デバッグ情報表示関数 ---
' デバッグモード時のみメッセージを表示
Sub DebugLog(message As String)
    If DEBUG_MODE Then
        Debug.Print "[" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "] " & message
    End If
End Sub

' ====================================================================
' 6. 初期化処理（Excelブック作成時に実行）
' ====================================================================
Sub InitializeWorkbook()
    On Error Resume Next
    
    ' 必要なワークシートを作成
    If Worksheets("ヒアリング情報") Is Nothing Then
        Worksheets.Add.Name = "ヒアリング情報"
    End If
    
    If Worksheets("制御パネル") Is Nothing Then
        Worksheets.Add.Name = "制御パネル"
        Call CreateControlPanel
    End If
    
    MsgBox "ワークブックの初期化が完了しました。", vbInformation
End Sub

' ====================================================================
' 7. 制御パネル作成
' ====================================================================
Sub CreateControlPanel()
    Dim ws As Worksheet
    Set ws = Worksheets("制御パネル")
    
    ' タイトル
    ws.Range("B2").Value = "端末マスタ管理半自動化ツール"
    ws.Range("B2").Font.Size = 16
    ws.Range("B2").Font.Bold = True
    
    ' ボタン説明
    ws.Range("B4").Value = "1. ヒアリングシート取り込み"
    ws.Range("B5").Value = "2. CSV→Excel変換"
    ws.Range("B6").Value = "3. 差分比較実行"
    ws.Range("B7").Value = "4. SQL生成"
    ws.Range("B8").Value = "5. 初期化"
    
    ' 注意事項
    ws.Range("B10").Value = "【使用手順】"
    ws.Range("B11").Value = "1. export_terminal.batを実行してCSVを取得"
    ws.Range("B12").Value = "2. 「CSV→Excel変換」で更新前データを保存"
    ws.Range("B13").Value = "3. DB更新作業を手動で実施"
    ws.Range("B14").Value = "4. 再度batを実行して更新後データを取得"
    ws.Range("B15").Value = "5. 「差分比較実行」で結果を確認"
    
    ws.Columns("B:B").AutoFit
End Sub 