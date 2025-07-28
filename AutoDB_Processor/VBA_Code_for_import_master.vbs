' ====================================================================
' 端末マスタ管理半自動化ツール VBAコード
' ファイル名: import_master.xlsm 用のマクロコード
' ====================================================================

' ====================================================================
' 1. ヒアリングシート取り込み処理
' ====================================================================
Sub ImportHearingSheet()
    On Error GoTo ErrorHandler
    
    Dim filePath As String
    Dim fileContent As String
    Dim lines As Variant
    Dim i As Long
    Dim colonPos As Integer
    Dim itemName As String
    Dim itemValue As String
    
    ' ファイル選択ダイアログ
    filePath = Application.GetOpenFilename("テキストファイル (*.txt),*.txt", , "ヒアリングシートを選択してください")
    
    If filePath = "False" Then
        Exit Sub
    End If
    
    ' ファイル読み込み
    Open filePath For Input As #1
    fileContent = Input$(LOF(1), 1)
    Close #1
    
    ' 行ごとに分割
    lines = Split(fileContent, vbCrLf)
    
    ' ヒアリング情報シートをクリア
    Worksheets("ヒアリング情報").Range("A:B").Clear
    Worksheets("ヒアリング情報").Range("A1").Value = "項目名"
    Worksheets("ヒアリング情報").Range("B1").Value = "値"
    
    ' データ取り込み
    For i = 0 To UBound(lines)
        If InStr(lines(i), "：") > 0 Then
            colonPos = InStr(lines(i), "：")
            itemName = Trim(Left(lines(i), colonPos - 1))
            itemValue = Trim(Mid(lines(i), colonPos + 1))
            
            Worksheets("ヒアリング情報").Range("A" & (i + 2)).Value = itemName
            Worksheets("ヒアリング情報").Range("B" & (i + 2)).Value = itemValue
        End If
    Next i
    
    MsgBox "ヒアリングシートの取り込みが完了しました。", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
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
' 5. ヘルパー関数
' ====================================================================
Function GetHearingValue(ws As Worksheet, itemName As String) As String
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = itemName Then
            GetHearingValue = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i
    
    GetHearingValue = ""
End Function

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