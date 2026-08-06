Attribute VB_Name = "Mod_KabutanImport_v4"
Option Explicit

' ================================================================
' 株探データ取込 v4
' 設計:
'   1. データ取込シート 2行目(本日行) と 16行目以降(履歴)を両方読む
'   2. 日付で降順ソート(最新→最古)
'   3. 対象銘柄行のE～IT列(250日分)をクリア
'   4. 5シート(終値/高値/安値/始値/出来高)のE列=最新, F列=1日前,...に書込
'   5. 3行目の日付ヘッダも5シート全部を同時更新
'   6. 出来高=0 の日はスキップ(欠損扱い)
' ================================================================

Private Const PWD As String = "ne19480314"

Public Sub 株探データ一括取込()
    Dim dataWs As Worksheet
    On Error Resume Next
    Set dataWs = ThisWorkbook.Sheets("データ取込")
    On Error GoTo 0

    If dataWs Is Nothing Then
        MsgBox "「データ取込」シートが見つかりません。", vbExclamation, "取込エラー"
        Exit Sub
    End If

    ' --- 銘柄コードをデータ取込シートB1から取得 ---
    Dim codeInput As String
    codeInput = Trim(CStr(dataWs.Cells(1, 2).Value))
    If codeInput = "" Then
        MsgBox "データ取込シートの B1 に銘柄コードが入力されていません。" & vbCrLf & vbCrLf & _
               "B1 に対象銘柄のコード(例: 5706)を入力してから" & vbCrLf & _
               "このマクロを実行してください。", _
               vbExclamation, "銘柄コード未入力"
        Exit Sub
    End If

    ' --- 銘柄管理シートから対象行を検索 ---
    Dim mws As Worksheet
    On Error Resume Next
    Set mws = ThisWorkbook.Sheets("銘柄管理")
    On Error GoTo 0
    If mws Is Nothing Then
        MsgBox "「銘柄管理」シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim targetRow As Long: targetRow = 0
    Dim searchR As Long
    Dim searchLast As Long
    searchLast = mws.Cells(mws.Rows.Count, 2).End(xlUp).Row
    If searchLast < 6 Then searchLast = 6

    For searchR = 6 To searchLast
        If Trim(CStr(mws.Cells(searchR, 2).Value)) = codeInput Then
            targetRow = searchR
            Exit For
        End If
    Next searchR

    If targetRow = 0 Then
        MsgBox "銘柄コード「" & codeInput & "」が銘柄管理シートに見つかりません。" & vbCrLf & vbCrLf & _
               "先に銘柄管理シートのB列に " & codeInput & " を追加してから" & vbCrLf & _
               "再度このマクロを実行してください。", _
               vbExclamation, "銘柄未登録"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' --- ヘッダ(データ取込シート1行目)から項目列を検出 ---
    Dim openCol As Integer: openCol = 2
    Dim highCol As Integer: highCol = 3
    Dim lowCol As Integer:  lowCol = 4
    Dim closeCol As Integer: closeCol = 5
    Dim volCol As Integer:  volCol = 8

    Dim c As Integer
    For c = 1 To 10
        Dim hVal As String
        hVal = CStr(dataWs.Cells(1, c).Value)
        Select Case hVal
            Case "始値":       openCol = c
            Case "高値":       highCol = c
            Case "安値":       lowCol = c
            Case "終値":       closeCol = c
            Case "出来高", "売買高", "売買高(株)", "売買高（株）"
                volCol = c
        End Select
    Next c

    ' --- 日付行を収集 (本日行 row2 + 履歴 row16以降) ---
    Dim dateRows(300) As Long
    Dim dateVals(300) As Date
    Dim dateCount As Long: dateCount = 0

    ' 本日行 row2
    If IsDate(dataWs.Cells(2, 1).Value) Then
        dateRows(dateCount) = 2
        dateVals(dateCount) = CDate(dataWs.Cells(2, 1).Value)
        dateCount = dateCount + 1
    End If

    ' 履歴 row16以降
    Dim lastDataRow As Long
    lastDataRow = dataWs.Cells(dataWs.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 16 To lastDataRow
        If IsDate(dataWs.Cells(i, 1).Value) Then
            If dateCount < 300 Then
                dateRows(dateCount) = i
                dateVals(dateCount) = CDate(dataWs.Cells(i, 1).Value)
                dateCount = dateCount + 1
            End If
        End If
    Next i

    If dateCount = 0 Then
        Application.ScreenUpdating = True
        MsgBox "データ取込シートに日付データが見つかりません。" & vbCrLf & _
               "2行目(本日)または16行目以降を確認してください。", _
               vbExclamation, "データなし"
        Exit Sub
    End If

    ' --- 日付で降順ソート (バブルソート) ---
    Dim j As Long, tmpRow As Long
    Dim tmpDate As Date
    For i = 0 To dateCount - 2
        For j = 0 To dateCount - 2 - i
            If dateVals(j) < dateVals(j + 1) Then
                tmpDate = dateVals(j)
                dateVals(j) = dateVals(j + 1)
                dateVals(j + 1) = tmpDate
                tmpRow = dateRows(j)
                dateRows(j) = dateRows(j + 1)
                dateRows(j + 1) = tmpRow
            End If
        Next j
    Next i

    ' --- 銘柄名を取得 ---
    Dim meiName As String
    meiName = CStr(mws.Cells(targetRow, 3).Value)

    ' --- 確認ダイアログ ---
    Dim ret As Integer
    ret = MsgBox("以下の内容で取込みます。" & vbCrLf & vbCrLf & _
                 "銘柄: " & codeInput & " " & meiName & vbCrLf & _
                 "取込件数: " & dateCount & " 日分" & vbCrLf & _
                 "日付範囲: " & Format(dateVals(0), "m/d") & " ～ " & _
                 Format(dateVals(dateCount - 1), "m/d") & vbCrLf & _
                 "取込先行: " & targetRow & " 行目(自動検索)" & vbCrLf & vbCrLf & _
                 "※ 対象銘柄行のE～IT列(250日分)が株探データで上書きされます" & vbCrLf & _
                 "※ 3行目の日付ヘッダも全OHLCVシートで株探日付に更新されます" & vbCrLf & vbCrLf & _
                 "進めますか？", _
                 vbYesNo + vbQuestion, "取込確認")
    If ret = vbNo Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' --- 5シートに書込 ---
    Dim ShNames As Variant
    ShNames = Array("終値", "高値", "安値", "始値", "出来高")
    Dim srcCols As Variant
    srcCols = Array(closeCol, highCol, lowCol, openCol, volCol)

    Dim writeCount As Long: writeCount = 0
    Dim skipCount As Long: skipCount = 0

    Dim s As Integer
    For s = 0 To 4
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(ShNames(s))
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextSheet

        ' E～IT (col 5～254) をクリア
        ws.Range(ws.Cells(targetRow, 5), ws.Cells(targetRow, 254)).ClearContents

        ' 書込 + 3行目日付ヘッダ更新
        Dim d As Long
        For d = 0 To dateCount - 1
            Dim col As Integer: col = 5 + d
            If col > 254 Then Exit For

            Dim srcR As Long: srcR = dateRows(d)
            Dim srcV As Variant
            srcV = dataWs.Cells(srcR, srcCols(s)).Value

            Dim numV As Double: numV = 0
            On Error Resume Next
            numV = CDbl(Replace(CStr(srcV), ",", ""))
            On Error GoTo 0

            If numV > 0 Then
                ws.Cells(targetRow, col).Value = numV
                ws.Cells(targetRow, col).NumberFormat = "#,##0"
                writeCount = writeCount + 1
            Else
                skipCount = skipCount + 1
            End If

            ' 3行目日付ヘッダ (全OHLCVシートで同期)
            ws.Cells(3, col).Value = dateVals(d)
            ws.Cells(3, col).NumberFormat = "m/d"
        Next d

NextSheet:
    Next s

    Application.ScreenUpdating = True

    ' --- 完了メッセージ + クリア確認 ---
    MsgBox "取込完了！" & vbCrLf & vbCrLf & _
           "銘柄: " & codeInput & " " & meiName & vbCrLf & _
           "取込件数: " & dateCount & " 日分" & vbCrLf & _
           "日付範囲: " & Format(dateVals(0), "m/d") & " ～ " & _
           Format(dateVals(dateCount - 1), "m/d") & vbCrLf & _
           "書込みセル数: " & writeCount & vbCrLf & _
           "スキップ(値0): " & skipCount, _
           vbOKOnly + vbInformation, "取込完了"

    ' ★「データ取込シートをクリアしますか？」は廃止した。
    '   A2:Z500 は 2行目のヘッダ（銘柄名称/市場名称/日付/…）と
    '   A3 の RssChartPast 数式まで含むので、消すと以後どのマクロも
    '   取り込めなくなる。「1行目のヘッダは残してあります」という
    '   案内は誤りだった（ヘッダは2行目にある）。
End Sub

' ================================================================
' データ取込シート初期化 (ヘッダ行だけ残す)
' ================================================================
Public Sub データ取込_クリア()
    Dim dataWs As Worksheet
    On Error Resume Next
    Set dataWs = ThisWorkbook.Sheets("データ取込")
    On Error GoTo 0
    If dataWs Is Nothing Then Exit Sub

    ' ★A2:Z500 を消すと 2行目のヘッダと A3 の RssChartPast 数式まで
    '   消えてしまい、以後どのマクロも取り込めなくなる。
    '   このシートは B1 の銘柄コードで動くので、B1 を空にすれば表示は消える。
    dataWs.Range("B1").ClearContents
    Application.Calculate
    MsgBox "銘柄コード(B1)を空にしました。" & vbCrLf & _
           "2行目のヘッダと A3 の数式は残してあります。", vbInformation, "クリア完了"
End Sub