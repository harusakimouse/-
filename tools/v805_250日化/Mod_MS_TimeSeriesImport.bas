Attribute VB_Name = "Mod_MS_TimeSeriesImport"
Option Explicit

' ================================================================
' マーケットスピード時系列データ取込
' データ取込シートに MS時系列をコピペしたものを OHLCV 5シートに展開
' ================================================================

Private Const PWD As String = "ne19480314"

Public Sub マーケットスピード過去データ取込()
    Dim dataWs As Worksheet
    On Error Resume Next
    Set dataWs = ThisWorkbook.Sheets("データ取込")
    On Error GoTo 0
    If dataWs Is Nothing Then
        MsgBox "「データ取込」シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    ' --- B1 から銘柄コード取得 ---
    Dim codeInput As String
    codeInput = Trim(CStr(dataWs.Cells(1, 2).Value))
    If codeInput = "" Then
        MsgBox "データ取込シートの B1 に銘柄コードが入力されていません。" & vbCrLf & _
               "例: 285A, 7203, 8306 など", _
               vbExclamation
        Exit Sub
    End If
    codeInput = UCase(codeInput)

    ' --- 銘柄管理シートで対象行を検索 ---
    Dim mws As Worksheet
    On Error Resume Next
    Set mws = ThisWorkbook.Sheets("銘柄管理")
    On Error GoTo 0
    If mws Is Nothing Then
        MsgBox "「銘柄管理」シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim targetRow As Long: targetRow = 0
    Dim r As Long
    For r = 6 To 505
        If UCase(Trim(CStr(mws.Cells(r, 2).Value))) = codeInput Then
            targetRow = r
            Exit For
        End If
    Next r
    If targetRow = 0 Then
        MsgBox "銘柄コード " & codeInput & " が銘柄管理に見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim meiName As String: meiName = CStr(mws.Cells(targetRow, 3).Value)

    ' --- ヘッダ行を自動検出 (「日付」「始値」「終値」を含む行) ---
    Dim headerRow As Long: headerRow = 0
    Dim hr As Long, hc As Long
    Dim hasDate As Boolean, hasOpen As Boolean, hasClose As Boolean
    For hr = 1 To 50
        hasDate = False: hasOpen = False: hasClose = False
        For hc = 1 To 30
            Dim hv As String: hv = Trim(CStr(dataWs.Cells(hr, hc).Value))
            If hv = "日付" Then hasDate = True
            If hv = "始値" Then hasOpen = True
            If hv = "終値" Then hasClose = True
        Next hc
        If hasDate And hasOpen And hasClose Then
            headerRow = hr
            Exit For
        End If
    Next hr
    If headerRow = 0 Then
        MsgBox "ヘッダ行が見つかりません。" & vbCrLf & _
               "「日付」「始値」「終値」を含む行が必要です。", _
               vbExclamation
        Exit Sub
    End If

    ' --- ヘッダから列番号を特定 ---
    Dim dateCol As Long: dateCol = 0
    Dim openCol As Long: openCol = 0
    Dim highCol As Long: highCol = 0
    Dim lowCol As Long:  lowCol = 0
    Dim closeCol As Long: closeCol = 0
    Dim volCol As Long: volCol = 0
    Dim cc As Long
    For cc = 1 To 30
        Dim cv As String: cv = Trim(CStr(dataWs.Cells(headerRow, cc).Value))
        Select Case cv
            Case "日付":   dateCol = cc
            Case "始値":   openCol = cc
            Case "高値":   highCol = cc
            Case "安値":   lowCol = cc
            Case "終値":   closeCol = cc
            Case "出来高", "売買高", "出来高(株)", "出来高（株）"
                volCol = cc
        End Select
    Next cc

    If dateCol = 0 Or openCol = 0 Or highCol = 0 Or lowCol = 0 Or closeCol = 0 Or volCol = 0 Then
        MsgBox "必要な列が見つかりません。" & vbCrLf & _
               "日付:" & dateCol & " 始値:" & openCol & " 高値:" & highCol & vbCrLf & _
               "安値:" & lowCol & " 終値:" & closeCol & " 出来高:" & volCol, _
               vbExclamation
        Exit Sub
    End If

    ' --- データ行を収集 (ヘッダ行+1 から) ---
    Dim dateRows() As Long
    Dim dateVals() As Date
    ReDim dateRows(1 To 400)
    ReDim dateVals(1 To 400)
    Dim cnt As Long: cnt = 0

    Dim lastDataRow As Long
    lastDataRow = dataWs.Cells(dataWs.Rows.Count, dateCol).End(xlUp).Row

    Dim dr As Long
    For dr = headerRow + 1 To lastDataRow
        Dim dv As Variant: dv = dataWs.Cells(dr, dateCol).Value
        If IsDate(dv) Then
            cnt = cnt + 1
            dateRows(cnt) = dr
            dateVals(cnt) = CDate(dv)
        End If
    Next dr

    If cnt = 0 Then
        MsgBox "データ行が見つかりません。" & vbCrLf & _
               "ヘッダ行 " & headerRow & " の下に有効な日付データがありません。", _
               vbExclamation
        Exit Sub
    End If

    ' --- 日付で降順ソート (バブルソート) ---
    Dim i As Long, j As Long, tmpD As Date, tmpR As Long
    For i = 1 To cnt - 1
        For j = 1 To cnt - i
            If dateVals(j) < dateVals(j + 1) Then
                tmpD = dateVals(j): dateVals(j) = dateVals(j + 1): dateVals(j + 1) = tmpD
                tmpR = dateRows(j): dateRows(j) = dateRows(j + 1): dateRows(j + 1) = tmpR
            End If
        Next j
    Next i

    ' --- 確認ダイアログ ---
    Dim ret As Integer
    ret = MsgBox("以下の内容で OHLCV 5シートに取込みます。" & vbCrLf & vbCrLf & _
                 "銘柄: " & codeInput & " " & meiName & vbCrLf & _
                 "取込先行: " & targetRow & " 行目" & vbCrLf & _
                 "ヘッダ検出: " & headerRow & " 行目" & vbCrLf & _
                 "取込件数: " & cnt & " 日分" & vbCrLf & _
                 "日付範囲: " & Format(dateVals(1), "m/d") & " ～ " & _
                 Format(dateVals(cnt), "m/d") & vbCrLf & vbCrLf & _
                 "OHLCV 5シートの " & targetRow & "行目 E列～IT列（250日分）を上書きします。" & vbCrLf & _
                 "3行目の日付ヘッダも更新されます。" & vbCrLf & vbCrLf & _
                 "進めますか？", _
                 vbYesNo + vbQuestion, "MS取込確認")
    If ret = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

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
        If ws Is Nothing Then GoTo NextSh

        On Error Resume Next
        ws.Unprotect Password:=PWD
        On Error GoTo 0

        ' E～IT (col 5～254) をクリア
        ws.Range(ws.Cells(targetRow, 5), ws.Cells(targetRow, 254)).ClearContents

        ' 書込 (E列=最新, F列=1日前, ... IT列=249日前)
        Dim d As Long
        For d = 1 To cnt
            Dim col As Long: col = 4 + d   ' E=5, F=6, ...
            If col > 254 Then Exit For

            Dim srcR As Long: srcR = dateRows(d)
            Dim srcV As Variant: srcV = dataWs.Cells(srcR, srcCols(s)).Value

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

            ' 3行目 日付ヘッダも更新
            ws.Cells(3, col).Value = dateVals(d)
            ws.Cells(3, col).NumberFormat = "m/d"
        Next d

        On Error Resume Next
        ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
                   DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                   AllowFiltering:=True, AllowSorting:=True
        On Error GoTo 0
NextSh:
    Next s

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "MS時系列取込 完了！" & vbCrLf & vbCrLf & _
           "銘柄: " & codeInput & " " & meiName & vbCrLf & _
           "取込件数: " & cnt & " 日分" & vbCrLf & _
           "日付範囲: " & Format(dateVals(1), "m/d") & " ～ " & _
           Format(dateVals(cnt), "m/d") & vbCrLf & _
           "書込みセル数: " & writeCount & vbCrLf & _
           "スキップ(値0): " & skipCount, _
           vbOKOnly + vbInformation, "完了"

    ' ★「データ取込シートをクリアしますか？」は廃止した。
    '   クリアは Range("A2:Z500").ClearContents で、2行目のヘッダ
    '   （銘柄名称/市場名称/日付/…）と A3 の RssChartPast 数式まで
    '   消してしまい、以後どのマクロも取り込めなくなる。
    '   RSS は B1 を書き換えれば自動で更新されるので、そもそも
    '   クリアする必要がない。
End Sub