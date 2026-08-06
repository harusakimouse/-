Attribute VB_Name = "RSS過去データ取り込み"
Sub ImportOHLCV()
    Dim wsData  As Worksheet
    Dim wsOpen  As Worksheet, wsHigh As Worksheet
    Dim wsLow   As Worksheet, wsClose As Worksheet, wsVol As Worksheet
    Set wsData = Sheets("データ取込")
    Set wsOpen = Sheets("始値")
    Set wsHigh = Sheets("高値")
    Set wsLow = Sheets("安値")
    Set wsClose = Sheets("終値")
    Set wsVol = Sheets("出来高")     ' ← 末尾のAを削除

    Dim targetCode As String
    targetCode = Trim(CStr(wsData.Cells(1, 2).Value))
    If targetCode = "" Then
        MsgBox "B1にコードを入力してください。", vbExclamation
        Exit Sub
    End If

    Dim colDate  As Integer: colDate = 0
    Dim colVol   As Integer: colVol = 0
    Dim colOpen  As Integer: colOpen = 0
    Dim colHigh  As Integer: colHigh = 0
    Dim colLow   As Integer: colLow = 0
    Dim colClose As Integer: colClose = 0
    Dim hc As Integer
    For hc = 1 To 20
        Select Case CStr(wsData.Cells(2, hc).Value)
            Case "日付":   colDate = hc
            Case "出来高": colVol = hc
            Case "始値":   colOpen = hc
            Case "高値":   colHigh = hc
            Case "安値":   colLow = hc
            Case "終値":   colClose = hc
        End Select
    Next hc
    If colDate = 0 Or colVol = 0 Or colOpen = 0 Or _
       colHigh = 0 Or colLow = 0 Or colClose = 0 Then
        MsgBox "ヘッダー行2に必要な列が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' ★修正: 検索範囲を245→500に拡張（銘柄が増えても対応）
    Dim codeRow As Long: codeRow = 0
    Dim i As Long
    For i = 6 To 500
        If Trim(CStr(wsVol.Cells(i, 1).Value)) = targetCode Then
            codeRow = i
            Exit For
        End If
    Next i
    If codeRow = 0 Then
        MsgBox "コード「" & targetCode & "」がOHLCVシートに見つかりません。" & vbCrLf & _
               "（検索範囲: 出来高シート A6:A500）", vbExclamation
        Exit Sub
    End If

    Dim dateToCol(50000) As Long
    Dim c As Long
    For c = 5 To 254
        Dim dv As Variant
        dv = wsVol.Cells(3, c).Value
        If dv <> "" And Not IsEmpty(dv) Then
            Dim dvSer As Long
            dvSer = 0
            On Error Resume Next
            dvSer = CLng(CDbl(dv))
            On Error GoTo 0
            If dvSer > 40000 And dvSer < 50000 Then
                dateToCol(dvSer) = c
            End If
        End If
    Next c

    Dim written As Long: written = 0
    Dim skipped As Long: skipped = 0
    Dim r As Long
    For r = 3 To 300
        Dim cellVal As Variant
        cellVal = wsData.Cells(r, colDate).Value
        If cellVal = "" Or IsEmpty(cellVal) Then Exit For
        Dim dateSer As Long: dateSer = 0
        On Error Resume Next
        dateSer = CLng(CDbl(CDate(CStr(cellVal))))
        On Error GoTo 0
        Dim targetCol As Long: targetCol = 0
        If dateSer > 40000 And dateSer < 50000 Then
            targetCol = dateToCol(dateSer)
        End If
        If targetCol > 0 Then
            wsVol.Cells(codeRow, targetCol).Value = wsData.Cells(r, colVol).Value
            wsOpen.Cells(codeRow, targetCol).Value = wsData.Cells(r, colOpen).Value
            wsHigh.Cells(codeRow, targetCol).Value = wsData.Cells(r, colHigh).Value
            wsLow.Cells(codeRow, targetCol).Value = wsData.Cells(r, colLow).Value
            wsClose.Cells(codeRow, targetCol).Value = wsData.Cells(r, colClose).Value
            written = written + 1
        Else
            skipped = skipped + 1
        End If
    Next r

    MsgBox "【" & targetCode & "】完了" & vbCrLf & _
           "書込: " & written & " 日分" & vbCrLf & _
           "スキップ（日付範囲外）: " & skipped & " 日分", vbInformation
End Sub

Sub AA()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("データ取込")
    Dim rng As Range
    Set rng = ws.Range("A3")
    If rng.HasFormula Then
        Dim strFormula As String
        strFormula = rng.Formula
        rng.Formula = strFormula
    End If
    Application.Calculate
    DoEvents
    Application.Wait Now + TimeValue("00:00:03")
End Sub