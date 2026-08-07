Attribute VB_Name = "Mod_再計算ボタン"
Option Explicit

' ==============================================================================
'  再計算ボタン
'
'   ボタンを設置する   … 厳選TOP2 シートに「再計算」ボタンを描く（1回だけ実行）
'   モード切替ボタンを設置 … 「大引け／場中」を切り替える大きなボタンを描く
'   厳選TOP2_再計算    … ボタンから呼ばれる。押すと計算し直して結果を表示する
'   厳選TOP2_再計算_静か … 日次更新から呼ばれる。ダイアログを出さず、
'                        結果を厳選TOP2 の H7 セルに書くだけ
'
'  F9 キーの代わりです。ボタンを押せば同じことをします。
' ==============================================================================

Private Const SH_TOP   As String = "厳選TOP2"
Private Const BTN_NAME  As String = "btn再計算"
Private Const MODE_NAME As String = "btnモード"
Private Const MODE_CELL As String = "H10"
Private Const PWD      As String = "ne19480314"


'==============================================================================
' ① ボタンを設置する（最初に1回だけ実行してください）
'==============================================================================
Public Sub 再計算ボタンを設置()
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_TOP)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "「" & SH_TOP & "」シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim wasP As Boolean: wasP = ws.ProtectContents
    On Error Resume Next
    ws.Unprotect Password:=PWD
    ws.Unprotect
    Err.Clear
    On Error GoTo 0

    ' 同じボタンが既にあれば消してから作り直す（押すたびに増えないように）
    Dim sh As Shape, i As Long
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If sh.Name = BTN_NAME Then sh.Delete
    Next i

    ' G4 のあたりに置く
    Dim tl As Range: Set tl = ws.Range("H4")
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                tl.Left, tl.Top, 132, 40)
    With sh
        .Name = BTN_NAME
        .Fill.ForeColor.RGB = RGB(0, 112, 192)
        .Line.ForeColor.RGB = RGB(0, 70, 130)
        .Line.Weight = 1.25
        With .TextFrame2
            .TextRange.Text = "再計算"
            .TextRange.Font.Size = 14
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .MarginLeft = 0: .MarginRight = 0
        End With
        .OnAction = "厳選TOP2_再計算"
    End With

    If wasP Then
        On Error Resume Next
        ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
                   DrawingObjects:=False, Contents:=True, Scenarios:=True
        On Error GoTo 0
    End If

    ws.Activate
    MsgBox "「再計算」ボタンを " & SH_TOP & " シートに置きました。" & vbCrLf & vbCrLf & _
           "以後はこのボタンを押してください。F9 と同じことをします。" & vbCrLf & vbCrLf & _
           "位置を変えたいときは、ボタンを右クリックしてドラッグしてください。", _
           vbInformation, "設置完了"
End Sub


'==============================================================================
' ② 再計算（ボタンから呼ばれる）
'==============================================================================
Public Sub 厳選TOP2_再計算()
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation

    On Error GoTo ErrHandler
    Application.EnableCancelKey = xlErrorHandler

    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_TOP)
    On Error GoTo ErrHandler
    If ws Is Nothing Then Err.Raise 5, , "「" & SH_TOP & "」シートが見つかりません。"

    Application.StatusBar = "再計算しています…": DoEvents

    ' RSS は自動計算でないと値が更新されない
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild
    DoEvents

    ' RSS の応答が届くまで少し待ってから、もう一度
    Dim t0 As Double: t0 = Timer
    Do While Timer - t0 < 2
        DoEvents
    Loop
    Application.Calculate
    DoEvents

    Dim nBuy As Long, nSell As Long
    nBuy = Val(CStr(ws.Range("F5").Value))
    nSell = Val(CStr(ws.Range("F6").Value))

    Dim now2 As Date: now2 = Now
    Application.Calculation = prevCalc
    Application.StatusBar = False

    Dim msg As String
    msg = "再計算しました。（" & Format$(now2, "m/d hh:nn") & "）" & vbCrLf & _
          String(30, "-") & vbCrLf & _
          "  買い 条件通過: " & nBuy & " 銘柄" & vbCrLf & _
          "  売り 条件通過: " & nSell & " 銘柄" & vbCrLf & vbCrLf

    If nBuy = 0 And nSell = 0 Then
        msg = msg & "該当なしです。" & vbCrLf & vbCrLf & _
              "・15:30 の大引け前は、出来高が1日ぶん溜まっていないので" & vbCrLf & _
              "  ほぼ0件になります。15:35 以降に押してください。" & vbCrLf & _
              "・大引け後でも0件なら、今日は見送りが正しい判断です。"
    Else
        msg = msg & "候補は 16行目以降（買い）と 24行目以降（売り）に出ています。"
    End If

    On Error Resume Next
    モードボタン表示を更新
    Err.Clear
    On Error GoTo 0

    MsgBox msg, vbInformation, "再計算"
    Exit Sub

ErrHandler:
    If Err.Number = 18 Then Resume Next
    Dim e As String: e = "Err " & Err.Number & ": " & Err.Description
    On Error Resume Next
    Application.Calculation = prevCalc
    Application.StatusBar = False
    On Error GoTo 0
    MsgBox "再計算を中断しました。" & vbCrLf & vbCrLf & e, vbCritical, "エラー"
End Sub


'==============================================================================
' ③ 静かに再計算する（日次更新から呼ばれる。ダイアログを出さない）
'
'   15:35 / 15:50 / 16:10 の自動実行から呼ばれるので、
'   MsgBox を出すと画面が止まってしまう。結果はセルに書く。
'==============================================================================
Public Sub 厳選TOP2_再計算_静か()
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    On Error GoTo Quiet

    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_TOP)
    On Error GoTo Quiet
    If ws Is Nothing Then GoTo Quiet

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild
    DoEvents
    Dim t0 As Double: t0 = Timer
    Do While Timer - t0 < 3
        DoEvents
    Loop
    Application.Calculate
    DoEvents

    Dim nBuy As Long, nSell As Long
    nBuy = Val(CStr(ws.Range("F5").Value))
    nSell = Val(CStr(ws.Range("F6").Value))

    Dim wasP As Boolean: wasP = ws.ProtectContents
    On Error Resume Next
    ws.Unprotect Password:=PWD
    ws.Unprotect
    Err.Clear
    On Error GoTo Quiet

    With ws.Range("H7")
        .Value = "最終更新 " & Format$(Now, "m/d hh:nn") & _
                 "   買い " & nBuy & " 件 / 売り " & nSell & " 件"
        .Font.Bold = True
        If nBuy > 0 Then
            .Font.Color = RGB(0, 112, 192)
        Else
            .Font.Color = RGB(120, 120, 120)
        End If
    End With

    If wasP Then
        On Error Resume Next
        ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
                   DrawingObjects:=False, Contents:=True, Scenarios:=True
        On Error GoTo Quiet
    End If

Quiet:
    On Error Resume Next
    Application.Calculation = prevCalc
    On Error GoTo 0
End Sub


'==============================================================================
' ④ モード切替ボタンを設置する（最初に1回だけ実行してください）
'
'   「大引け」と「場中」を押すたびに切り替えます。
'   色と文字で、いまどちらなのかが常に見えます。
'==============================================================================
Public Sub モード切替ボタンを設置()
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_TOP)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "「" & SH_TOP & "」シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim wasP As Boolean: wasP = ws.ProtectContents
    On Error Resume Next
    ws.Unprotect Password:=PWD
    ws.Unprotect
    Err.Clear
    On Error GoTo 0

    Dim sh As Shape, i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Name = MODE_NAME Then ws.Shapes(i).Delete
    Next i

    ' 「再計算」ボタン(H4)の右どなり。再計算ボタンの約2倍の大きさ
    Dim tl As Range: Set tl = ws.Range("K4")
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, tl.Left, tl.Top, 224, 62)
    With sh
        .Name = MODE_NAME
        .Line.Weight = 2
        With .TextFrame2
            .TextRange.Font.Size = 18
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .MarginLeft = 0: .MarginRight = 0
            .MarginTop = 0: .MarginBottom = 0
        End With
        .OnAction = "モードを切り替える"
    End With

    モードボタン表示を更新

    If wasP Then
        On Error Resume Next
        ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
                   DrawingObjects:=False, Contents:=True, Scenarios:=True
        On Error GoTo 0
    End If

    ws.Activate
    MsgBox "モード切替ボタンを置きました。" & vbCrLf & vbCrLf & _
           "押すたびに「大引け」⇔「場中」が切り替わります。" & vbCrLf & vbCrLf & _
           "  青  = 大引けモード（15:35以降に使う）" & vbCrLf & _
           "  橙  = 場中モード（15:00頃に使う。しきい値を下げる）", _
           vbInformation, "設置完了"
End Sub


'==============================================================================
' ⑤ モードを切り替える（ボタンから呼ばれる）
'==============================================================================
Public Sub モードを切り替える()
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SH_TOP)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim wasP As Boolean: wasP = ws.ProtectContents
    On Error Resume Next
    ws.Unprotect Password:=PWD
    ws.Unprotect
    Err.Clear
    On Error GoTo 0

    If Trim$(CStr(ws.Range(MODE_CELL).Value)) = "場中" Then
        ws.Range(MODE_CELL).Value = "大引け"
    Else
        ws.Range(MODE_CELL).Value = "場中"
    End If

    If wasP Then
        On Error Resume Next
        ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
                   DrawingObjects:=False, Contents:=True, Scenarios:=True
        On Error GoTo 0
    End If

    モードボタン表示を更新
    厳選TOP2_再計算
End Sub


'==============================================================================
' ⑥ ボタンの文字と色を、いまのモードに合わせる
'==============================================================================
Public Sub モードボタン表示を更新()
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_TOP)
    If ws Is Nothing Then Exit Sub

    Dim sh As Shape: Set sh = Nothing
    Dim i As Long
    For i = 1 To ws.Shapes.Count
        If ws.Shapes(i).Name = MODE_NAME Then Set sh = ws.Shapes(i): Exit For
    Next i
    If sh Is Nothing Then Exit Sub

    Dim isNaka As Boolean
    isNaka = (Trim$(CStr(ws.Range(MODE_CELL).Value)) = "場中")

    Dim sc As String, vo As String
    sc = CStr(ws.Range("H12").Value)
    vo = Format$(ws.Range("H13").Value, "0.0")

    With sh
        If isNaka Then
            .Fill.ForeColor.RGB = RGB(230, 120, 0)      ' 橙
            .Line.ForeColor.RGB = RGB(170, 85, 0)
            .TextFrame2.TextRange.Text = "場中モード" & vbCrLf & _
                "スコア" & sc & "以上 / 出来高" & vo & "倍以上"
        Else
            .Fill.ForeColor.RGB = RGB(0, 70, 130)       ' 濃紺
            .Line.ForeColor.RGB = RGB(0, 40, 80)
            .TextFrame2.TextRange.Text = "大引けモード" & vbCrLf & _
                "スコア" & sc & "以上 / 出来高" & vo & "倍以上"
        End If
        With .TextFrame2.TextRange
            .Font.Size = 18
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            ' 2行目は小さく
            .Paragraphs(2).Font.Size = 11
            .Paragraphs(2).Font.Bold = msoFalse
        End With
    End With
    Err.Clear
End Sub
