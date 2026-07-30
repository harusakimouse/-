Attribute VB_Name = "Module4"
' ==================================================================
' ★ Module4（完全版 v2）
'   ・メニューに「サマリー更新／買銘柄更新／売銘柄更新」ボタンを設置
'   ・3シートは数式駆動なので「更新＝全再計算」でOK
'   ★ v2の改善点：更新ボタンが「効いたか分からない」問題を解消
'      - 更新のたびに確認ダイアログ（MsgBox）を表示
'      - サマリーの現在の判定内容（TOPIX比率・該当条件）も一緒に表示
'      - 集計シートは何件表示されているかを表示
'   使い方：貼り付け後、マクロ「メニューに更新ボタン設置」を1回だけ実行
' ==================================================================
Option Explicit

' ===== ボタンの動作 =====
Public Sub 更新_サマリー()
    全再計算
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("サマリー")
    ws.Activate

    Dim topixRate As String, uriCond As String, kaiCond As String
    topixRate = FmtPct(ws.Range("E4").Value)
    uriCond = CStr(ws.Range("C6").Value)
    kaiCond = CStr(ws.Range("E6").Value)

    Application.StatusBar = Format(Now, "hh:nn:ss") & " サマリー更新しました"
    MsgBox "サマリーを更新しました。" & vbCrLf & vbCrLf & _
           "更新時刻　　　: " & Format(Now, "yyyy/mm/dd hh:nn:ss") & vbCrLf & _
           "TOPIX前日比率: " & topixRate & vbCrLf & _
           "売り該当条件　: " & uriCond & vbCrLf & _
           "買い該当条件　: " & kaiCond, _
           vbInformation, "更新完了（サマリー）"
End Sub

Public Sub 更新_買銘柄()
    全再計算
    NotifyList "全体買銘柄"
End Sub

Public Sub 更新_売銘柄()
    全再計算
    NotifyList "全体売銘柄"
End Sub

Private Sub 全再計算()
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild
    DoEvents
End Sub

' 集計シートを更新し、表示件数つきで確認ダイアログを出す
Private Sub NotifyList(ByVal shName As String)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(shName)
    ws.Activate
    ' B列（コード）に値が入っている件数を数える（ヘッダー除く）
    Dim cnt As Long, r As Long, lastr As Long
    lastr = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = 2 To lastr
        If Len(Trim(CStr(ws.Cells(r, 2).Value))) > 0 Then cnt = cnt + 1
    Next r
    Application.StatusBar = Format(Now, "hh:nn:ss") & " " & shName & " 更新しました"
    MsgBox shName & " を更新しました。" & vbCrLf & vbCrLf & _
           "更新時刻: " & Format(Now, "yyyy/mm/dd hh:nn:ss") & vbCrLf & _
           "表示銘柄数: " & cnt & " 件", _
           vbInformation, "更新完了（" & shName & "）"
End Sub

' 前日比率セル（%値/小数どちらでも）を "＋0.79%" 形式に整える
Private Function FmtPct(ByVal v As Variant) As String
    On Error Resume Next
    If Not IsNumeric(v) Then FmtPct = CStr(v): Exit Function
    Dim d As Double: d = CDbl(v)
    ' |値|<=1 なら小数（0.0079）とみなして%へ、それ以外は既に%値とみなす
    If Abs(d) <= 1 Then d = d * 100
    FmtPct = Format(d, "+0.00;-0.00") & "%"
End Function

' ===== メニューに3ボタンを設置（1回だけ実行）=====
Public Sub メニューに更新ボタン設置()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("メニュー")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "「メニュー」シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    ' 既存の同名ボタンを消してから作り直し（重複防止）
    Dim nm As Variant
    For Each nm In Array("btnUpdSummary", "btnUpdBuy", "btnUpdSell", "lblUpdSection")
        On Error Resume Next
        ws.Shapes(CStr(nm)).Delete
        On Error GoTo 0
    Next nm

    ' 設置する行（既存レイアウトの下＝13行目付近）。重なる場合は下の baseRow を変更
    Dim baseRow As Long: baseRow = 13
    Dim topBtn As Double: topBtn = ws.Rows(baseRow).Top + 5

    ' 見出しラベル
    Dim lbl As Shape
    Set lbl = ws.Shapes.AddShape(msoShapeRectangle, 40, ws.Rows(baseRow - 1).Top + 3, 400, 22)
    lbl.Name = "lblUpdSection"
    lbl.Fill.ForeColor.RGB = RGB(230, 230, 230)
    lbl.Line.Visible = msoFalse
    With lbl.TextFrame2.TextRange
        .Text = "【 表示更新（再計算）】"
        .Font.Bold = msoTrue
        .Font.Size = 11
        .Font.Fill.ForeColor.RGB = RGB(60, 60, 60)
    End With
    lbl.TextFrame2.HorizontalAnchor = msoAnchorCenter
    lbl.TextFrame2.VerticalAnchor = msoAnchorMiddle

    ' 3ボタン
    AddUpdBtn ws, "btnUpdSummary", "サマリー更新", "更新_サマリー", 40, topBtn, RGB(0, 150, 80)
    AddUpdBtn ws, "btnUpdBuy", "買銘柄更新", "更新_買銘柄", 180, topBtn, RGB(0, 150, 80)
    AddUpdBtn ws, "btnUpdSell", "売銘柄更新", "更新_売銘柄", 320, topBtn, RGB(0, 150, 80)

    ws.Activate
    MsgBox "メニューに更新ボタンを設置しました。", vbInformation
End Sub

Private Sub AddUpdBtn(ws As Worksheet, shpName As String, caption As String, _
                      macroName As String, leftPos As Double, topPos As Double, colorRGB As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, 120, 35)
    shp.Name = shpName
    shp.Fill.ForeColor.RGB = colorRGB
    shp.Line.ForeColor.RGB = RGB(80, 80, 80)
    shp.Line.Weight = 1
    With shp.TextFrame2
        .WordWrap = msoFalse
        .HorizontalAnchor = msoAnchorCenter
        .VerticalAnchor = msoAnchorMiddle
        With .TextRange
            .Text = caption
            .Font.Bold = msoTrue
            .Font.Size = 11
            .Font.Fill.ForeColor.RGB = vbWhite
        End With
    End With
    ' 確実に解決するようブック名で修飾
    shp.OnAction = "'" & ThisWorkbook.Name & "'!" & macroName
End Sub

Public Sub ボタン再登録()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("メニュー")
    Dim m As Variant, cnt As Long, ng As String
    For Each m In Array( _
        Array("15:00取得", "Auto1500"), Array("15:20取得", "Auto1520"), Array("15:30取得", "Auto1530"), _
        Array("1500蓄積", "Auto1500"), Array("1520蓄積", "Auto1520"), Array("1530蓄積", "Auto1530"), _
        Array("btnUpdSummary", "更新_サマリー"), Array("btnUpdBuy", "更新_買銘柄"), Array("btnUpdSell", "更新_売銘柄"))
        On Error Resume Next
        Err.Clear
        ws.Shapes(CStr(m(0))).OnAction = "'" & ThisWorkbook.Name & "'!" & CStr(m(1))
        If Err.Number = 0 Then
            cnt = cnt + 1
        Else
            ng = ng & vbCrLf & "  " & CStr(m(0))
        End If
        On Error GoTo 0
    Next m
    MsgBox cnt & " 個のボタンを再登録しました。" & _
           IIf(ng <> "", vbCrLf & "見つからなかった図形:" & ng, ""), vbInformation
End Sub
