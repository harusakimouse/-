Attribute VB_Name = "Mod_HeikinAshi_Date"
Option Explicit
'==================================================================
' 平均足　日付確認　v1.0
'
'   実行 : Alt+F8 →「平均足_日付確認」
'
'   このブックの「終値」シートと、元のOHLCVブックの「終値」シートの
'   日付行（D3から右へ8個）を並べて表示します。
'   どちらが古いのかが一目で分かります。
'==================================================================

Public Sub 平均足_日付確認()

    Dim s As String
    s = "◆このブック（平均足研究）の 終値シート" & vbCrLf
    s = s & HA_D_Line(ThisWorkbook) & vbCrLf & vbCrLf

    Dim wb As Workbook, src As Workbook
    For Each wb In Application.Workbooks
        If wb.Name Like "OHLCV*" Then Set src = wb
    Next wb

    If src Is Nothing Then
        s = s & "◆OHLCVブック" & vbCrLf & "　いま開いていません。" & vbCrLf & _
            "　①「取込＋買い候補」を押せば、自動で開いて取り込みます。"
    Else
        s = s & "◆" & src.Name & vbCrLf & HA_D_Line(src)
    End If

    MsgBox s, vbInformation, "日付の確認"

End Sub

'--- 1ブック分の日付を文字にする ---
Private Function HA_D_Line(ByVal wb As Workbook) As String

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("終値")
    On Error GoTo 0
    If ws Is Nothing Then
        HA_D_Line = "　「終値」シートがありません。"
        Exit Function
    End If

    Dim c As Long, t As String, v As Variant
    t = ""
    For c = 4 To 11
        v = ws.Cells(3, c).Value2
        If IsNumeric(v) Then
            If CDbl(v) > 40000 And CDbl(v) < 80000 Then
                t = t & Format(CDate(CDbl(v)), "m/d") & " / "
            Else
                t = t & "?" & " / "
            End If
        ElseIf Len(CStr(v)) = 0 Then
            t = t & "空 / "
        Else
            t = t & "文字 / "
        End If
    Next c

    Dim p As String
    p = CStr(ws.Cells(6, 5).Value2)
    If Len(p) = 0 Then p = "空"

    HA_D_Line = "　D3から右へ： " & t & vbCrLf & _
                "　E6（一番新しい日の値）： " & p

End Function
