Attribute VB_Name = "Mod_空行ガード適用"
Option Explicit

' ==================================================================
' 管理 / 銘柄管理 シートの「空コード行に対する RSS 無駄打ち」を止める
'
' 【現状（V805 実測）】
'   管理シート    : RSS数式のある行 458行(4〜461) のうち 415行 が
'                   B列(コード) 空。1行4本なので 約1,660本 が空打ち。
'   銘柄管理シート: RSS数式のある行 488行(5〜505) のうち 200行(306〜505) が
'                   B列 空。
'
'   分析 / 売分析 には  =IF(C3<>"",RssMarket(C3,"現在値"),"")  という
'   ガードが入っているのに、管理 / 銘柄管理 には入っていない。
'
' 【この処理】
'   コード参照が空のときに "" を返すよう、既存の RSS 数式を
'   IF(参照="","",元の式) で包み直す。式の中身（項目名・参照）は変えない。
'   既に IF( で始まる式は二重に包まない。
'
'   ※ 数式そのものを書き換えるので、必ずバックアップを取ってから。
' ==================================================================

Private Const PWD As String = "ne19480314"


Public Sub 空行ガード_適用()

    If MsgBox("管理 / 銘柄管理 シートの RSS 数式を" & vbCrLf & _
              "  =IF(<コード参照>="""","""",<元の式>)" & vbCrLf & _
              "の形に包み直します。" & vbCrLf & vbCrLf & _
              "★ 必ず事前にブックのバックアップを取ってください ★" & vbCrLf & vbCrLf & _
              "続行しますか?", vbYesNo + vbExclamation, "空行ガード適用") <> vbYes Then Exit Sub

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Dim prevCalc As XlCalculation: prevCalc = Application.Calculation
    Application.Calculation = xlCalculationManual

    Dim report As String, total As Long

    total = total + WrapSheet("管理", 1, 4, 461, "B")     ' コードは B列
    report = report & "管理: " & total & " セル" & vbCrLf

    Dim n2 As Long
    n2 = WrapSheet("銘柄管理", 1, 5, 505, "B")            ' コードは B列
    report = report & "銘柄管理: " & n2 & " セル" & vbCrLf
    total = total + n2

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.Calculate

    MsgBox "空行ガードを適用しました。" & vbCrLf & String(30, "-") & vbCrLf & _
           report & String(30, "-") & vbCrLf & "合計 " & total & " セル", _
           vbInformation, "完了"
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "エラー: " & Err.Number & " " & Err.Description, vbCritical, "エラー"
End Sub


' ------------------------------------------------------------------
' 1 シート分。RSS 数式を持つセルを IF() で包む。包んだ数を返す。
' ------------------------------------------------------------------
Private Function WrapSheet(ByVal shName As String, _
                           ByVal colFirst As Long, _
                           ByVal rowFirst As Long, _
                           ByVal rowLast As Long, _
                           ByVal codeCol As String) As Long
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    ws.Unprotect Password:=PWD
    On Error GoTo 0

    Dim colLast As Long
    colLast = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1

    Dim n As Long, r As Long, c As Long
    For r = rowFirst To rowLast
        For c = colFirst To colLast
            Dim cell As Range
            Set cell = ws.Cells(r, c)
            If cell.HasFormula Then
                Dim f As String: f = cell.Formula
                If InStr(1, f, "Rss", vbTextCompare) > 0 Then
                    ' 既にガード済み / 指数系はスキップ
                    If Left$(UCase$(Replace(f, " ", "")), 4) <> "=IF(" _
                       And InStr(1, f, "RssIndexMarket", vbTextCompare) = 0 Then
                        Dim guard As String
                        guard = "=IF($" & codeCol & r & "="""","""","
                        Dim newF As String
                        newF = guard & Mid$(f, 2) & ")"
                        On Error Resume Next
                        Err.Clear
                        cell.Formula = newF
                        If Err.Number = 0 Then n = n + 1
                        Err.Clear
                        On Error GoTo 0
                    End If
                End If
            End If
        Next c
    Next r

    On Error Resume Next
    ws.Protect Password:=PWD, UserInterfaceOnly:=True, _
               DrawingObjects:=True, Contents:=True, Scenarios:=True
    On Error GoTo 0

    WrapSheet = n
End Function


' ==================================================================
' 事前確認用（無変更）: どれだけ空打ちしているか数える
' ==================================================================
Public Sub 空行ガード_点検()
    Dim msg As String
    msg = "【空コード行の RSS 数式】（変更しません）" & vbCrLf & String(40, "-") & vbCrLf
    msg = msg & CountSheet("管理", 4, 461, "B")
    msg = msg & CountSheet("銘柄管理", 5, 505, "B")
    msg = msg & CountSheet("分析", 3, 206, "C")
    msg = msg & CountSheet("売分析", 3, 206, "C")
    MsgBox msg, vbInformation, "空行ガード 点検"
End Sub

Private Function CountSheet(ByVal shName As String, _
                            ByVal rowFirst As Long, ByVal rowLast As Long, _
                            ByVal codeCol As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then CountSheet = shName & ": シートなし" & vbCrLf: Exit Function

    Dim colLast As Long
    colLast = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1

    Dim rowsWithRss As Long, emptyRows As Long, cellsWasted As Long
    Dim r As Long, c As Long
    For r = rowFirst To rowLast
        Dim hasRss As Long: hasRss = 0
        For c = 1 To colLast
            If ws.Cells(r, c).HasFormula Then
                If InStr(1, ws.Cells(r, c).Formula, "Rss", vbTextCompare) > 0 Then hasRss = hasRss + 1
            End If
        Next c
        If hasRss > 0 Then
            rowsWithRss = rowsWithRss + 1
            If Trim$(CStr(ws.Range(codeCol & r).Value)) = "" Then
                emptyRows = emptyRows + 1
                cellsWasted = cellsWasted + hasRss
            End If
        End If
    Next r

    CountSheet = shName & ": RSS行=" & rowsWithRss & _
                 " / コード空=" & emptyRows & _
                 " / 無駄セル=" & cellsWasted & vbCrLf
End Function
