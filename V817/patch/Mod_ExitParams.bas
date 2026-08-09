Attribute VB_Name = "Mod_ExitParams"
'==================================================================
'  Mod_ExitParams  （V817 で追加）
'
'  損切幅・利確幅を V811設定シート 1か所から取得するための共通関数。
'  抽出系モジュールに直接書かれていた 1.08 / 0.96 / 0.92 / 1.04 を
'  これらの関数呼び出しに置き換えて使う。
'
'  【導入手順】
'   1) VBE (Alt+F11) → ファイル → ファイルのインポート → このファイルを選択
'   2) Mod_買抽出v13 の 964〜965 行目を下記のように差し替え
'         .Cells(r, 15).Value = Format(ExitTP(closePrice, True), "#,##0")
'         .Cells(r, 16).Value = Format(ExitSL(closePrice, True), "#,##0")
'   3) Mod_SellExtrac の 1022 / 1024 行目を下記のように差し替え
'         .Cells(r, 15).Value = ExitTP(closePrice, False)   'O:利確ライン
'         .Cells(r, 16).Value = ExitSL(closePrice, False)   'P:損切ライン
'   4) Mod_買抽出_順張りVWAP / Mod_売抽出_下降VWAP にも同じ定数があるため
'      使用している場合は同様に差し替える
'==================================================================
Option Explicit

Private Const SET_SHEET As String = "V811設定"

'--- 設定値の取得（シートが無い場合は安全側の既定値を返す） ---
Private Function P(ByVal addr As String, ByVal defaultVal As Double) As Double
    Dim ws As Worksheet, v As Variant
    On Error GoTo Fallback
    Set ws = ThisWorkbook.Worksheets(SET_SHEET)
    v = ws.Range(addr).Value
    If IsNumeric(v) Then
        If CDbl(v) <> 0 Then
            P = CDbl(v)
            Exit Function
        End If
    End If
Fallback:
    P = defaultVal
End Function

'==================================================================
' 銘柄コードの 14日平均レンジ率(%) を返す。取得できないときは 0。
'   V811指標!C列 と同じ定義： (Σ高値 − Σ安値) ÷ Σ終値 × 100
'==================================================================
Public Function AtrPct(ByVal code As String) As Double
    Dim wsMei As Worksheet, wsH As Worksheet, wsL As Worksheet, wsC As Worksheet
    Dim r As Long, i As Long, n As Long
    Dim sh As Double, sl As Double, sc As Double
    Dim v As Variant

    On Error GoTo Fail
    Set wsMei = ThisWorkbook.Worksheets("銘柄管理")
    Set wsH = ThisWorkbook.Worksheets("高値")
    Set wsL = ThisWorkbook.Worksheets("安値")
    Set wsC = ThisWorkbook.Worksheets("終値")

    r = 0
    For i = 6 To 305
        If Trim(CStr(wsMei.Cells(i, 2).Value)) = Trim(code) Then
            r = i
            Exit For
        End If
    Next i
    If r = 0 Then GoTo Fail

    n = CLng(P("$H$10", 14))
    If n < 2 Then n = 14

    For i = 5 To 4 + n                      ' E列(=5)から n 列分
        v = wsH.Cells(r, i).Value: If Not IsNumeric(v) Then GoTo Fail Else sh = sh + CDbl(v)
        v = wsL.Cells(r, i).Value: If Not IsNumeric(v) Then GoTo Fail Else sl = sl + CDbl(v)
        v = wsC.Cells(r, i).Value: If Not IsNumeric(v) Then GoTo Fail Else sc = sc + CDbl(v)
    Next i

    If sc <= 0 Then GoTo Fail
    AtrPct = (sh - sl) / sc * 100
    Exit Function
Fail:
    AtrPct = 0
End Function

'==================================================================
' 損切幅(%) を返す。
'   V811設定!H3 = 1 … ATR × H4 を H5〜H6 でクランプ
'   V811設定!H3 = 0 … B3 × 100（％固定）
'==================================================================
Public Function ExitStopPct(Optional ByVal code As String = "") As Double
    Dim mode As Double, w As Double, a As Double
    mode = P("$H$3", 1)
    If mode = 1 And Len(code) > 0 Then
        a = AtrPct(code)
        If a > 0 Then
            w = a * P("$H$4", 2)
            If w < P("$H$5", 3) Then w = P("$H$5", 3)
            If w > P("$H$6", 10) Then w = P("$H$6", 10)
            ExitStopPct = w
            Exit Function
        End If
    End If
    ExitStopPct = P("$B$3", 0.06) * 100
End Function

Public Function ExitTargetPct(Optional ByVal code As String = "") As Double
    ExitTargetPct = ExitStopPct(code) * P("$B$4", 2)
End Function

'==================================================================
' 価格ベースのライン。isBuy=True で買い、False で売り。
'   code を渡すと ATR連動、省略すると％固定にフォールバックする。
'==================================================================
Public Function ExitTP(ByVal price As Double, ByVal isBuy As Boolean, _
                       Optional ByVal code As String = "") As Double
    Dim p2 As Double: p2 = ExitTargetPct(code) / 100
    If isBuy Then ExitTP = Round(price * (1 + p2), 0) Else ExitTP = Round(price * (1 - p2), 0)
End Function

Public Function ExitSL(ByVal price As Double, ByVal isBuy As Boolean, _
                       Optional ByVal code As String = "") As Double
    Dim p2 As Double: p2 = ExitStopPct(code) / 100
    If isBuy Then ExitSL = Round(price * (1 - p2), 0) Else ExitSL = Round(price * (1 + p2), 0)
End Function

'==================================================================
' 推奨株数（1回の許容損失額から逆算・窓抜け安全率込み・単元100株）
'   1単元でも予算を超える場合は 0（＝見送り）を返す
'==================================================================
Public Function SuggestQty(ByVal price As Double, Optional ByVal code As String = "") As Long
    Dim w As Double, risk As Double
    w = ExitStopPct(code) / 100
    If price <= 0 Or w <= 0 Then Exit Function
    risk = price * w * P("$H$7", 1.07)
    If risk <= 0 Then Exit Function
    SuggestQty = Int(P("$B$5", 20000) / risk / 100) * 100
End Function
