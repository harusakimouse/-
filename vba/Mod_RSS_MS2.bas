Attribute VB_Name = "Mod_RSS_MS2"
Option Explicit

'==================================================================
' MarketSpeed2 RSS 補助モジュール
'   ・Auto_Open  … ブックを開くと自動でRSSアドインを読み込み再計算
'   ・RSS今すぐ再接続 … 値が入らないときに押す
'   ・RSS更新間隔を遅くする … RTDサーバー応答なし対策
'   ・RSS関数の数を数える … 何個RSSを使っているか調べる
'   ・少し待つ … ブックを1本ずつ間隔をあけて開くために使う
'==================================================================

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

#If Win64 Then
    Private Const XLL_NAME As String = "MarketSpeed2_RSS_64bit.xll"
#Else
    Private Const XLL_NAME As String = "MarketSpeed2_RSS_32bit.xll"
#End If

'------------------------------------------------------------------
' ブックを開くと自動で走る（ThisWorkbookへの貼り付けは不要）
'------------------------------------------------------------------
Public Sub Auto_Open()
    On Error Resume Next
    RSSアドイン読込
    少し待つ 3
    Application.CalculateFullRebuild
    On Error GoTo 0
End Sub

'------------------------------------------------------------------
' ★値が入らないときに押す
'------------------------------------------------------------------
Public Sub RSS今すぐ再接続()
    On Error Resume Next
    RSSアドイン読込
    少し待つ 1
    Application.CalculateFullRebuild
    On Error GoTo 0
    MsgBox "RSSを読み直しました。" & vbCrLf & _
           "値が入らない場合は MarketSpeed2 のログインを確認してください。", _
           vbInformation, "RSS"
End Sub

'------------------------------------------------------------------
' ★RTDサーバー応答なし対策：更新間隔を5秒にする
'   （一度実行すればExcelを再起動しても効いたまま）
'------------------------------------------------------------------
Public Sub RSS更新間隔を遅くする()
    On Error Resume Next
    Application.RTD.ThrottleInterval = 5000
    If Err.Number <> 0 Then
        MsgBox "設定できませんでした。", vbExclamation, "RSS"
    Else
        MsgBox "RSSの更新間隔を5秒にしました。" & vbCrLf & _
               "元に戻すときは「RSS更新間隔を標準に戻す」を実行。", vbInformation, "RSS"
    End If
    On Error GoTo 0
End Sub

Public Sub RSS更新間隔を標準に戻す()
    On Error Resume Next
    Application.RTD.ThrottleInterval = 2000
    MsgBox "RSSの更新間隔を2秒（標準）に戻しました。", vbInformation, "RSS"
    On Error GoTo 0
End Sub

'------------------------------------------------------------------
' ★開いている全ブックのRSS関数の数を数える
'------------------------------------------------------------------
Public Sub RSS関数の数を数える()
    Dim wb As Workbook, ws As Worksheet, rng As Range, c As Range
    Dim n As Long, total As Long, msg As String

    Application.ScreenUpdating = False
    For Each wb In Application.Workbooks
        n = 0
        For Each ws In wb.Worksheets
            Set rng = Nothing
            On Error Resume Next
            Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0
            If Not rng Is Nothing Then
                For Each c In rng
                    If InStr(1, c.Formula, "Rss", vbTextCompare) > 0 Then n = n + 1
                Next c
            End If
        Next ws
        msg = msg & wb.Name & "  :  " & n & vbCrLf
        total = total + n
    Next wb
    Application.ScreenUpdating = True

    MsgBox msg & vbCrLf & "合計 " & total & " 個", vbInformation, "RSS関数の数"
End Sub

'------------------------------------------------------------------
' 指定秒だけ待つ（CPUを使わない＝RTDを止めない待ち方）
'------------------------------------------------------------------
Public Sub 少し待つ(ByVal sec As Long)
    Dim t As Double
    t = Timer
    Do While Timer - t < sec
        Sleep 200
        DoEvents
    Loop
End Sub

'------------------------------------------------------------------
' RSSアドイン(xll)を読み込む（既に有効なら何も起きない）
'------------------------------------------------------------------
Private Sub RSSアドイン読込()
    Dim cand As Variant, p As Variant

    cand = Array(Environ$("LOCALAPPDATA") & "\MarketSpeed2\Bin\rss\" & XLL_NAME, _
                 Environ$("USERPROFILE") & "\AppData\Local\MarketSpeed2\Bin\rss\" & XLL_NAME, _
                 "C:\MarketSpeed2\Bin\rss\" & XLL_NAME)

    For Each p In cand
        If Dir$(CStr(p)) <> "" Then
            On Error Resume Next
            Application.RegisterXLL CStr(p)
            On Error GoTo 0
            Exit Sub
        End If
    Next p
End Sub
