Attribute VB_Name = "Mod_RSS_MS2"
Option Explicit

'==================================================================
' MarketSpeed2 RSS 補助モジュール
'   ・RSSアドイン(xll)を読み込む
'   ・ブックを1本ずつ間隔をあけて開く（少し待つ）
'   ・つなぎ直しは「RSS今すぐ再接続」を押す
'  ※タイマー監視は入れない（RTDサーバーが不安定になるため）
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
' ブックを開いたときに呼ばれる（ThisWorkbookから）
'------------------------------------------------------------------
Public Sub 自動RSS準備()
    On Error Resume Next
    RSSアドイン読込
    少し待つ 3
    Application.CalculateFullRebuild
    On Error GoTo 0
End Sub

'------------------------------------------------------------------
' ★手動用：RSSをつなぎ直す（値が入らないときに押す）
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
' ★ボタン用：RSSを読み込んでから、起動リストのブックを開く
'------------------------------------------------------------------
Public Sub 一発起動()
    On Error Resume Next
    RSSアドイン読込
    Application.Run "Mod_Launcher.起動"
    少し待つ 2
    Application.CalculateFullRebuild
    On Error GoTo 0
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

'------------------------------------------------------------------
' 指定秒だけ待つ（CPUを使わない待ち方＝RTDを止めない）
'   ブックを1本ずつ間隔をあけて開くために使う
'------------------------------------------------------------------
Public Sub 少し待つ(ByVal sec As Long)
    Dim t As Double
    t = Timer
    Do While Timer - t < sec
        Sleep 200
        DoEvents
    Loop
End Sub
