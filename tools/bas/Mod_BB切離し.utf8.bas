Attribute VB_Name = "Mod_BB切離し"
' ==================================================================
'  Mod_BB切離し ― V816 と BB816 をつなぐモジュール
'
'  V816   本体ブック（ボリンジャーバンド判定を降ろした版）
'  BB816  ボリンジャーバンド判定の独立ブック（マクロなし・数式だけ）
'
'  BB816 は自分の中に 終値/高値/安値/出来高 を持っているので、単独で開いて
'  計算できます。ただし価格は自動では更新されないので、日次更新のあとに
'  「BB816へ株価を送る」を1回実行してください。
'
'  入れ方
'    Alt+F11 → ファイル → ファイルのインポート → このファイルを選ぶ
'    （または新しい標準モジュールを作って、この中身を貼り付ける）
'
'  使うマクロ（Alt+F8 で実行できます）
'    BB816へ株価を送る    価格4シートを BB816.xlsx に流し込み、候補を表示する
'    BB816を開く          BB816.xlsx を開いて候補シートを見る
'    BB_シートを完全削除   V816 に残っている空のBBシート6枚を消す（1回だけ）
' ==================================================================
Option Explicit

Private Const BB_FILE As String = "BB816.xlsx"
Private Const PRICE_SHEETS As String = "終値,高値,安値,出来高"
Private Const COPY_RANGE As String = "A1:IT510"     ' 価格シートの全域（250営業日分）
Private Const BB_SHEETS As String = "BB使い方,BB設定,BBスクリーニング,BB買い候補TOP,BB_帯幅,BB_スイング"


' ------------------------------------------------------------------
' 価格データを BB816 に送る（日次更新のあとに実行）
' ------------------------------------------------------------------
Public Sub BB816へ株価を送る()
    Call 価格送信(False)
End Sub

' 日次更新のマクロの最後に  Call BB816へ株価を送る_無言  と書いておくと、
' メッセージを出さずに自動で連携できます。
Public Sub BB816へ株価を送る_無言()
    Call 価格送信(True)
End Sub


Private Sub 価格送信(ByVal silent As Boolean)
    Dim path As String, wb As Workbook, nm As Variant
    Dim srcWs As Worksheet, dstWs As Worksheet
    Dim calcMode As XlCalculation, opened As Boolean
    Dim msg As String, hits As Variant

    path = BB816のパス(silent)
    If path = "" Then Exit Sub

    On Error GoTo 失敗
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wb = 既に開いているか(BB_FILE)
    If wb Is Nothing Then
        Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0)
        opened = True
    End If

    For Each nm In Split(PRICE_SHEETS, ",")
        Set srcWs = Nothing: Set dstWs = Nothing
        On Error Resume Next
        Set srcWs = ThisWorkbook.Worksheets(CStr(nm))
        Set dstWs = wb.Worksheets(CStr(nm))
        On Error GoTo 失敗
        If srcWs Is Nothing Then
            MsgBox "本体ブックに「" & nm & "」シートがありません。", vbExclamation, "BB816連携"
            GoTo 後始末
        End If
        If dstWs Is Nothing Then
            MsgBox BB_FILE & " に「" & nm & "」シートがありません。", vbExclamation, "BB816連携"
            GoTo 後始末
        End If
        ' 値だけを丸ごと転送する（RSSの数式は持ち込まない）
        dstWs.Range(COPY_RANGE).Value = srcWs.Range(COPY_RANGE).Value
    Next nm

    Application.Calculation = xlCalculationAutomatic
    wb.Application.CalculateFullRebuild
    wb.Save

    hits = ""
    On Error Resume Next
    hits = wb.Worksheets("BB816候補").Range("C3").Value
    On Error GoTo 失敗

    If Not silent Then
        wb.Activate
        On Error Resume Next
        wb.Worksheets("BB816候補").Activate
        On Error GoTo 失敗
        msg = "BB816 に株価を送りました。" & vbCrLf & vbCrLf & _
              "データ最終日 : " & 最終日表示() & vbCrLf & _
              "◎買いの件数 : " & hits & " 件" & vbCrLf & vbCrLf & _
              "上から最大3銘柄まで、翌営業日の寄付で買ってください。" & vbCrLf & _
              "買えたらすぐ、損切り価格に逆指値を置いてください。"
        MsgBox msg, vbInformation, "BB816 連携"
    End If

後始末:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

失敗:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Not silent Then
        MsgBox "BB816への転送に失敗しました。" & vbCrLf & _
               Err.Number & " : " & Err.Description, vbExclamation, "BB816 連携"
    End If
End Sub


' ------------------------------------------------------------------
' BB816 を開くだけ
' ------------------------------------------------------------------
Public Sub BB816を開く()
    Dim path As String, wb As Workbook
    path = BB816のパス(False)
    If path = "" Then Exit Sub
    Set wb = 既に開いているか(BB_FILE)
    If wb Is Nothing Then Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0)
    wb.Activate
    On Error Resume Next
    wb.Worksheets("BB816候補").Activate
End Sub


' ------------------------------------------------------------------
' V816 に残っている空のBBシートを完全に削除する
'   ZIP編集ではシートを消さず「空にして非表示」にしてあります。
'   VBAのDocumentモジュールごと安全に消せるのは Excel 自身だけなので、
'   完全に消したい場合だけ、このマクロを1回実行してください。
' ------------------------------------------------------------------
Public Sub BB_シートを完全削除()
    Dim nm As Variant, ws As Worksheet, n As Long, 残 As String

    If MsgBox("V816 に残っている空のBBシート6枚を削除します。" & vbCrLf & vbCrLf & _
              "・中身はすでに空なので、判定結果は変わりません" & vbCrLf & _
              "・削除後は必ず上書き保存してください" & vbCrLf & vbCrLf & _
              "実行しますか？", vbYesNo + vbQuestion, "BBシートの削除") <> vbYes Then Exit Sub

    Application.DisplayAlerts = False
    For Each nm In Split(BB_SHEETS, ",")
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        On Error GoTo 0
        If Not ws Is Nothing Then
            If Application.CountA(ws.UsedRange) <= 12 Then
                ws.Visible = xlSheetVisible
                ws.Delete
                n = n + 1
            Else
                残 = 残 & vbCrLf & "  " & nm & "（中身が残っているので消しませんでした）"
            End If
        End If
    Next nm
    Application.DisplayAlerts = True

    MsgBox n & " 枚のシートを削除しました。" & 残 & vbCrLf & vbCrLf & _
           "上書き保存してください。", vbInformation, "BBシートの削除"
End Sub


' ------------------------------------------------------------------
' 補助
' ------------------------------------------------------------------
Private Function BB816のパス(ByVal silent As Boolean) As String
    Dim p As String, f As Variant
    p = ThisWorkbook.path & Application.PathSeparator & BB_FILE
    If Dir(p) <> "" Then
        BB816のパス = p
        Exit Function
    End If
    If silent Then Exit Function

    MsgBox BB_FILE & " が本体ブックと同じフォルダにありません。" & vbCrLf & _
           "場所を選んでください。", vbInformation, "BB816 連携"
    f = Application.GetOpenFilename("Excelブック,*.xlsx;*.xlsm", , "BB816.xlsx を選ぶ")
    If VarType(f) = vbBoolean Then Exit Function
    BB816のパス = CStr(f)
End Function


Private Function 既に開いているか(ByVal fileName As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, fileName, vbTextCompare) = 0 Then
            Set 既に開いているか = wb
            Exit Function
        End If
    Next wb
End Function


Private Function 最終日表示() As String
    On Error Resume Next
    最終日表示 = ThisWorkbook.Worksheets("終値").Range("E3").Text
    If 最終日表示 = "" Then 最終日表示 = "（不明）"
End Function
