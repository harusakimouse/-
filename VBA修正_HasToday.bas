' =====================================================================
'  自動発火（自動取得）不具合の修正コード
'  対象ブック: 打ち出のこづち V0728  /  対象モジュール: Module1
' =====================================================================
'
'  【貼り付け方法】
'   1) Excel でブックを開く
'   2) Alt + F11 で VBE（VBA エディタ）を開く
'   3) 左のプロジェクトから「標準モジュール → Module1」をダブルクリック
'   4) 既存の HasToday 関数（下記【修正前】と同じもの）を丸ごと選択して
'      【修正A】の内容に置き換える
'   5) （任意）【修正B】を Module1 と ThisWorkbook にそれぞれ追記
'   6) 上書き保存（.xlsm のまま）
'
' =====================================================================
'  ■ なぜ直すのか（根本原因）
' ---------------------------------------------------------------------
'  日付列（A列）は「日付書式」付きセルなので、VBA で .Value を読むと
'  Date 型で返ります。ところが元コードは IsNumeric() で判定しており、
'  IsNumeric(Date型) は必ず False。そのため HasToday は「本日分あり」を
'  一度も True と判定できず、次の3つが同時に起きていました。
'    (1) 30秒ごとに何度でも再取得 → 大量の重複行（7/29は最大73回分）
'    (2) 実際は書けているのにログが毎回「書込0件 → 次回リトライ」
'    (3) 手動ボタンの重複確認ダイアログも出ない
' =====================================================================


' ---------------------------------------------------------------------
'  【修正前】（参考・これを削除）
' ---------------------------------------------------------------------
'  Public Function HasToday(ByVal nm As String) As Boolean
'      On Error Resume Next
'      Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(nm)
'      Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
'      If lr < 2 Then Exit Function
'      If IsNumeric(ws.Cells(lr, 1).Value) Then
'          HasToday = (CLng(ws.Cells(lr, 1).Value) = CLng(Date))
'      End If
'  End Function


' =====================================================================
'  【修正A】← 必須。Module1 の HasToday をこれに置き換える
' =====================================================================
Public Function HasToday(ByVal nm As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(nm)
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    If lr < 2 Then Exit Function

    Dim v As Variant: v = ws.Cells(lr, 1).Value
    If IsDate(v) Then
        ' 日付書式セル（.Value が Date 型で返る）→ シリアルの整数部で比較
        HasToday = (Int(CDbl(CDate(v))) = CLng(Date))
    ElseIf IsNumeric(v) Then
        ' 生のシリアル値が入っている場合の保険
        HasToday = (CLng(v) = CLng(Date))
    End If
End Function


' =====================================================================
'  【修正B】← 任意（推奨）。自動発火チェーンの自己復旧
'  ・OnTime の30秒チェーンは、6時間の間に一度でもタイミングを取り逃すと
'    そのまま停止し、開き直すまで復活しません（7/30に一度も自動取得
'    されなかった原因はこれと考えられます）。
'  ・下記を入れると、シートを切り替えるたびにチェーンの生存を確認し、
'    切れていれば静かに立て直します。修正Aで再発火は無害（冪等）に
'    なっているため、この再武装で重複は発生しません。
' =====================================================================

' ↓↓↓ Module1 の末尾に追記 ↓↓↓
Public Sub ReArmScheduler()
    On Error Resume Next
    If Time >= TimeValue("15:45:00") Then Exit Sub   ' 稼働時間外は何もしない
    If Not IsRunning Then Exit Sub                   ' そもそも停止中なら触らない

    Dim nx As Variant: nx = LogSheet.Range("F1").Value
    If IsDate(nx) Then
        ' 次回予約が「まだ生きている」なら何もしない（90秒の猶予）
        If CDate(nx) > Now - TimeValue("00:01:30") Then Exit Sub
    End If

    ' ここに来た＝チェーンが切れている → 立て直す
    ScheduleNext
    LogWrite "自己復旧", "ポーリング再武装"
End Sub

' ↓↓↓ ThisWorkbook モジュールに追記 ↓↓↓
'  Private Sub Workbook_SheetActivate(ByVal Sh As Object)
'      On Error Resume Next
'      ReArmScheduler
'  End Sub
