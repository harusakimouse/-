Attribute VB_Name = "MS2_Menu"
Option Explicit

'=========================================================
' 操作ボタンを専用シート「操作パネル_MS2」にまとめる
'  - DATA_MS2 などの表・数式は一切変更しません
'  - 既存の各処理マクロ(Module1)を呼び出すボタンを並べます
'  ※事前に、重複している Module2 は削除しておいてください
'    （MS2_Init_Book / MS2_Create_Buttons の二重定義対策）
'=========================================================
Sub MS2_Build_Menu()

    Dim ws As Worksheet
    Dim btn As Object
    Dim y As Single
    Dim items As Variant
    Dim i As Long

    On Error Resume Next
    Set ws = Sheets("操作パネル_MS2")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = Sheets.Add(Before:=Sheets(1))
        ws.Name = "操作パネル_MS2"
    End If

    ws.Cells.Clear
    ws.Buttons.Delete

    ws.Range("A1").Value = "◆ MS2 操作パネル"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ' Caption, 実行マクロ名 の順
    items = Array( _
        "銘柄反映_MS2", "MS2_Update_StockList_To_DATA", _
        "RSS式セット_MS2", "MS2_Set_RssMarket_Formulas", _
        "ATR更新_MS2", "MS2_Calc_ATR_5min_14", _
        "売買抽出_MS2", "MS2_Stock_Logic_Run", _
        "ランキング_MS2", "MS2_Update_Ranking", _
        "ログ_MS2", "MS2_Show_Log", _
        "全自動_MS2", "MS2_Auto_All")

    y = 36
    For i = LBound(items) To UBound(items) Step 2
        Set btn = ws.Buttons.Add(20, y, 170, 28)
        btn.Caption = items(i)
        btn.OnAction = items(i + 1)
        y = y + 36
    Next i

    ws.Range("A" & (Int(y / 15) + 4)).Value = _
        "※このシートのボタンから各処理を実行できます。"

    MsgBox "「操作パネル_MS2」シートにボタンを作成しました。", vbInformation, "MS2"

End Sub
