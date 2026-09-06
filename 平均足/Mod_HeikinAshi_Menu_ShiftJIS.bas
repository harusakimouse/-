Attribute VB_Name = "Mod_HeikinAshi_Menu"
Option Explicit
'==================================================================
' 平均足　メニュー（ボタン）作成　v1.0
'
'   実行 : Alt+F8 →「平均足_メニュー作成」
'   「メニュー」シートを作り、ボタンを並べます。
'   何度実行しても、古いボタンを消してから作り直します。
'==================================================================

Private Const SH As String = "メニュー"

'Excelがおかしくなった時の復旧（画面が固まった・反応しない時に押す）
Public Sub 平均足_画面を元に戻す()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Application.Cursor = xlDefault
    Application.Interactive = True
    On Error GoTo 0
    MsgBox "Excelの設定を元に戻しました。" & vbCrLf & _
           "（画面更新・イベント・自動計算をすべてONにしました）", vbInformation
End Sub

'ブックの構造保護を外す（保護されているとシートを追加できません）
Private Function HA_M_UnlockBook() As Boolean
    HA_M_UnlockBook = True
    If Not ThisWorkbook.ProtectStructure Then Exit Function
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=HA_SHEET_PW
    On Error GoTo 0
    If ThisWorkbook.ProtectStructure Then
        HA_M_UnlockBook = False
        MsgBox "このブックは「ブックの保護（構造）」がかかっているため、" & vbCrLf & _
               "シートを追加できません。" & vbCrLf & vbCrLf & _
               "校閲タブ → ブックの保護 → クリックして解除してください。" & vbCrLf & _
               "（パスワードを聞かれる場合は、Mod_HeikinAshi_Auto の先頭にある" & vbCrLf & _
               "　HA_SHEET_PW に入れてから、もう一度実行してください）", vbExclamation
    End If
End Function

Public Sub 平均足_メニュー作成()

    If Not HA_M_UnlockBook() Then Exit Sub

    On Error GoTo MenuErr
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = SH
    End If

    Application.ScreenUpdating = False

    '--- 古いボタンを消す ---
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ws.Cells.UnMerge
    ws.Cells.Clear

    '--- 見出し ---
    With ws.Range("B2:H2")
        .Merge
        .Value = "  平均足トレード　メニュー"
        .Font.Name = "Meiryo UI": .Font.Size = 20: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 70, 127)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 12
    ws.Rows(2).RowHeight = 44
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C:H").ColumnWidth = 12

    '--- ボタンを並べる ---
    Dim r As Long
    r = 4
    HA_M_Btn ws, r, "① 取込 ＋ 買い候補", "平均足_取込して抽出", _
             "OHLCVブックからデータを取り込み、そのまま買い候補を出します（引け後15:31以降）", RGB(0, 112, 192)
    HA_M_Btn ws, r, "② 買い候補だけ出す", "平均足_買い抽出", _
             "取込はせず、今あるデータで買い候補を出し直します", RGB(0, 112, 192)
    HA_M_Btn ws, r, "④ 自動実行を開始", "平均足_自動開始", _
             "毎営業日15:35に、取込→抽出を自動で行います（Excelを開いたままに）", RGB(0, 153, 68)
    HA_M_Btn ws, r, "⑤ 自動実行を止める", "平均足_自動停止", _
             "自動実行の予約を取り消します", RGB(128, 128, 128)
    HA_M_Btn ws, r, "⑥ 記録シート", "平均足_記録シート作成", _
             "トレード記録シートを作る／計算式を最新にする（記録は消えません）", RGB(197, 90, 17)
    HA_M_Btn ws, r, "⑦ 売買ルールを見る", "平均足_ルール表示", _
             "買う条件・点数・出口・資金管理を1枚のシートに出します", RGB(197, 90, 17)
    HA_M_Btn ws, r, "⑧ 記録シートの書式を戻す", "平均足_記録シート_書式を作り直す", _
             "色やフォントを初期状態に戻します（記録は消えません）", RGB(128, 128, 128)
    HA_M_Btn ws, r, "⑨ データシート作り直し", "平均足_データシート作り直し", _
             "始値～出来高の5枚を保護なしで作り直す（パスワードを聞かれる時だけ使う）", RGB(128, 128, 128)
    HA_M_Btn ws, r, "⑩ 画面を元に戻す", "平均足_画面を元に戻す", _
             "固まった後などに押す。画面更新・自動計算をすべてONに戻します", RGB(192, 0, 0)
    HA_M_Btn ws, r, "⑪ データチェック", "平均足_データチェック", _
             "データが古くないか・日付の欠損がないかを調べます（RSSが切れた日の事故防止）", RGB(0, 112, 192)

    '--- 毎日の手順 ---
    r = r + 1
    ws.Cells(r, 2).Value = "【毎日の手順】"
    ws.Cells(r + 1, 2).Value = "1. 15:35に自動で候補が出ます（④を一度押しておけば、あとは自動）"
    ws.Cells(r + 2, 2).Value = "2. 翌朝、「平均足」シートの上位から寄り成りで買う。同時5銘柄まで"
    ws.Cells(r + 3, 2).Value = "3. 寄り値が「見送りライン」より高く始まった銘柄は買わない"
    ws.Cells(r + 4, 2).Value = "4. 買ったらすぐ　逆指値－8%　と　売り指値+8%　を出す"
    ws.Cells(r + 5, 2).Value = "5. どちらも当たらなければ、5営業日後の引けで成行手仕舞い"
    ws.Cells(r + 6, 2).Value = "6. 記録シートに　買日・コード・銘柄名・買値・株数　を入れる"
    ws.Cells(r + 7, 2).Value = "※「今の指示」が『休む』になっていたら、その間は新規で買わないこと"
    ws.Cells(r + 8, 2).Value = "※このシステムは【買いのみ】です。空売りは検証で負けたため、機能ごと削除しました"
    With ws.Range(ws.Cells(r, 2), ws.Cells(r + 8, 2))
        .Font.Name = "Meiryo UI": .Font.Size = 12
    End With
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r + 7, 2).Font.Color = RGB(192, 0, 0)
    ws.Cells(r + 7, 2).Font.Bold = True
    ws.Cells(r + 8, 2).Font.Color = RGB(192, 0, 0)
    ws.Cells(r + 8, 2).Font.Bold = True

    On Error Resume Next
    ws.Activate
    ws.Range("A1").Select
    On Error GoTo 0

    Application.ScreenUpdating = True
    MsgBox "「" & SH & "」シートにボタンを作りました。" & vbCrLf & _
           "ボタンは⑩まであります。", vbInformation
    Exit Sub

MenuErr:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "メニュー作成でエラーが出ました。" & vbCrLf & _
           "番号 " & Err.Number & "：" & Err.Description, vbExclamation
End Sub

'ボタンを1つ作る（rは自動で進みます）
Private Sub HA_M_Btn(ByVal ws As Worksheet, ByRef r As Long, _
                     ByVal cap As String, ByVal macroName As String, _
                     ByVal note As String, ByVal col As Long)

    ws.Rows(r).RowHeight = 42

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                 ws.Cells(r, 2).Left + 3, ws.Cells(r, 2).Top + 3, _
                                 ws.Columns("B").Width - 6, 36)
    With shp
        .Name = "btn_" & macroName
        .Fill.ForeColor.RGB = col
        .Fill.Solid
        .Line.ForeColor.RGB = col
        .Shadow.Visible = msoFalse
        .OnAction = macroName
        With .TextFrame2.TextRange
            .Text = cap
            .Font.Name = "Meiryo UI"
            .Font.Size = 14
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

    With ws.Cells(r, 3)
        .Value = note
        .Font.Name = "Meiryo UI"
        .Font.Size = 11
        .VerticalAlignment = xlCenter
    End With

    r = r + 1
End Sub
