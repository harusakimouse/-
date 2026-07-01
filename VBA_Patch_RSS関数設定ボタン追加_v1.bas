Attribute VB_Name = "Mod_Patch_RSSButton_v1"
'=============================================================
' RSS関数設定ボタン追加パッチ v1（0701v8.xlsm 用）
'
' 症状: 「15:00 スナップ」「15:20 スナップ」ボタンを押してもデータが
'        取得できない。
'
' 原因: RSSデータシートのC〜L列（現在値・前日比率・始値・安値・高値等）
'        は本来 SetRSSFormulas プロシージャが設定する
'        =RssMarket(...) / =RssIndexMarket(...) 式で自動更新される
'        設計になっているが、RSSデータシートには
'        「← RSS関数設定ボタンを押すとB〜N列に自動入力」という案内が
'        あるだけで、実際にSetRSSFormulasを呼び出す「RSS関数設定」
'        ボタンが CreateButtons プロシージャで作られていない。
'        そのため各セルは固定値0のまま更新されず、
'        15:00/15:20スナップボタンはRSSデータシートの0をそのまま
'        コピーするだけになっていた。
'
' 対処: CreateButtons に、RSSデータシート上へ「RSS関数設定」ボタン
'        （OnAction = SetRSSFormulas）を作成する処理を追加する。
'
' 適用方法:
'   1. Alt+F11 でVBAエディタを開く
'   2. このファイルをインポート（ファイル → ファイルのインポート）
'   3. イミディエイトウィンドウ (Ctrl+G) で
'      「Call RSSボタン追加パッチ適用」と入力してEnter
'      ※「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」
'        （Excelオプション→セキュリティセンター→マクロの設定）が
'        オンになっている必要があります。
'   4. 完了メッセージを確認後、このモジュール(Mod_Patch_RSSButton_v1)を
'      削除してOK
'   5. マクロ「CreateButtons」を実行
'      → RSSデータシートに「RSS関数設定」ボタンが追加されます
'   6. 追加された「RSS関数設定」ボタンを1回クリック
'      （RssMarket式がセットされ、以後は場中ずっと自動更新されます）
'   7. これで 15:00 / 15:20 のスナップボタンで実データが
'      取得できるようになります
'=============================================================

Public Sub RSSボタン追加パッチ適用()
    Dim vbProj As Object
    Dim modMain As Object

    On Error GoTo ErrHandler
    Set vbProj = ThisWorkbook.VBProject

    Dim comp As Object
    For Each comp In vbProj.VBComponents
        If comp.name = "Module1" Then Set modMain = comp.CodeModule
    Next comp

    If modMain Is Nothing Then MsgBox "Module1 が見つかりません", vbCritical: Exit Sub

    Dim log As String: log = ""

    Dim oldTxt As String, newTxt As String
    oldTxt = "Sub CreateButtons()" & vbCrLf & _
             "    Dim ws3 As Worksheet, ws4 As Worksheet" & vbCrLf & _
             "    Dim btn As Object" & vbCrLf & _
             "    Dim shp As Shape" & vbCrLf & _
             "    Set ws3 = ThisWorkbook.Sheets(3)" & vbCrLf & _
             "    Set ws4 = ThisWorkbook.Sheets(4)" & vbCrLf & _
             "    For Each shp In ws3.Shapes" & vbCrLf & _
             "        shp.Delete" & vbCrLf & _
             "    Next shp" & vbCrLf & _
             "    For Each shp In ws4.Shapes" & vbCrLf & _
             "        shp.Delete" & vbCrLf & _
             "    Next shp"
    newTxt = "Sub CreateButtons()" & vbCrLf & _
             "    Dim ws2 As Worksheet, ws3 As Worksheet, ws4 As Worksheet" & vbCrLf & _
             "    Dim btn As Object" & vbCrLf & _
             "    Dim shp As Shape" & vbCrLf & _
             "    Set ws2 = ThisWorkbook.Sheets(2)" & vbCrLf & _
             "    Set ws3 = ThisWorkbook.Sheets(3)" & vbCrLf & _
             "    Set ws4 = ThisWorkbook.Sheets(4)" & vbCrLf & _
             "    For Each shp In ws2.Shapes" & vbCrLf & _
             "        shp.Delete" & vbCrLf & _
             "    Next shp" & vbCrLf & _
             "    For Each shp In ws3.Shapes" & vbCrLf & _
             "        shp.Delete" & vbCrLf & _
             "    Next shp" & vbCrLf & _
             "    For Each shp In ws4.Shapes" & vbCrLf & _
             "        shp.Delete" & vbCrLf & _
             "    Next shp" & vbCrLf & _
             "    Set btn = ws2.Shapes.AddFormControl(xlButtonControl, 760, 4, 180, 30)" & vbCrLf & _
             "    btn.Name = ""btnSetRSS""" & vbCrLf & _
             "    btn.TextFrame.Characters.Text = ""RSS関数設定""" & vbCrLf & _
             "    btn.TextFrame.Characters.Font.Size = 16" & vbCrLf & _
             "    btn.OnAction = ""SetRSSFormulas"""

    If ReplaceInModule(modMain, oldTxt, newTxt) Then
        log = log & "OK: CreateButtons に「RSS関数設定」ボタン作成処理を追加しました" & vbCrLf
    Else
        log = log & "SKIP: 対象コードが見つかりませんでした（適用済み、または既にコードが変更されている可能性があります）" & vbCrLf
    End If

    MsgBox "【RSS関数設定ボタン追加パッチ 適用結果】" & vbCrLf & vbCrLf & log & vbCrLf & _
           "次の手順:" & vbCrLf & _
           "1. このモジュール(Mod_Patch_RSSButton_v1)を削除" & vbCrLf & _
           "2. マクロ「CreateButtons」を実行" & vbCrLf & _
           "   （RSSデータシートに「RSS関数設定」ボタンが追加されます）" & vbCrLf & _
           "3. 追加された「RSS関数設定」ボタンを1回クリック" & vbCrLf & _
           "   （RssMarket式がセットされ、場中は自動更新されます）" & vbCrLf & _
           "4. その後、15:00 / 15:20 のスナップボタンで実データが取得できます", _
           vbInformation, "RSS関数設定ボタン追加パッチ v1 完了"
    Exit Sub

ErrHandler:
    MsgBox "エラー " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "VBIDEアクセスが無効な場合: Excel → オプション → セキュリティセンター" & vbCrLf & _
           "→ マクロの設定 → 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」をONに", _
           vbCritical, "エラー"
End Sub

Private Function ReplaceInModule(ByVal cm As Object, _
                                  ByVal oldText As String, _
                                  ByVal newText As String) As Boolean
    Dim totalLines As Long: totalLines = cm.CountOfLines
    Dim fullCode As String: fullCode = cm.Lines(1, totalLines)
    Dim pos As Long: pos = InStr(fullCode, oldText)
    If pos = 0 Then ReplaceInModule = False: Exit Function
    Dim newCode As String
    newCode = Left(fullCode, pos - 1) & newText & Mid(fullCode, pos + Len(oldText))
    cm.DeleteLines 1, totalLines
    cm.InsertLines 1, newCode
    ReplaceInModule = True
End Function
