Attribute VB_Name = "Mod_Patch_VWAP_v1"
'=============================================================
' VWAP(出来高加重平均)取得パッチ v1（0701v8.xlsm 用）
'
' 目的: スナップショットシートのM列(VWAP)は見出しと表示形式だけが
'        用意されており、実際のVWAP値を取得・転記する処理が
'        入っていなかった（SortByScoreの判定材料としては既に
'        参照されているが、常に空欄のため機能していなかった）。
'
' 対処:
'   1. RSSデータシートのN列（従来未使用）に
'      =RssMarket(コード,"出来高加重平均") 式を追加
'   2. SetRSSFormulasのクリア範囲を A3:M302 → A3:N302 に拡張
'   3. Snapshot1500で、RSSデータシートN列の値を
'      スナップショットシートM列へコピーする処理を追加
'
' 適用方法:
'   1. Alt+F11 でVBAエディタを開く
'   2. このファイルをインポート（ファイル → ファイルのインポート）
'   3. イミディエイトウィンドウ (Ctrl+G) で
'      「Call VWAP追加パッチ適用」と入力してEnter
'      ※「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」
'        がオンになっている必要があります。
'   4. 完了メッセージを確認後、このモジュール(Mod_Patch_VWAP_v1)を
'      削除してOK
'   5. 「RSS関数設定」ボタンをもう一度クリック
'      （RSSデータシートN列にVWAP式がセットされます）
'   6. 15:00 / 15:20 のスナップボタンを押すと、
'      スナップショットシートM列に実際のVWAP値が入るようになります
'=============================================================

Public Sub VWAP追加パッチ適用()
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

    '-----------------------------------------------------------
    ' 修正1: SetRSSFormulas のクリア範囲拡張 + N2ヘッダー追加
    '-----------------------------------------------------------
    Dim old1 As String, new1 As String
    old1 = "    ws_rss.Range(""A3:M302"").ClearContents"
    new1 = "    ws_rss.Range(""A3:N302"").ClearContents" & vbCrLf & _
             "    ws_rss.Cells(2, 14).Value = ""VWAP"""

    If ReplaceInModule(modMain, old1, new1) Then
        log = log & "OK: SetRSSFormulas のクリア範囲をA3:N302に拡張し、N2見出しを追加しました" & vbCrLf
    Else
        log = log & "SKIP: 修正1 対象コードが見つかりませんでした（適用済みの可能性）" & vbCrLf
    End If

    '-----------------------------------------------------------
    ' 修正2: 各行にVWAP取得式(N列)を追加
    '-----------------------------------------------------------
    Dim old2b As String, new2b As String
    old2b = "        ws_rss.Cells(r, 13).NumberFormat = ""0.00%""" & vbCrLf & _
             "        r = r + 1"
    new2b = "        ws_rss.Cells(r, 13).NumberFormat = ""0.00%""" & vbCrLf & _
             "        ws_rss.Cells(r, 14).Formula = ""=RssMarket("""""" & code & """""","""""" & ChrW(&H51FA) & ChrW(&H6765) & ChrW(&H9AD8) & ChrW(&H52A0) & ChrW(&H91CD) & ChrW(&H5E73) & ChrW(&H5747) & """""")""" & vbCrLf & _
             "        r = r + 1"

    If ReplaceInModule(modMain, old2b, new2b) Then
        log = log & "OK: RSSデータシート N列にVWAP取得式(出来高加重平均)を追加しました" & vbCrLf
    Else
        log = log & "SKIP: 修正2 対象コードが見つかりませんでした（適用済みの可能性）" & vbCrLf
    End If

    '-----------------------------------------------------------
    ' 修正3: Snapshot1500 で N列 → M列へVWAPをコピー
    '-----------------------------------------------------------
    Dim old3 As String, new3 As String
    old3 = "    ws_snap.Cells(2, 13).Value = ""VWAP""" & vbCrLf & _
             "    ws_snap.Range(""M3:M"" & (cnt + 2)).NumberFormat = ""#,##0"""
    new3 = "    ws_snap.Cells(2, 13).Value = ""VWAP""" & vbCrLf & _
             "    ws_snap.Range(""M3:M"" & (cnt + 2)).Value = ws_rss.Range(""N3:N"" & (cnt + 2)).Value" & vbCrLf & _
             "    ws_snap.Range(""M3:M"" & (cnt + 2)).NumberFormat = ""#,##0"""

    If ReplaceInModule(modMain, old3, new3) Then
        log = log & "OK: Snapshot1500 にVWAP(N列→M列)のコピー処理を追加しました" & vbCrLf
    Else
        log = log & "SKIP: 修正3 対象コードが見つかりませんでした（適用済みの可能性）" & vbCrLf
    End If

    MsgBox "【VWAP追加パッチ 適用結果】" & vbCrLf & vbCrLf & log & vbCrLf & _
           "次の手順:" & vbCrLf & _
           "1. このモジュール(Mod_Patch_VWAP_v1)を削除" & vbCrLf & _
           "2. 「RSS関数設定」ボタンをもう一度クリック" & vbCrLf & _
           "   （RSSデータシートN列にVWAP式がセットされます）" & vbCrLf & _
           "3. 15:00 / 15:20 のスナップボタンを押すとスナップショットM列に" & vbCrLf & _
           "   実際のVWAP値が入ります", _
           vbInformation, "VWAP追加パッチ v1 完了"
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
