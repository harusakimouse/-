Attribute VB_Name = "Mod_Patch_改善v1"
'=============================================================
' 打ち出のこづち 改善パッチ v1
' 適用方法:
'   1. Alt+F11 でVBAエディタを開く
'   2. このファイルをインポート (ファイル → ファイルのインポート)
'   3. イミディエイトウィンドウ (Ctrl+G) で
'      「Call 改善パッチ適用」と入力してEnter
'   4. 完了メッセージを確認後、このモジュールを削除してOK
'=============================================================

Public Sub 改善パッチ適用()
    Dim vbProj As Object
    Dim modBuy  As Object
    Dim modSell As Object

    On Error GoTo ErrHandler

    ' VBIDEアクセス（Excelオプション→セキュリティセンター→
    '   マクロの設定→「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」が必要）
    Set vbProj = ThisWorkbook.VBProject

    ' モジュールを取得
    Dim comp As Object
    For Each comp In vbProj.VBComponents
        If comp.name = "Mod_買抽出v13" Then Set modBuy = comp.CodeModule
        If comp.name = "Mod_SellExtrac" Then Set modSell = comp.CodeModule
    Next comp

    If modBuy Is Nothing Then MsgBox "Mod_買抽出v13 が見つかりません", vbCritical: Exit Sub
    If modSell Is Nothing Then MsgBox "Mod_SellExtrac が見つかりません", vbCritical: Exit Sub

    Dim log As String: log = ""
    Dim cnt As Long: cnt = 0

    '===========================================================
    ' 修正1: 買い RSI≥70 ハードフィルター追加（Mod_買抽出v13）
    '===========================================================
    Dim old1 As String, new1 As String
    old1 = "        If doMark And canExtract And BUY_REQUIRE_OSIME Then" & vbCrLf & _
           "            If osime <> ""25日線押し目●"" And osime <> ""5日線押し目○"" Then" & vbCrLf & _
           "                canExtract = False" & vbCrLf & _
           "            End If" & vbCrLf & _
           "        End If"
    new1 = "        If doMark And canExtract And BUY_REQUIRE_OSIME Then" & vbCrLf & _
           "            If osime <> ""25日線押し目●"" And osime <> ""5日線押し目○"" Then" & vbCrLf & _
           "                canExtract = False" & vbCrLf & _
           "            End If" & vbCrLf & _
           "        End If" & vbCrLf & _
           "        ' 修正1: RSI過熱（≥70）はハード除外（踏み上げ・調整リスク）" & vbCrLf & _
           "        If rsiVal >= 70 Then canExtract = False"

    If ReplaceInModule(modBuy, old1, new1) Then
        log = log & "✅ 修正1: 買いRSI≥70ハードフィルター 追加" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "⚠️ 修正1: 対象コードが見つかりませんでした（適用済みの可能性）" & vbCrLf
    End If

    '===========================================================
    ' 修正2: 売り RSI<15 ハードフィルター追加（Mod_SellExtrac）
    '        S_抽出実行 と S_抽出実行_RSS の両方に適用
    '===========================================================
    Dim old2 As String, new2 As String
    old2 = "        If doMark And canExtract Then" & vbCrLf & _
           "            If momCount < 2 Then canExtract = False" & vbCrLf & _
           "            If S_REQUIRE_DOWNTREND And Not ok_PD Then canExtract = False" & vbCrLf & _
           "        End If"
    new2 = "        ' 修正2: RSI極値（<15）は踏み上げリスクで除外" & vbCrLf & _
           "        If canExtract And rsiVal < 15 Then canExtract = False" & vbCrLf & _
           "        If doMark And canExtract Then" & vbCrLf & _
           "            If momCount < 2 Then canExtract = False" & vbCrLf & _
           "            If S_REQUIRE_DOWNTREND And Not ok_PD Then canExtract = False" & vbCrLf & _
           "        End If"

    Dim replaced2 As Long: replaced2 = 0
    Do
        If ReplaceInModule(modSell, old2, new2) Then
            replaced2 = replaced2 + 1
        Else
            Exit Do
        End If
    Loop Until replaced2 >= 2  ' S_抽出実行 と S_抽出実行_RSS の2箇所

    If replaced2 > 0 Then
        log = log & "✅ 修正2: 売りRSI<15ハードフィルター 追加（" & replaced2 & "箇所）" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "⚠️ 修正2: 対象コードが見つかりませんでした" & vbCrLf
    End If

    '===========================================================
    ' 修正3a: 買い抽出シート 利確ライン表示 +5%→+8%（writeRow）
    '===========================================================
    Dim old3a As String, new3a As String
    old3a = "        .Cells(r, 15).Value = Format(closePrice * 1.05, ""#,##0"")"
    new3a = "        .Cells(r, 15).Value = Format(closePrice * 1.08, ""#,##0"")  ' 修正3: +5%→+8%"

    If ReplaceInModule(modBuy, old3a, new3a) Then
        log = log & "✅ 修正3a: 買い利確ライン表示 +5%→+8% 変更" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "⚠️ 修正3a: writeRow 利確ライン対象未発見" & vbCrLf
    End If

    '===========================================================
    ' 修正3b: 買い抽出シート 損切りライン表示 -3%→-4%（writeRow）
    '         自動注文決済のSL=4%と統一
    '===========================================================
    Dim old3b As String, new3b As String
    old3b = "        .Cells(r, 16).Value = Format(closePrice * 0.97, ""#,##0"")"
    new3b = "        .Cells(r, 16).Value = Format(closePrice * 0.96, ""#,##0"")  ' 修正3: -3%→-4%（SLと統一）"

    If ReplaceInModule(modBuy, old3b, new3b) Then
        log = log & "✅ 修正3b: 買い損切りライン表示 -3%→-4% 変更" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "⚠️ 修正3b: writeRow 損切りライン対象未発見" & vbCrLf
    End If

    '===========================================================
    ' 修正4: 緩和モードでも momCount≥2 を要求（Mod_買抽出v13）
    '        8銘柄→27銘柄の急増を抑制
    '===========================================================
    Dim old4 As String, new4 As String
    old4 = "        If doMark And canExtract Then" & vbCrLf & _
           "            If momCount < BUY_REQUIRE_MOMENTUM_COUNT Then canExtract = False" & vbCrLf & _
           "        End If"
    new4 = "        ' 修正4: 緩和モード(doMark=False)でも最低momCount≥2を要求" & vbCrLf & _
           "        If canExtract And momCount < 2 Then canExtract = False" & vbCrLf & _
           "        If doMark And canExtract Then" & vbCrLf & _
           "            If momCount < BUY_REQUIRE_MOMENTUM_COUNT Then canExtract = False" & vbCrLf & _
           "        End If"

    If ReplaceInModule(modBuy, old4, new4) Then
        log = log & "✅ 修正4: 緩和モード momCount≥2 追加（過剰抽出を抑制）" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "⚠️ 修正4: 対象コードが見つかりませんでした" & vbCrLf
    End If

    '===========================================================
    ' 自動注文決済シート 利確%セルは openpyxl で変更済み (6→8)
    '===========================================================

    MsgBox "【改善パッチ適用結果】" & vbCrLf & vbCrLf & log & vbCrLf & _
           "適用件数: " & cnt & " / 5" & vbCrLf & vbCrLf & _
           "※自動注文決済シートの利確% は 8% に変更済みです" & vbCrLf & _
           "  （このパッチ適用後、このモジュールを削除してください）", _
           vbInformation, "改善パッチ v1 完了"
    Exit Sub

ErrHandler:
    MsgBox "エラー " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "VBIDEアクセスが無効な場合: Excel → オプション → セキュリティセンター" & vbCrLf & _
           "→ マクロの設定 → 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」をONに", _
           vbCritical, "エラー"
End Sub

'-----------------------------------------------------------
' モジュール内テキスト置換ヘルパー
'-----------------------------------------------------------
Private Function ReplaceInModule(ByVal cm As Object, _
                                  ByVal oldText As String, _
                                  ByVal newText As String) As Boolean
    Dim totalLines As Long: totalLines = cm.CountOfLines
    Dim fullCode As String: fullCode = cm.Lines(1, totalLines)

    Dim pos As Long: pos = InStr(fullCode, oldText)
    If pos = 0 Then
        ReplaceInModule = False
        Exit Function
    End If

    Dim newCode As String
    newCode = Left(fullCode, pos - 1) & newText & Mid(fullCode, pos + Len(oldText))

    cm.DeleteLines 1, totalLines
    cm.InsertLines 1, newCode
    ReplaceInModule = True
End Function
