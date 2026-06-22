Attribute VB_Name = "Mod_Patch_金のこづちV2"
'=============================================================
' 金のこづち 0622 総合改善パッチ v2
'
' ■ 改善の根拠（4者の意見を統合）
'   ・私の分析：売り構造はミラー設計で精度不十分、RSI過熱フィルター必要
'   ・専門家①（トレンドフォロー）：GCルックバック短縮、緩和モード厳格化
'   ・専門家②（リスク管理）：RR1:2統一（買売とも TP8%/SL4%）
'   ・専門家③（日本株専門）：RSI≥80売りエントリー除外、売りスコア閾値引上
'
' 適用方法:
'   1. Alt+F11 でVBAエディタを開く
'   2. このファイルをインポート (ファイル → ファイルのインポート)
'   3. イミディエイトウィンドウ (Ctrl+G) で
'      「Call 金のこづちパッチ適用」と入力してEnter
'   4. 完了メッセージを確認後、このモジュールを削除してOK
'=============================================================

Public Sub 金のこづちパッチ適用()
    Dim vbProj As Object
    Dim modBuy  As Object
    Dim modSell As Object

    On Error GoTo ErrHandler

    Set vbProj = ThisWorkbook.VBProject

    Dim comp As Object
    For Each comp In vbProj.VBComponents
        If comp.name = "Mod_買抽出v13" Then Set modBuy = comp.CodeModule
        If comp.name = "Mod_SellExtrac" Then Set modSell = comp.CodeModule
    Next comp

    If modBuy Is Nothing Then MsgBox "Mod_買抽出v13 が見つかりません", vbCritical: Exit Sub
    If modSell Is Nothing Then MsgBox "Mod_SellExtrac が見つかりません", vbCritical: Exit Sub

    Dim log As String: log = ""
    Dim cnt As Long:  cnt = 0

    '===========================================================
    ' 【買い】修正1: RSI≥70 ハードフィルター追加
    '  根拠: RSI過熱銘柄はエントリー後の調整リスクが高い
    '        現行はスコア-1点のみで除外できていない
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
           "        ' [修正1] RSI過熱（≥70）はハード除外（専門家①②共通意見）" & vbCrLf & _
           "        If rsiVal >= 70 Then canExtract = False"

    If ReplaceInModule(modBuy, old1, new1) Then
        log = log & "✅ 買い修正1: RSI≥70ハードフィルター 追加" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "⚠️ 買い修正1: 適用済みの可能性" & vbCrLf
    End If

    '===========================================================
    ' 【買い】修正2: 緩和モードでも momCount≥2 を要求
    '  根拠: 緩和モード時（TOPIX上昇日）に8→27銘柄に急増するのを抑制
    '        専門家①「トレンド一致なしはエントリー見送り」
    '===========================================================
    Dim old2 As String, new2 As String
    old2 = "        If doMark And canExtract Then" & vbCrLf & _
           "            If momCount < BUY_REQUIRE_MOMENTUM_COUNT Then canExtract = False" & vbCrLf & _
           "        End If"
    new2 = "        ' [修正2] 緩和モードでも最低momCount≥2を要求（過剰抽出を抑制）" & vbCrLf & _
           "        If canExtract And momCount < 2 Then canExtract = False" & vbCrLf & _
           "        If doMark And canExtract Then" & vbCrLf & _
           "            If momCount < BUY_REQUIRE_MOMENTUM_COUNT Then canExtract = False" & vbCrLf & _
           "        End If"

    If ReplaceInModule(modBuy, old2, new2) Then
        log = log & "✅ 買い修正2: 緩和モードmomCount≥2 追加" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "⚠️ 買い修正2: 適用済みの可能性" & vbCrLf
    End If

    '===========================================================
    ' 【買い】修正3: 緩和モード（TOPIX上昇日）のスコア閾値を+2引き上げ
    '  根拠: TOPIX上昇日は地合い良く「なんでも上がる」状況
    '        真に強い銘柄のみ抽出するためスコア閾値を高める
    '        専門家①「地合いに乗じた玉石混交エントリーを避ける」
    '        専門家②「銘柄数を10以内に絞り集中管理」
    '===========================================================
    Dim old3 As String, new3 As String
    old3 = "                Case Else" & vbCrLf & _
           "                    canExtract = True" & vbCrLf & _
           "            End Select"
    new3 = "                Case Else" & vbCrLf & _
           "                    ' [修正3] 緩和モードはスコア閾値+2（専門家①②）" & vbCrLf & _
           "                    If totalScore >= minScore + 2 Then canExtract = True" & vbCrLf & _
           "            End Select"

    If ReplaceInModule(modBuy, old3, new3) Then
        log = log & "✅ 買い修正3: 緩和モードスコア閾値+2 変更" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "⚠️ 買い修正3: 対象未発見（確認要）" & vbCrLf
    End If

    '===========================================================
    ' 【買い】修正4: GCルックバック 5→3日（Const変更）
    '  根拠: 5日前のGCはシグナルが古く、その後に上昇の勢いが失速している
    '        専門家①「直近3日以内のGCのみ有効と判断」
    '===========================================================
    Dim old4 As String, new4 As String
    old4 = "Public Const GOLDEN_CROSS_LOOKBACK As Long = 5"
    new4 = "Public Const GOLDEN_CROSS_LOOKBACK As Long = 3  ' [修正4] 5→3日（専門家①）"

    If ReplaceInModule(modBuy, old4, new4) Then
        log = log & "✅ 買い修正4: GCルックバック 5→3日 変更" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "⚠️ 買い修正4: 適用済みの可能性" & vbCrLf
    End If

    '===========================================================
    ' 【買い】修正5a: 利確ライン +5%→+8%（writeRow）
    '  根拠: RR比を1:1.5から1:2に改善（専門家②「最低RR1:2が鉄則」）
    '===========================================================
    Dim old5a As String, new5a As String
    old5a = "        .Cells(r, 15).Value = Format(closePrice * 1.05, ""#,##0"")"
    new5a = "        .Cells(r, 15).Value = Format(closePrice * 1.08, ""#,##0"")  ' [修正5a] +5%→+8%（RR1:2）"

    If ReplaceInModule(modBuy, old5a, new5a) Then
        log = log & "✅ 買い修正5a: 利確ライン +5%→+8% 変更" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "⚠️ 買い修正5a: 適用済みの可能性" & vbCrLf
    End If

    '===========================================================
    ' 【買い】修正5b: 損切りライン -3%→-4%（writeRow）
    '  根拠: 自動注文決済のSL=4%と表示を統一、RR計算を正確に
    '===========================================================
    Dim old5b As String, new5b As String
    old5b = "        .Cells(r, 16).Value = Format(closePrice * 0.97, ""#,##0"")"
    new5b = "        .Cells(r, 16).Value = Format(closePrice * 0.96, ""#,##0"")  ' [修正5b] -3%→-4%（SLと統一）"

    If ReplaceInModule(modBuy, old5b, new5b) Then
        log = log & "✅ 買い修正5b: 損切りライン -3%→-4% 変更" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "⚠️ 買い修正5b: 適用済みの可能性" & vbCrLf
    End If

    '===========================================================
    ' 【売り】修正S1: S_MIN_SCORE 10→11（スコア閾値引き上げ）
    '  根拠: 売り抽出は買いより条件が厳しくあるべき
    '        空売り制約（貸借銘柄）+ 踏み上げリスクを考慮
    '        専門家③「信用売り残高が多い銘柄ほど踏み上げリスク大」
    '===========================================================
    Dim oldS1 As String, newS1 As String
    oldS1 = "Private Const S_MIN_SCORE  As Long = 10  ' 抽出最低スコア"
    newS1 = "Private Const S_MIN_SCORE  As Long = 11  ' [修正S1] 10→11（専門家③）抽出最低スコア"

    If ReplaceInModule(modSell, oldS1, newS1) Then
        log = log & "✅ 売り修正S1: S_MIN_SCORE 10→11 変更" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "⚠️ 売り修正S1: 対象未発見（確認要）" & vbCrLf
    End If

    '===========================================================
    ' 【売り】修正S2: RSI<15 ハードフィルター追加
    '  根拠: RSI極値（<15）は売られ過ぎで買い戻しリスクが極めて高い
    '        新規空売りエントリーは危険（専門家②「RR崩壊域）
    '===========================================================
    Dim oldS2 As String, newS2 As String
    oldS2 = "        If doMark And canExtract Then" & vbCrLf & _
            "            If momCount < 2 Then canExtract = False" & vbCrLf & _
            "            If S_REQUIRE_DOWNTREND And Not ok_PD Then canExtract = False" & vbCrLf & _
            "        End If"
    newS2 = "        ' [修正S2] RSI極値（<15）は踏み上げリスクで除外（専門家②③）" & vbCrLf & _
            "        If canExtract And rsiVal < 15 Then canExtract = False" & vbCrLf & _
            "        ' [修正S3] RSI≥80は過熱天井だが踏み上げ加速帯 → 新規エントリー除外" & vbCrLf & _
            "        If canExtract And rsiVal >= 80 Then canExtract = False" & vbCrLf & _
            "        If doMark And canExtract Then" & vbCrLf & _
            "            If momCount < 2 Then canExtract = False" & vbCrLf & _
            "            If S_REQUIRE_DOWNTREND And Not ok_PD Then canExtract = False" & vbCrLf & _
            "        End If"

    Dim replacedS2 As Long: replacedS2 = 0
    Do
        If ReplaceInModule(modSell, oldS2, newS2) Then
            replacedS2 = replacedS2 + 1
        Else
            Exit Do
        End If
    Loop Until replacedS2 >= 2

    If replacedS2 > 0 Then
        log = log & "✅ 売り修正S2/S3: RSI<15除外＋RSI≥80除外 追加（" & replacedS2 & "箇所）" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "⚠️ 売り修正S2/S3: 適用済みの可能性" & vbCrLf
    End If

    '===========================================================
    ' 【売り】修正S4a: 利確ライン -5%→-8%（S_WriteRow）
    '  根拠: 買い側と同様にRR1:2統一（専門家②）
    '===========================================================
    Dim oldS4a As String, newS4a As String
    oldS4a = "        .Cells(r, 15).Value = Round(closePrice * 0.95, 0)  'O:利確ライン -5%"
    newS4a = "        .Cells(r, 15).Value = Round(closePrice * 0.92, 0)  'O:[修正S4a]利確ライン -8%（RR1:2）"

    Dim replacedS4a As Long: replacedS4a = 0
    Do
        If ReplaceInModule(modSell, oldS4a, newS4a) Then
            replacedS4a = replacedS4a + 1
        Else
            Exit Do
        End If
    Loop Until replacedS4a >= 2

    If replacedS4a > 0 Then
        log = log & "✅ 売り修正S4a: 利確ライン -5%→-8% 変更（" & replacedS4a & "箇所）" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "⚠️ 売り修正S4a: 適用済みの可能性" & vbCrLf
    End If

    '===========================================================
    ' 【売り】修正S4b: 損切りライン +3%→+4%（S_WriteRow）
    '  根拠: RR1:2統一、+4%SLで踏み上げ許容範囲を確保（専門家②）
    '===========================================================
    Dim oldS4b As String, newS4b As String
    oldS4b = "        .Cells(r, 16).Value = Round(closePrice * 1.03, 0)  'P:損切ライン +3%"
    newS4b = "        .Cells(r, 16).Value = Round(closePrice * 1.04, 0)  'P:[修正S4b]損切ライン +4%（RR1:2）"

    Dim replacedS4b As Long: replacedS4b = 0
    Do
        If ReplaceInModule(modSell, oldS4b, newS4b) Then
            replacedS4b = replacedS4b + 1
        Else
            Exit Do
        End If
    Loop Until replacedS4b >= 2

    If replacedS4b > 0 Then
        log = log & "✅ 売り修正S4b: 損切りライン +3%→+4% 変更（" & replacedS4b & "箇所）" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "⚠️ 売り修正S4b: 適用済みの可能性" & vbCrLf
    End If

    '===========================================================
    ' 完了メッセージ
    '===========================================================
    MsgBox "【金のこづち 0622 パッチ適用結果】" & vbCrLf & vbCrLf & log & vbCrLf & _
           "適用件数: " & cnt & " / 10" & vbCrLf & vbCrLf & _
           "■ 改善内容サマリー" & vbCrLf & _
           "  買い: RSI≥70除外 / GC3日/緩和スコア+2 / momCount≥2 / TP8% SL4%" & vbCrLf & _
           "  売り: RSI<15除外 / RSI≥80除外 / スコア≥11 / TP8% SL4%" & vbCrLf & vbCrLf & _
           "※ 自動注文決済: SL4% / TP8% / GU上限3% は設定済みです" & vbCrLf & _
           "  （適用後、このモジュールを削除してください）", _
           vbInformation, "金のこづち 0622 パッチ完了"
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
