Attribute VB_Name = "Mod_Patch_FuturesFilter_V1"
'=============================================================
' 日経225先物フィルター パッチ v1
'
' ■ 事前準備（Excelで手作業）:
'   銘柄管理シートに以下を追加:
'     B6 セル: 先物　（ラベル）
'     C6 セル: =RssIndexMarket("N225.FUT01.OS", "現在値")
'     D6 セル: =RssIndexMarket("N225.FUT01.OS", "前日比率")
'
' ■ 適用後の動作:
'   先物 < -1.5%  → TOPIX関係なく厳格モード強制（全銘柄スコア基準を最高に）
'   先物 < -0.5%  → 緩和モード禁止（MID以下に制限）
'   先物 ≥ -0.5%  → TOPIXの判定を通常通り使用
'
' ■ 適用方法:
'   1. 上記セルをExcelで先に設定
'   2. Alt+F11 でVBAエディタを開く
'   3. このファイルをインポート
'   4. イミディエイトウィンドウで「Call 先物フィルター適用」→ Enter
'   5. 完了後このモジュールを削除
'=============================================================

Public Sub ApplyFuturesFilter()
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
    ' 修正B: 買いモジュール 先物フィルター追加
    '  topixMode 決定直後に先物前日比率でオーバーライド
    '  読み取り元: 銘柄管理シート D6 (=RssIndexMarket("N225.FUT01.OS","前日比率"))
    '===========================================================
    Dim oldB As String, newB As String
    oldB = "    Dim topixMode As String" & vbCrLf & _
           "    If topixPct <= TOPIX_LARGE_DROP_PCT Then" & vbCrLf & _
           "        topixMode = ""LARGE""" & vbCrLf & _
           "    ElseIf topixPct <= TOPIX_SMALL_DROP_PCT Then" & vbCrLf & _
           "        topixMode = ""MID""" & vbCrLf & _
           "    Else" & vbCrLf & _
           "        topixMode = ""OK""" & vbCrLf & _
           "    End If"
    newB = "    Dim topixMode As String" & vbCrLf & _
           "    If topixPct <= TOPIX_LARGE_DROP_PCT Then" & vbCrLf & _
           "        topixMode = ""LARGE""" & vbCrLf & _
           "    ElseIf topixPct <= TOPIX_SMALL_DROP_PCT Then" & vbCrLf & _
           "        topixMode = ""MID""" & vbCrLf & _
           "    Else" & vbCrLf & _
           "        topixMode = ""OK""" & vbCrLf & _
           "    End If" & vbCrLf & _
           "    ' ★先物フィルター: 銘柄管理D6 =RssIndexMarket(""N225.FUT01.OS"",""前日比率"")" & vbCrLf & _
           "    Dim futuresPct As Double" & vbCrLf & _
           "    futuresPct = SafeNum(mws.Cells(6, 4).Value, 0)" & vbCrLf & _
           "    If futuresPct < -1.5 Then" & vbCrLf & _
           "        topixMode = ""LARGE""  ' 先物急落(-1.5%超): 問答無用で厳格モード" & vbCrLf & _
           "    ElseIf futuresPct < -0.5 Then" & vbCrLf & _
           "        If topixMode = ""OK"" Then topixMode = ""MID""  ' 先物軟調: 緩和モード禁止" & vbCrLf & _
           "    End If"

    If ReplaceInModule(modBuy, oldB, newB) Then
        log = log & "✅ 買い: 先物フィルター追加（D6参照）" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "⚠️ 買い: 対象コードが見つかりません（適用済みの可能性）" & vbCrLf
    End If

    '===========================================================
    ' 修正S: 売りモジュール 先物フィルター追加
    '  先物急騰時は新規空売りを禁止
    '  先物 > +1.5%  → 強い踏み上げリスク → 全銘柄売りNG
    '  先物 > +0.5%  → 緩和売りモード禁止
    '===========================================================
    Dim oldS As String, newS As String
    oldS = "    Dim topixMode As String" & vbCrLf & _
           "    If topixPct <= -1.5 Then" & vbCrLf & _
           "        topixMode = ""LARGE""" & vbCrLf & _
           "    ElseIf topixPct <= -0.5 Then" & vbCrLf & _
           "        topixMode = ""MID""" & vbCrLf & _
           "    Else" & vbCrLf & _
           "        topixMode = ""OK""" & vbCrLf & _
           "    End If"
    newS = "    Dim topixMode As String" & vbCrLf & _
           "    If topixPct <= -1.5 Then" & vbCrLf & _
           "        topixMode = ""LARGE""" & vbCrLf & _
           "    ElseIf topixPct <= -0.5 Then" & vbCrLf & _
           "        topixMode = ""MID""" & vbCrLf & _
           "    Else" & vbCrLf & _
           "        topixMode = ""OK""" & vbCrLf & _
           "    End If" & vbCrLf & _
           "    ' ★先物フィルター（売り）: 急騰時は空売り禁止" & vbCrLf & _
           "    Dim futuresPct As Double" & vbCrLf & _
           "    futuresPct = S_SafeNum(mws.Cells(6, 4).Value, 0)" & vbCrLf & _
           "    If futuresPct > 1.5 Then" & vbCrLf & _
           "        topixMode = ""LARGE""  ' 先物急騰(+1.5%超): 空売り全禁止" & vbCrLf & _
           "    ElseIf futuresPct > 0.5 Then" & vbCrLf & _
           "        If topixMode = ""OK"" Then topixMode = ""MID""  ' 先物強い: 売り緩和モード禁止" & vbCrLf & _
           "    End If"

    If ReplaceInModule(modSell, oldS, newS) Then
        log = log & "✅ 売り: 先物フィルター追加（急騰時空売り禁止）" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "⚠️ 売り: 対象コードが見つかりません（適用済みの可能性）" & vbCrLf
    End If

    '===========================================================
    ' 完了
    '===========================================================
    MsgBox "【先物フィルター パッチ適用結果】" & vbCrLf & vbCrLf & log & vbCrLf & _
           "適用件数: " & cnt & " / 2" & vbCrLf & vbCrLf & _
           "■ 動作確認チェックリスト" & vbCrLf & _
           "  □ 銘柄管理 B6: 先物（ラベル）" & vbCrLf & _
           "  □ 銘柄管理 C6: =RssIndexMarket(""N225.FUT01.OS"",""現在値"")" & vbCrLf & _
           "  □ 銘柄管理 D6: =RssIndexMarket(""N225.FUT01.OS"",""前日比率"")" & vbCrLf & _
           "  □ D6に数値が表示されている（例: 0.54）" & vbCrLf & vbCrLf & _
           "適用後このモジュールを削除してください", _
           vbInformation, "先物フィルター v1 完了"
    Exit Sub

ErrHandler:
    MsgBox "エラー " & Err.Number & ": " & Err.Description, vbCritical, "エラー"
End Sub

'-----------------------------------------------------------
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
