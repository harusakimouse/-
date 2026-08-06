Attribute VB_Name = "modRiskGate"
'==============================================================================
' modRiskGate  ―  打ち出の小づち V806 リスク管理モジュール
'
'  V805 の問題点に対する処置：
'   (1) スコアの合算だけで採用を決めていたため、過熱銘柄が高得点で通過していた
'       → 減点では防げないので「ハードゲート」で足切りする
'   (2) 相場環境シグナル（"相場上昇・見送り"）が算出されているのに無視されていた
'       → ゲートで強制的に反映する
'   (3) 全銘柄一律100株のため1トレードのリスク額が最大145倍ばらついていた
'       → リスク額を固定して株数を逆算する
'   (4) 同一銘柄への連続エントリー（ナンピン）を止める仕組みが無かった
'       → 建玉存在チェックを入れる
'   (5) 1日あたりの新規建玉数・総リスクに上限が無かった
'       → 日次上限を設ける
'
'  使い方： 標準モジュールとしてインポートし、発注リスト生成の直前に
'           CheckEntry() を呼んで False なら採用しない。
'==============================================================================
Option Explicit

'------------------------------------------------------------ 設定（ここを触る）
Public Const RISK_PER_TRADE   As Double = 20000   ' 1トレードの許容損失額（円）
Public Const STOP_PCT         As Double = 0.03    ' 損切幅（MAE分析より。V806のAG2と揃えること）
Public Const MAX_NEW_PER_DAY  As Long = 3         ' 1日の新規建玉数の上限
Public Const MAX_OPEN_TOTAL   As Long = 8         ' 同時保有の上限
Public Const MAX_DAILY_RISK   As Double = 60000   ' 1日に晒す合計リスク額の上限（円）

' 買いゲート
Public Const BUY_RSI_MAX      As Double = 70      ' これ以上は過熱として採用しない
Public Const BUY_RSI_MIN      As Double = 45      ' これ未満はトレンド不成立
Public Const BUY_VOL_MIN      As Double = 1.5     ' 出来高倍率の下限

' 売りゲート
Public Const SELL_RSI_MIN     As Double = 30      ' これ以下は売られ過ぎ＝踏み上げリスク
Public Const SELL_RSI_MAX     As Double = 55
Public Const SELL_VOL_MIN     As Double = 1.5
Public Const STOP_TP          As Double = 0.08    ' 利確幅（V806のAG3と揃えること）

Private Const SH_KANRI        As String = "管理"
Private Const ROW_FIRST       As Long = 4


'==============================================================================
' エントリー可否判定。採用してよければ True、見送るなら False を返し
' rejectReason に理由を格納する。
'==============================================================================
Public Function CheckEntry(ByVal code As String, _
                           ByVal side As String, _
                           ByVal rsi As Double, _
                           ByVal volRatio As Double, _
                           ByVal emaTrend As String, _
                           ByVal macd As Double, _
                           ByVal finalSignal As String, _
                           ByRef rejectReason As String) As Boolean

    rejectReason = ""

    '--- ゲート0: システム自身が「見送り」と判定しているものは採用しない -------
    '    V805 ではこのシグナルが算出されながら発注に反映されておらず、
    '    8/7 の ＳＵＭＣＯ・ミネベアミツミ はここで止まるはずだった。
    If InStr(finalSignal, "見送り") > 0 Then
        rejectReason = "最終シグナルが見送り: " & finalSignal
        Exit Function
    End If

    '--- ゲート1: 建玉数・重複 ------------------------------------------------
    If CountOpenPositions() >= MAX_OPEN_TOTAL Then
        rejectReason = "同時保有上限 " & MAX_OPEN_TOTAL & " 件に到達"
        Exit Function
    End If
    If CountNewToday() >= MAX_NEW_PER_DAY Then
        rejectReason = "本日の新規建玉が上限 " & MAX_NEW_PER_DAY & " 件に到達"
        Exit Function
    End If
    If HasOpenPosition(code, side) Then
        rejectReason = code & " は同方向の建玉を保有中（ナンピン禁止）"
        Exit Function
    End If
    If WasStoppedRecently(code, 5) Then
        rejectReason = code & " は直近5営業日に損切済み（再エントリー禁止）"
        Exit Function
    End If

    '--- ゲート2: 方向別のテクニカル条件 --------------------------------------
    If side = "買" Then
        If rsi >= BUY_RSI_MAX Then
            rejectReason = "RSI " & Format(rsi, "0.0") & " が過熱（" & BUY_RSI_MAX & "以上）"
            Exit Function
        End If
        If rsi < BUY_RSI_MIN Then
            rejectReason = "RSI " & Format(rsi, "0.0") & " が弱すぎる（" & BUY_RSI_MIN & "未満）"
            Exit Function
        End If
        If volRatio < BUY_VOL_MIN Then
            rejectReason = "出来高倍率 " & Format(volRatio, "0.00") & " が不足"
            Exit Function
        End If
        If emaTrend <> "パーフェクト▲" And emaTrend <> "上昇トレンド▲" Then
            rejectReason = "EMAトレンドが上昇でない（" & emaTrend & "）"
            Exit Function
        End If
        If macd <= 0 Then
            rejectReason = "MACDがマイナス"
            Exit Function
        End If
    Else
        If rsi <= SELL_RSI_MIN Then
            rejectReason = "RSI " & Format(rsi, "0.0") & " が売られ過ぎ（踏み上げリスク）"
            Exit Function
        End If
        If rsi > SELL_RSI_MAX Then
            rejectReason = "RSI " & Format(rsi, "0.0") & " が高い（下降トレンド不成立）"
            Exit Function
        End If
        If volRatio < SELL_VOL_MIN Then
            rejectReason = "出来高倍率 " & Format(volRatio, "0.00") & " が不足"
            Exit Function
        End If
        If emaTrend <> "下降トレンド▼" Then
            rejectReason = "EMAトレンドが下降でない（" & emaTrend & "）"
            Exit Function
        End If
        If macd >= 0 Then
            rejectReason = "MACDがプラス"
            Exit Function
        End If
    End If

    CheckEntry = True
End Function


'==============================================================================
' リスク額を一定にした発注株数。単元(100株)未満は切り捨て、最低1単元。
' 建値が高い銘柄ほど株数が減り、1トレードの損失額が RISK_PER_TRADE に揃う。
'==============================================================================
Public Function CalcQty(ByVal entryPrice As Double, _
                        Optional ByVal riskYen As Double = RISK_PER_TRADE, _
                        Optional ByVal stopPct As Double = STOP_PCT) As Long
    Dim units As Long
    If entryPrice <= 0 Or stopPct <= 0 Then Exit Function
    units = Int(riskYen / (entryPrice * stopPct) / 100)
    If units < 1 Then units = 1          ' 1単元でも上限超過なら見送り判断は呼び出し側で
    CalcQty = units * 100
End Function


'==============================================================================
' 本日これから建てる分を含めた合計リスク額が上限内かを確認する
'==============================================================================
Public Function FitsDailyRisk(ByVal entryPrice As Double, ByVal qty As Long) As Boolean
    FitsDailyRisk = (TodayRiskYen() + entryPrice * STOP_PCT * qty) <= MAX_DAILY_RISK
End Function


'------------------------------------------------------------------- 内部関数
Private Function LastRow() As Long
    LastRow = ThisWorkbook.Sheets(SH_KANRI).Cells(ThisWorkbook.Sheets(SH_KANRI).Rows.Count, 2).End(xlUp).Row
End Function

Private Function CountOpenPositions() As Long
    Dim ws As Worksheet, r As Long, n As Long
    Set ws = ThisWorkbook.Sheets(SH_KANRI)
    For r = ROW_FIRST To LastRow()
        If ws.Cells(r, 2).Value <> "" And InStr(ws.Cells(r, 19).Value, "保有") > 0 Then n = n + 1
    Next r
    CountOpenPositions = n
End Function

Private Function CountNewToday() As Long
    Dim ws As Worksheet, r As Long, n As Long
    Set ws = ThisWorkbook.Sheets(SH_KANRI)
    For r = ROW_FIRST To LastRow()
        If IsDate(ws.Cells(r, 5).Value) Then
            If CDate(ws.Cells(r, 5).Value) = Date Then n = n + 1
        End If
    Next r
    CountNewToday = n
End Function

Private Function HasOpenPosition(ByVal code As String, ByVal side As String) As Boolean
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets(SH_KANRI)
    For r = ROW_FIRST To LastRow()
        If CStr(ws.Cells(r, 2).Value) = code _
           And CStr(ws.Cells(r, 3).Value) = side _
           And InStr(ws.Cells(r, 19).Value, "保有") > 0 Then
            HasOpenPosition = True
            Exit Function
        End If
    Next r
End Function

' 直近 nDays 営業日以内に同一銘柄を損切していたら True
' （V805 では 助川電気工業 を 8/4 に -16,800円 で損切した3営業日後に同方向で再エントリーしていた）
Private Function WasStoppedRecently(ByVal code As String, ByVal nDays As Long) As Boolean
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Sheets(SH_KANRI)
    For r = ROW_FIRST To LastRow()
        If CStr(ws.Cells(r, 2).Value) = code And InStr(ws.Cells(r, 24).Value, "損切") > 0 Then
            If IsDate(ws.Cells(r, 20).Value) Then
                If Date - CDate(ws.Cells(r, 20).Value) <= nDays Then
                    WasStoppedRecently = True
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Function TodayRiskYen() As Double
    Dim ws As Worksheet, r As Long, t As Double
    Set ws = ThisWorkbook.Sheets(SH_KANRI)
    For r = ROW_FIRST To LastRow()
        If IsDate(ws.Cells(r, 5).Value) Then
            If CDate(ws.Cells(r, 5).Value) = Date And IsNumeric(ws.Cells(r, 6).Value) Then
                t = t + ws.Cells(r, 6).Value * STOP_PCT * ws.Cells(r, 7).Value
            End If
        End If
    Next r
    TodayRiskYen = t
End Function


'==============================================================================
' 既存の管理シートに対して、現行の建玉がゲートを通るかを一括点検する。
' 実行するとイミディエイトウィンドウに却下理由が出る。
'==============================================================================
Public Sub AuditOpenPositions()
    Dim ws As Worksheet, r As Long, reason As String, ng As Long
    Set ws = ThisWorkbook.Sheets(SH_KANRI)
    Debug.Print "=== 建玉ゲート点検 " & Format(Now, "yyyy/mm/dd hh:nn") & " ==="
    For r = ROW_FIRST To LastRow()
        If InStr(ws.Cells(r, 19).Value, "保有") > 0 Then
            If ws.Cells(r, 6).Value = "" Then
                Debug.Print "行" & r & " " & ws.Cells(r, 4).Value & ": 購入価格が空欄。損切ライン算出不能"
                ng = ng + 1
            End If
        End If
    Next r
    Debug.Print "要修正 " & ng & " 件"
End Sub


'==============================================================================
'  以下は V805 で発見した既存VBAのバグに対する置換コードです。
'  該当モジュールの当該箇所を、それぞれ差し替えてください。
'==============================================================================

'------------------------------------------------------------------------------
' 【バグ1・最優先】Module1.損切りアラート が一度も発火しない
'
'  現行:  If InStr(ステータス, "損切推奨") > 0 Or InStr(ステータス, "STOP推奨") > 0 Then
'
'  S列が生成する文字列は「🔴損切り」「⚠️損切注意」であり、「損切推奨」は
'  旧22列レイアウト時代の文字列。列参照だけ新レイアウトに直され、比較文字列が
'  取り残された。結果、損失拡大中でも起動時に「全ポジション正常です」を表示する。
'
'  置換後（部分一致にする）:
'     If InStr(ステータス, "損切") > 0 Or InStr(ステータス, "STOP") > 0 Then
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' 【バグ2】Module6.追記_本日の候補 の重複チェックが機能していない
'
'  現行:  If WorksheetFunction.CountIfs(ws.Range("B:B"), cd, ws.Range("E:E"), bd) = 0 Then
'         → 「コード AND 購入日」の複合条件のため、日が変われば必ず通過する
'
'  下の IsEntryAllowed に置き換えること。
'------------------------------------------------------------------------------
Public Function IsEntryAllowed(ByVal ws As Worksheet, ByVal cd As String, _
                               ByVal bd As Double, ByRef reason As String) As Boolean
    Dim wf As WorksheetFunction: Set wf = Application.WorksheetFunction
    reason = ""

    ' ① 同一コード＋同一購入日（従来の条件）
    If wf.CountIfs(ws.Range("B:B"), cd, ws.Range("E:E"), bd) > 0 Then
        reason = "本日すでに追記済み"
        Exit Function
    End If

    ' ② 同一コードで未決済ポジションが残っている（T列=売却日が空）
    '    ダイヘン -100,500円、タカラスタンダード3連続はここで止まるはずだった
    If wf.CountIfs(ws.Range("B:B"), cd, ws.Range("T:T"), "") > 0 Then
        reason = "同一銘柄を保有中（ナンピン禁止）"
        Exit Function
    End If

    ' ③ 直近10営業日以内に同一コードでエントリー済み
    '    助川電気工業は損切の3営業日後に再エントリーされていた
    If wf.CountIfs(ws.Range("B:B"), cd, _
                   ws.Range("E:E"), ">=" & CDbl(wf.WorkDay(bd, -10))) > 0 Then
        reason = "直近10営業日に同一銘柄をエントリー済み"
        Exit Function
    End If

    IsEntryAllowed = True
End Function


'------------------------------------------------------------------------------
' 【バグ3】損切/利確価格が東証の呼値に乗っておらず、145件中71件(49%)が発注不可
'          → 逆指値が証券会社に弾かれ、損切注文が存在しないポジションになる
'
'  V806 のシート数式側では対応済み。VBAから価格を書き込む場合はこの関数を通すこと。
'------------------------------------------------------------------------------
Public Function SnapTick(ByVal p As Double, ByVal roundUp As Boolean) As Double
    Dim t As Double
    ' 東証の呼値（TOPIX100構成銘柄はより細かいため、別途対応が必要）
    Select Case p
        Case Is <= 3000#:   t = 1
        Case Is <= 5000#:   t = 5
        Case Is <= 30000#:  t = 10
        Case Is <= 50000#:  t = 50
        Case Is <= 300000#: t = 100
        Case Else:          t = 500
    End Select
    If roundUp Then
        SnapTick = -Int(-p / t) * t          ' 切り上げ
    Else
        SnapTick = Int(p / t) * t            ' 切り下げ
    End If
End Function

' 買い＝切り下げ／売り＝切り上げ。どちらも「利確は約定しやすい側・損切は余裕のある側」に倒れる。
Public Sub WriteExitLines(ByVal ws As Worksheet, ByVal r As Long)
    Dim px As Double, isSell As Boolean
    If Not IsNumeric(ws.Cells(r, "F").Value) Then Exit Sub
    px = CDbl(ws.Cells(r, "F").Value)
    isSell = (ws.Cells(r, "C").Value = "売")
    ws.Cells(r, "M").Value = SnapTick(px * IIf(isSell, 1 - STOP_TP, 1 + STOP_TP), isSell)
    ws.Cells(r, "N").Value = SnapTick(px * IIf(isSell, 1 + STOP_PCT, 1 - STOP_PCT), isSell)
End Sub


'------------------------------------------------------------------------------
' 【バグ4】抽出Sub 4本にエラーハンドラがなく、Calculation = xlCalculationManual が
'          残留しうる。残ると H列(RSS現在値)が更新されず、L/O/S/AH列が凍結し
'          損切判定が完全に沈黙する。トレードシステムとして最悪の失敗モード。
'
'  対象: Mod_買抽出_順張りVWAP.抽出実行_順張り / Mod_売抽出_下降VWAP.抽出実行_下降
'        Mod_買抽出v13.抽出実行           / Mod_SellExtrac.S_抽出実行
'
'  4本すべてを次の骨格で囲むこと。
'------------------------------------------------------------------------------
'   Private Sub 抽出実行_順張り(...)
'       Dim prevCalc As XlCalculation
'       prevCalc = Application.Calculation
'       On Error GoTo ErrHandler
'       … 既存処理 …
'       GoTo CleanExit
'   ErrHandler:
'       MsgBox "【抽出処理エラー】" & vbCrLf & _
'              "Err " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
'              "アプリ状態を復元しました。データは不完全な可能性があります。", _
'              vbCritical, "抽出中断"
'   CleanExit:
'       On Error Resume Next
'       Application.StatusBar = False
'       Application.Calculation = prevCalc
'       Application.ScreenUpdating = True
'       Application.EnableEvents = True
'       Application.CutCopyMode = False
'       On Error GoTo 0
'   End Sub


'------------------------------------------------------------------------------
' 【バグ5】Mod_KanriSheet は旧22列レイアウトのまま残存している死にモジュール。
'          管理シート作成 は Application.DisplayAlerts = False の後 ws.Delete を
'          実行するため、押すと全取引履歴が無警告で消滅する。
'
'  → モジュールごと削除するか、最低限ボタンの割り当てを外すこと。
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' 【バグ6】末尾行探索が Do While Cells(r,"B").Value <> "" で、B列に空白行が
'          1行でも混ざるとそこで停止し、以降の既存データを上書きする。
'------------------------------------------------------------------------------
Public Function NextWriteRow(ByVal ws As Worksheet) As Long
    NextWriteRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 1
    If NextWriteRow < ROW_FIRST Then NextWriteRow = ROW_FIRST
End Function
