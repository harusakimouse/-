Attribute VB_Name = "Mod_銘柄名キャッシュ"
'==============================================================================
' Mod_銘柄名キャッシュ  v1.0 (2026/08/31)
'
' 【何のためのモジュールか】
'   このBOOKの銘柄名は、すべて 銘柄管理!C列 の
'       =@RssMarket($B6,"銘柄名称")
'   という「RSSが生きているときしか値が返らない式」だけが出所でした。
'   終値!B列 も 分析!D列 も、中身は 銘柄管理!C列 への参照です。
'
'   そのためRSSが止まっていると、銘柄名の出所が全滅します。
'   実際にこのBOOKでは 銘柄管理!C列 が全行 "エラー" になっており、
'   終値!B列 のうち 87銘柄分 が式のままなので、同じく "エラー" になります。
'
'   このモジュールは「一度でも取れた正しい銘柄名」を
'   銘柄管理シートの AD列（既定・非表示）に文字列として貯めておき、
'   RSSが死んでいる時間帯でもそこから名前を出せるようにします。
'
' 【使い方】
'   ・RSSが繋がっている取引時間中に一度 銘柄名キャッシュ更新 を実行してください。
'     （売り抽出／買い抽出からも自動で呼ばれるので、普通は意識不要です）
'   ・キャッシュは既に正しく入っている名前を、空や "エラー" で上書きしません。
'==============================================================================
Option Explicit

' 銘柄管理シートの列
Public Const NC_SHEET      As String = "銘柄管理"
Public Const NC_CODE_COL   As Long = 2      'B列：コード
Public Const NC_NAME_COL   As Long = 3      'C列：銘柄名（RSS式）
Public Const NC_CACHE_COL  As Long = 30     'AD列：★名称キャッシュ（このモジュールが管理）
Public Const NC_HDR_ROW    As Long = 4
Public Const NC_FIRST_ROW  As Long = 6      '5行目はTOPIX行なので6行目から
Private Const NC_PW        As String = "ne19480314"

'==============================================================================
' 名称として使えない値かどうかを判定する
'   "" / "エラー" / "#N/A" などのエラー文字 / 数値 は名称ではない。
'   ★数値をはじくのが重要：以前は名前が取れないと "(3936)" と書いていたため、
'     Excelが会計表記とみなして -3936 という数値に化けていました。
'==============================================================================
Public Function NC_名称が無効(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    t = Replace(t, ChrW(&H3000), "")          '全角スペースだけの値も無効扱い
    NC_名称が無効 = True
    If t = "" Then Exit Function
    If Left$(t, 1) = "#" Then Exit Function                 '#N/A #VALUE! #REF! …
    If t = "エラー" Or UCase$(t) = "ERROR" Then Exit Function 'RSSが返すエラー文字
    If t = "取得中" Or t = "待機中" Then Exit Function        'RSS初期化中の暫定値
    If IsNumeric(t) Then Exit Function                       '数値は名称ではない
    NC_名称が無効 = False
End Function

Public Function NC_安全文字(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        NC_安全文字 = ""
    Else
        NC_安全文字 = Trim$(CStr(v))
    End If
End Function

Private Function NC_Ws(ByVal n As String) As Worksheet
    On Error Resume Next
    Set NC_Ws = ThisWorkbook.Sheets(n)
    On Error GoTo 0
End Function

'==============================================================================
' キャッシュ列の見出しを整える（初回だけ意味がある）
'==============================================================================
Private Sub NC_見出し整備(ByVal mws As Worksheet)
    On Error Resume Next
    With mws.Cells(NC_HDR_ROW, NC_CACHE_COL)
        If NC_安全文字(.Value) = "" Then .Value = "名称キャッシュ"
        .Font.Bold = True
    End With
    mws.Columns(NC_CACHE_COL).NumberFormat = "@"
    mws.Columns(NC_CACHE_COL).Hidden = True
    On Error GoTo 0
End Sub

'==============================================================================
' ■ 銘柄名キャッシュ更新（ボタン割り当て可）
'   銘柄管理!C列 / 終値!B列 / 分析!D列 のうち「まともな値」を拾って
'   銘柄管理!AD列 に文字列で保存する。
'   すでにキャッシュに正しい名前があるものは、無効値では上書きしない。
'   戻り値 = 新しく保存できた件数
'==============================================================================
Public Function 銘柄名キャッシュ更新(Optional ByVal silent As Boolean = True) As Long
    Dim mws As Worksheet: Set mws = NC_Ws(NC_SHEET)
    If mws Is Nothing Then Exit Function

    On Error Resume Next
    mws.Unprotect Password:=NC_PW
    On Error GoTo 0
    NC_見出し整備 mws

    Dim clWs As Worksheet: Set clWs = NC_Ws("終値")
    Dim anWs As Worksheet: Set anWs = NC_Ws("分析")
    Dim kjWs As Worksheet: Set kjWs = NC_Ws("検証明細")   '過去の検証明細にも銘柄名が残っている

    '各シートを「コード → 行」で引けるようにしておく
    Dim clDic As Object: Set clDic = NC_行辞書(clWs, 1, 6)
    Dim anDic As Object: Set anDic = NC_行辞書(anWs, 3, 3)
    Dim kjDic As Object: Set kjDic = NC_行辞書(kjWs, 3, 4)

    Dim last As Long
    last = mws.Cells(mws.Rows.Count, NC_CODE_COL).End(xlUp).Row
    If last < NC_FIRST_ROW Then GoTo Fin

    Dim r As Long, added As Long
    For r = NC_FIRST_ROW To last
        Dim code As String: code = NC_安全文字(mws.Cells(r, NC_CODE_COL).Value)
        If code <> "" And code <> "TOPX" Then
            Dim cur As String: cur = NC_安全文字(mws.Cells(r, NC_CACHE_COL).Value)
            If NC_名称が無効(cur) Then
                Dim nm As String
                nm = NC_安全文字(mws.Cells(r, NC_NAME_COL).Value)          '① 銘柄管理C列(RSS)
                If NC_名称が無効(nm) Then nm = NC_他シートから(clWs, clDic, code, 2)  '② 終値B列
                If NC_名称が無効(nm) Then nm = NC_他シートから(anWs, anDic, code, 4)  '③ 分析D列
                If NC_名称が無効(nm) Then nm = NC_他シートから(kjWs, kjDic, code, 4)  '④ 検証明細D列
                If Not NC_名称が無効(nm) Then
                    mws.Cells(r, NC_CACHE_COL).NumberFormat = "@"
                    mws.Cells(r, NC_CACHE_COL).Value = nm
                    added = added + 1
                End If
            End If
        End If
    Next r

Fin:
    銘柄名キャッシュ更新 = added
    If Not silent Then
        MsgBox "銘柄名キャッシュを更新しました。" & vbCrLf & _
               "新しく保存できた銘柄名: " & added & " 件" & vbCrLf & vbCrLf & _
               "※ まだ埋まっていない銘柄は、RSSが繋がっている取引時間中に" & vbCrLf & _
               "　 もう一度このボタンを押すと保存されます。", _
               vbInformation, "銘柄名キャッシュ"
    End If
End Function

Private Function NC_行辞書(ByVal ws As Worksheet, ByVal codeCol As Long, _
                            ByVal firstRow As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set NC_行辞書 = d
    If ws Is Nothing Then Exit Function
    Dim last As Long: last = ws.Cells(ws.Rows.Count, codeCol).End(xlUp).Row
    Dim r As Long
    For r = firstRow To last
        Dim c As String: c = NC_安全文字(ws.Cells(r, codeCol).Value)
        If c <> "" Then
            If Not d.Exists(c) Then d.Add c, r
        End If
    Next r
End Function

Private Function NC_他シートから(ByVal ws As Worksheet, ByVal dic As Object, _
                                  ByVal code As String, ByVal nameCol As Long) As String
    NC_他シートから = ""
    If ws Is Nothing Or dic Is Nothing Then Exit Function
    If Not dic.Exists(code) Then Exit Function
    NC_他シートから = NC_安全文字(ws.Cells(CLng(dic(code)), nameCol).Value)
End Function

'==============================================================================
' ■ 名称辞書作成（抽出処理から呼ぶ本体）
'   コード → 銘柄名 の辞書を返す。
'   キャッシュを優先し、RSSが生きていればその値でキャッシュを育てる。
'==============================================================================
Public Function 名称辞書作成() As Object
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Set 名称辞書作成 = dic

    Dim mws As Worksheet: Set mws = NC_Ws(NC_SHEET)
    If mws Is Nothing Then Exit Function

    '取れる範囲でキャッシュを育ててから読む（RSSが死んでいても実害なし）
    On Error Resume Next
    銘柄名キャッシュ更新 True
    On Error GoTo 0

    Dim last As Long
    last = mws.Cells(mws.Rows.Count, NC_CODE_COL).End(xlUp).Row
    If last < NC_FIRST_ROW Then Exit Function

    Dim r As Long
    For r = NC_FIRST_ROW To last
        Dim code As String: code = NC_安全文字(mws.Cells(r, NC_CODE_COL).Value)
        If code <> "" And code <> "TOPX" Then
            Dim nm As String
            nm = NC_安全文字(mws.Cells(r, NC_NAME_COL).Value)        'RSSが生きていれば最新
            If NC_名称が無効(nm) Then nm = NC_安全文字(mws.Cells(r, NC_CACHE_COL).Value)
            If NC_名称が無効(nm) Then nm = ""
            If Not dic.Exists(code) Then dic.Add code, nm
        End If
    Next r
End Function

'==============================================================================
' ■ 銘柄名を1件解決する
'   辞書 → 元シートの名称列 → 「コード（名称未取得）」の順で必ず文字列を返す。
'   ★ここで絶対に "(3936)" 形式を返さないこと（Excelが -3936 に変換するため）
'==============================================================================
Public Function 銘柄名解決(ByVal code As String, ByVal nameDic As Object, _
                            Optional ByVal srcWs As Worksheet, _
                            Optional ByVal srcRow As Long = 0, _
                            Optional ByVal srcNameCol As Long = 2) As String
    Dim nm As String: nm = ""

    If Not nameDic Is Nothing Then
        If nameDic.Exists(code) Then nm = NC_安全文字(nameDic(code))
    End If

    If NC_名称が無効(nm) Then
        If Not srcWs Is Nothing And srcRow > 0 Then
            nm = NC_安全文字(srcWs.Cells(srcRow, srcNameCol).Value)
        End If
    End If

    If NC_名称が無効(nm) Then
        '全角カッコを使う。半角 "(1234)" はExcelが負数に変換してしまう。
        nm = code & ChrW(&HFF08) & "名称未取得" & ChrW(&HFF09)
    End If

    銘柄名解決 = nm
End Function

'==============================================================================
' ■ 名称セルへの書き込み（必ず文字列として書く）
'   数値化・日付化を防ぐため、書く直前に表示形式を「文字列」に固定する。
'==============================================================================
Public Sub 名称セル書込(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, _
                        ByVal nm As String)
    With ws.Cells(r, c)
        .NumberFormat = "@"
        .Value = nm
    End With
End Sub

'==============================================================================
' ■ 既に -3936 のように数値化してしまった名称セルを修復する
'   買抽出v13 / 売抽出v13 / 厳選TOP2 のD列を見て、負の数値になっているものを
'   キャッシュの銘柄名に戻す。
'==============================================================================
Public Sub 名称セル修復_全抽出シート()
    Dim names As Variant
    names = Array("買抽出v13", "売抽出v13", "厳選TOP2")
    Dim dic As Object: Set dic = 名称辞書作成()

    Dim i As Long, fixed As Long
    For i = LBound(names) To UBound(names)
        Dim ws As Worksheet: Set ws = NC_Ws(CStr(names(i)))
        If Not ws Is Nothing Then
            On Error Resume Next
            ws.Unprotect Password:=NC_PW
            On Error GoTo 0
            Dim last As Long
            last = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
            Dim r As Long
            For r = 4 To last
                Dim code As String: code = NC_安全文字(ws.Cells(r, 3).Value)
                If code <> "" Then
                    Dim cur As String: cur = NC_安全文字(ws.Cells(r, 4).Value)
                    If NC_名称が無効(cur) Then
                        名称セル書込 ws, r, 4, 銘柄名解決(code, dic)
                        fixed = fixed + 1
                    End If
                End If
            Next r
        End If
    Next i

    MsgBox "銘柄名セルを修復しました： " & fixed & " 件", vbInformation, "銘柄名キャッシュ"
End Sub
