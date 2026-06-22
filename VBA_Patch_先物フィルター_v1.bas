Attribute VB_Name = "Mod_Patch_FuturesFilter_V1"

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

    If modBuy Is Nothing Then MsgBox "Mod_買抽出v13 not found", vbCritical: Exit Sub
    If modSell Is Nothing Then MsgBox "Mod_SellExtrac not found", vbCritical: Exit Sub

    Dim log As String: log = ""
    Dim cnt As Long:  cnt = 0

    ' --- BUY: add futures filter after topixMode block ---
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
           "    ' Futures filter: KanriSheet D6 = RssIndexMarket(N225.FUT01.OS, prev day %)" & vbCrLf & _
           "    Dim futuresPct As Double" & vbCrLf & _
           "    futuresPct = SafeNum(mws.Cells(5, 8).Value, 0)" & vbCrLf & _
           "    If futuresPct < -1.5 Then" & vbCrLf & _
           "        topixMode = ""LARGE""" & vbCrLf & _
           "    ElseIf futuresPct < -0.5 Then" & vbCrLf & _
           "        If topixMode = ""OK"" Then topixMode = ""MID""" & vbCrLf & _
           "    End If"

    If ReplaceInModule(modBuy, oldB, newB) Then
        log = log & "OK: Buy futures filter added" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "SKIP: Buy - already applied or not found" & vbCrLf
    End If

    ' --- SELL: add futures filter after topixMode block ---
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
           "    ' Futures filter for SELL: block short entry when futures surge" & vbCrLf & _
           "    Dim futuresPct As Double" & vbCrLf & _
           "    futuresPct = S_SafeNum(mws.Cells(5, 8).Value, 0)" & vbCrLf & _
           "    If futuresPct > 1.5 Then" & vbCrLf & _
           "        topixMode = ""LARGE""" & vbCrLf & _
           "    ElseIf futuresPct > 0.5 Then" & vbCrLf & _
           "        If topixMode = ""OK"" Then topixMode = ""MID""" & vbCrLf & _
           "    End If"

    If ReplaceInModule(modSell, oldS, newS) Then
        log = log & "OK: Sell futures filter added" & vbCrLf
        cnt = cnt + 1
    Else
        log = log & "SKIP: Sell - already applied or not found" & vbCrLf
    End If

    MsgBox "Futures Filter Result:" & vbCrLf & vbCrLf & log & vbCrLf & _
           "Applied: " & cnt & " / 2" & vbCrLf & vbCrLf & _
           "** Before using, set KanriSheet D6 = RssIndexMarket(N225.FUT01.OS, prev%)" & vbCrLf & _
           "Delete this module after applying.", _
           vbInformation, "FuturesFilter v1"
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
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
