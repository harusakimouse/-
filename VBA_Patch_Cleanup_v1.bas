Attribute VB_Name = "Mod_Patch_Cleanup_V1"

Public Sub CleanupFuturesFilter()
    Dim vbProj As Object
    Dim modBuy  As Object
    Dim modSell As Object

    On Error GoTo ErrHandler
    Set vbProj = ThisWorkbook.VBProject

    ' Find modules by partial ASCII name match (avoids Japanese encoding issues)
    Dim comp As Object
    For Each comp In vbProj.VBComponents
        ' Buy module contains "v13" and not "Sell"
        If InStr(comp.name, "v13") > 0 And InStr(comp.name, "Sell") = 0 Then
            Set modBuy = comp.CodeModule
        End If
        ' Sell module is ASCII-only name
        If comp.name = "Mod_SellExtrac" Then
            Set modSell = comp.CodeModule
        End If
    Next comp

    If modBuy Is Nothing Then MsgBox "Buy module (v13) not found", vbCritical: Exit Sub
    If modSell Is Nothing Then MsgBox "Mod_SellExtrac not found", vbCritical: Exit Sub

    Dim log As String: log = ""

    ' --- BUY: remove duplicate futures block and update to Cells(5,8) ---
    Dim oldB As String, newB As String

    ' Try to remove duplicate (two blocks back to back)
    oldB = "    ' Futures filter: KanriSheet D6 = RssIndexMarket(N225.FUT01.OS, prev day %)" & vbCrLf & _
           "    Dim futuresPct As Double" & vbCrLf & _
           "    futuresPct = SafeNum(mws.Cells(6, 4).Value, 0)" & vbCrLf & _
           "    If futuresPct < -1.5 Then" & vbCrLf & _
           "        topixMode = ""LARGE""" & vbCrLf & _
           "    ElseIf futuresPct < -0.5 Then" & vbCrLf & _
           "        If topixMode = ""OK"" Then topixMode = ""MID""" & vbCrLf & _
           "    End If" & vbCrLf & _
           "    ' Futures filter: KanriSheet D6 = RssIndexMarket(N225.FUT01.OS, prev day %)" & vbCrLf & _
           "    Dim futuresPct As Double" & vbCrLf & _
           "    futuresPct = SafeNum(mws.Cells(6, 4).Value, 0)" & vbCrLf & _
           "    If futuresPct < -1.5 Then" & vbCrLf & _
           "        topixMode = ""LARGE""" & vbCrLf & _
           "    ElseIf futuresPct < -0.5 Then" & vbCrLf & _
           "        If topixMode = ""OK"" Then topixMode = ""MID""" & vbCrLf & _
           "    End If"
    newB = "    ' Futures filter: H5 = RssIndexMarket(N225.FUT01.OS, prev day %)" & vbCrLf & _
           "    Dim futuresPct As Double" & vbCrLf & _
           "    futuresPct = SafeNum(mws.Cells(5, 8).Value, 0)" & vbCrLf & _
           "    If futuresPct < -1.5 Then" & vbCrLf & _
           "        topixMode = ""LARGE""" & vbCrLf & _
           "    ElseIf futuresPct < -0.5 Then" & vbCrLf & _
           "        If topixMode = ""OK"" Then topixMode = ""MID""" & vbCrLf & _
           "    End If"

    If ReplaceInModule(modBuy, oldB, newB) Then
        log = log & "OK: Buy - duplicate removed, cell -> H5" & vbCrLf
    Else
        ' Single block: just fix cell reference
        Dim oldB2 As String, newB2 As String
        oldB2 = "    futuresPct = SafeNum(mws.Cells(6, 4).Value, 0)"
        newB2 = "    futuresPct = SafeNum(mws.Cells(5, 8).Value, 0)"
        If ReplaceInModule(modBuy, oldB2, newB2) Then
            log = log & "OK: Buy - cell fixed to H5" & vbCrLf
        Else
            log = log & "SKIP: Buy - no change needed" & vbCrLf
        End If
    End If

    ' --- SELL: fix cell reference ---
    Dim oldS As String, newS As String
    oldS = "    futuresPct = S_SafeNum(mws.Cells(6, 4).Value, 0)"
    newS = "    futuresPct = S_SafeNum(mws.Cells(5, 8).Value, 0)"
    If ReplaceInModule(modSell, oldS, newS) Then
        log = log & "OK: Sell - cell fixed to H5" & vbCrLf
    Else
        log = log & "SKIP: Sell - no change needed" & vbCrLf
    End If

    MsgBox "Cleanup Result:" & vbCrLf & vbCrLf & log & vbCrLf & _
           "Done. Delete this module.", vbInformation, "Cleanup v1"
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
