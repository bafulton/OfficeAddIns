Attribute VB_Name = "FunctionReplace"
'Blackman & Sloop Excel Add-In, v1.2 (5/15/14)

Sub RunReplaceFunctions(control As IRibbonControl)
    On Error Resume Next

    repFormulas = Array( _
        "CLIENTNAME(", "CLIENTNAME2(", "CLIENTID(", "CLIENTADDRESS1(", "CLIENTADDRESS2(", "CLIENTCITY(", _
        "CLIENTSTATE(", "CLIENTZIP(", "CLIENTCOUNTRY(", "CLIENTPHONE(", "CLIENTFAX(", "CLIENTURL(", _
        "PRIMARYEMAIL(", "SECONDARYEMAIL(", "CLIENTTYPE(", "CLIENTINDUSTRY(", "CLIENTFEIN(", "CLIENTSTATEID(", _
        "FIRMNAME(", "FIRMADDRESS1(", "FIRMADDRESS2(", "FIRMCITY(", "FIRMSTATE(", "FIRMZIP(", "FIRMCOUNTRY(", _
        "FIRMPHONE(", "FIRMFAX(", "FIRMURL(", "CY(", "PY(", "CYBDATE(", "CYEDATE(", "CPBDATE(", "CPEDATE(", _
        "PYEDATE(", "PPBDATE(", "PPEDATE(", "PERIODSQ(", "PJNAME(", "BINDERID(", "BINDERDESC(", "BINDERDELIVDT", _
        "BINDERTYPE(", "BINDERCHRGCODE(", "BINDERLEAD(", "BINDERDATEOFREPORT(", "BINDERREPORTRELEASEDATE(", _
        "WPNAME(", "WPINDEX(", "ADIFF(", "AORAND(", "APDIFF(", "DDIFF(", "PDIFF(", "XFOOT(", "TBLINK(")

    'limit the selection by the end row & column
    If WorksheetFunction.CountA(Cells) > 0 Then
        'search for any entry, by searching backwards by rows
        LastR = Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        'search for any entry, by searching backwards by columns
        lastC = Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    End If
    Dim sel As Range, max As Range, inter As Range
    Set sel = Selection
    Set max = Range(Cells(1, 1), Cells(LastR, lastC))
    Set inter = Intersect(sel, max)
    If inter.Cells.count > 100 Then
        msg = MsgBox("This may take some time. Continue?", vbYesNo, "Warning")
        If msg = vbNo Then Exit Sub
    End If

    'loop through the selection and parse formulas
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim re As New RegExp
    Dim matches As MatchCollection
    Dim repLevel, repStart, repFormula As String
    With re
      .Pattern = "([a-z]+\()|(\))|([""].+[""])"
      .Global = True
      .IgnoreCase = True
    End With
    failure = False: failures = 0
    i = 1
    For Each c In inter.Cells
        failure = False
        Application.StatusBar = "Replacing formula " & i & " of " & inter.Cells.count
        If c.Formula <> "" Then
            'check if the cell contains formulas
            If re.Test(c.Formula) Then
                'get the collection of functions in the cell
                Set matches = re.Execute(c.Formula)
                repLevel = -999: repStart = 2: repFormula = "="

                For Each m In matches
                    If m = ")" Then
                        repLevel = repLevel - 1

                        If repLevel = 0 Then
                            'copy nonstandard formula into a blank cell
                            Cells(LastR + 1, lastC + 1) = "=" & Mid(c.Formula, repStart, m.FirstIndex + m.Length - repStart + 1)
                            repFormula = repFormula & Cells(LastR + 1, lastC + 1)
                            Cells(LastR + 1, lastC + 1).Delete
                            repStart = m.FirstIndex + m.Length + 1

                            repLevel = -999
                        End If
                    ElseIf left(m, 1) <> """" Then
                        repLevel = repLevel + 1

                        If (UBound(Filter(repFormulas, UCase(m))) > -1) Then
                            'start of nonstandard formula
                            If repLevel < 0 Then repLevel = 1
                        End If
                    End If
                    
                    If repLevel < 0 Then
                        'update the replacement formula
                        repFormula = repFormula & Mid(c.Formula, repStart, m.FirstIndex + m.Length - repStart + 1)
                        repStart = m.FirstIndex + m.Length + 1
                    End If
                Next

                'append any text remaining after the last close paren
                MsgBox Len(c.Formula)
                MsgBox InStrRev(c.Formula, ")")
                If InStrRev(c.Formula, ")") <> Len(c.Formula) Then
                    repFormula = repFormula & right(c.Formula, Len(c.Formula) - InStrRev(c.Formula, ")"))
                End If

                'check for any formula errors
                Cells(LastR + 1, lastC + 1) = repFormula
                If IsError(Cells(LastR + 1, lastC + 1)) Then
                    repFormula = right(repFormula, Len(repFormula) - 1)
                    Cells(LastR + 1, lastC + 1) = repFormula
                    If IsError(Cells(LastR + 1, lastC + 1)) Then failure = True: failures = failures + 1
                End If
                Cells(LastR + 1, lastC + 1).Delete

                'copy over the new formula
                If Not failure Then Range(c.Address) = repFormula
            End If
        End If

        i = i + 1
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

    If failures > 0 Then MsgBox failures & " formulas could not be replaced", vbCritical, "Formula Error"
End Sub
