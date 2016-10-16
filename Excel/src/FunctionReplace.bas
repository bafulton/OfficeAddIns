Attribute VB_Name = "FunctionReplace"
'Blackman & Sloop Excel Add-In, v1.4 (10/16/16)

Sub RunReplaceFunctions(control As IRibbonControl)
    On Error Resume Next

    repFormulas = Array( _
        "CLIENTNAME(", "CLIENTNAME2(", "CLIENTID(", "CLIENTADDRESS1(", "CLIENTADDRESS2(", "CLIENTCITY(", _
        "CLIENTSTATE(", "CLIENTZIP(", "CLIENTCOUNTRY(", "CLIENTPHONE(", "CLIENTFAX(", "CLIENTURL(", _
        "PRIMARYEMAIL(", "SECONDARYEMAIL(", "CLIENTTYPE(", "CLIENTINDUSTRY(", "CLIENTFEIN(", "CLIENTSTATEID(", _
        "FIRMNAME(", "FIRMADDRESS1(", "FIRMADDRESS2(", "FIRMCITY(", "FIRMSTATE(", "FIRMZIP(", "FIRMCOUNTRY(", _
        "FIRMPHONE(", "FIRMFAX(", "FIRMURL(", "CY(", "PY(", "CYBDATE(", "CYEDATE(", "CPBDATE(", "CPEDATE(", _
        "PYEDATE(", "PPBDATE(", "PPEDATE(", "PERIODSQ(", "PJNAME(", "BINDERID(", "BINDERDESC(", "BINDERDELIVDT(", _
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
    
    'warn if selection is large
    If inter.Cells.count > 100 Then
        msg = MsgBox("This may take some time. Continue?", vbYesNo, "Warning")
        If msg = vbNo Then Exit Sub
    End If

    'disable updating (saves time)
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'loop through the selection and parse formulas
    curFormula = c.Formula
    For Each c In inter.Cells
        If left(curFormula, 1) = "=" Then
            For Each f In repFormulas
                Do Until InStr(UCase(curFormula), f) = 0
                    curLevel = 1
                    For i = InStr(UCase(curFormula), f) + Len(f) To Len(curFormula)
                        If Mid(curFormula, i, 1) = "(" Then
                            curLevel = curLevel + 1
                        ElseIf Mid(curFormula, i, 1) = ")" Then
                            curLevel = curLevel - 1
                        End If
                        If curLevel = 0 Then
                            MsgBox "happy"
                        End If
                    Next i
                Loop
            Next f
        End If
    Next

    're-enable updating
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

