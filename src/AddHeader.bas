Attribute VB_Name = "AddHEader"
'Blackman & Sloop Excel Add-In, v1.2 (5/15/14)

Sub AddAHeader(control As IRibbonControl)
    On Error Resume Next

    'ask if include materiality calculations
    Dim cols, offset As Integer
    offset = 0
    msg = MsgBox("Include materiality calculations?", vbYesNo, "Materiality")
    If msg = vbYes Then offset = 4
    cols = Selection.Columns.count
    If offset <> 0 Then
        If cols < 7 Then cols = 7
    Else
        If cols < 3 Then cols = 7
    End If

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    'insert the header rows & set to default
    Range("A1:A9").EntireRow.Insert
    Range(Cells(1, 1), Cells(9, cols)).ClearFormats
    Range(Cells(1, 1), Cells(9, cols)).Font.Name = Range("A10").Font.Name
    Range(Cells(1, 1), Cells(9, cols)).Font.Size = Range("A10").Font.Size

'---NAME/INDEX/DATE------------------------------------------------------------------

    'date
    Range("A1") = "=pjname()"
    Range(Cells(1, 1), Cells(1, cols - offset)).Merge

    'wp name & index
    Range("A2") = "=wpname()&"" (""&wpindex()&"")"""
    Range(Cells(2, 1), Cells(2, cols - offset)).Merge

    'client name
    Range("A3") = "=cyedate()"
    Range(Cells(3, 1), Cells(3, cols - offset)).Merge
    Range(Cells(3, 1), Cells(3, cols - offset)).NumberFormat = "mmmm dd, yyyy"

    'format all rows
    Range(Cells(1, 1), Cells(3, cols - offset)).HorizontalAlignment = xlCenter
    Range(Cells(1, 1), Cells(3, cols - offset)).Font.Bold = True
    Range(Cells(1, 1), Cells(3, cols - offset)).Interior.Color = RGB(216, 216, 216)
    Range(Cells(3, 1), Cells(3, cols)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range(Cells(1, cols - 4), Cells(3, cols - 4)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Range(Cells(1, cols), Cells(3, cols)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Range(Cells(1, 1), Cells(3, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    Range(Cells(1, 1), Cells(1, cols)).Borders(xlEdgeTop).LineStyle = xlContinuous

'---PURPOSE/PROCEDURES/CONCLUSION----------------------------------------------------

    'purpose
    Range("A5") = "Purpose:"
    Range(Cells(5, 2), Cells(5, cols)).Merge

    'procedures
    Range("A6") = "Procedures:"
    Range(Cells(6, 2), Cells(7, cols)).Merge

    'conclusion
    Range("A8") = "Conclusion:"
    Range(Cells(8, 2), Cells(8, cols)).Merge

    'format all rows
    Range("A5:A8").HorizontalAlignment = xlRight
    Range("A5:A8").Font.Bold = True
    Range(Cells(5, 2), Cells(8, cols)).HorizontalAlignment = xlLeft
    Range(Cells(5, 2), Cells(8, cols)).VerticalAlignment = xlTop
    Range(Cells(5, 2), Cells(8, cols)).WrapText = True

'---MATERIALITY----------------------------------------------------------------------

    If offset <> 0 Then
        'materiality, performance, & trivial
        Cells(1, cols - 3) = "Materiality:"
        Cells(1, cols - 2) = 0
        Cells(1, cols - 2).Name = "Materiality"
        Cells(2, cols - 3) = "Performance:"
        Cells(2, cols - 2) = "=ROUNDDOWN(" & Cells(1, cols - 2).Address & "*0.75,-LEN(INT(" & Cells(1, cols - 2).Address & "*0.75))+2)"
        Cells(2, cols - 2).Name = "Performance"
        Cells(3, cols - 3) = "Trivial:"
        Cells(3, cols - 2) = "=ROUNDDOWN(" & Cells(1, cols - 2).Address & "*0.05,-LEN(INT(" & Cells(1, cols - 2).Address & "*0.05))+2)"
        Cells(3, cols - 2).Name = "Trivial"
        'formatting
        Range(Cells(1, cols - 3), Cells(3, cols - 3)).HorizontalAlignment = xlRight
        Range(Cells(1, cols - 2), Cells(3, cols - 2)).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* "" - ""_);_(@_)"
        Cells(1, cols - 2).Interior.Color = RGB(146, 208, 80)
    
        'risk level, threshold %, & threshold $
        Cells(1, cols - 1) = "Assessed risk:"
        Cells(1, cols) = "High"
        Cells(2, cols - 1) = "Scope %:"
        Cells(2, cols) = "=IF(" & Cells(1, cols).Address & "=""Low"",0.2,IF(" & Cells(1, cols).Address & "=""Moderate"",0.15,0.1))"
        Cells(3, cols - 1) = "Scope $:"
        Cells(3, cols) = "=" & Cells(2, cols - 2).Address & "*" & Cells(2, cols).Address
        Cells(3, cols).Name = "Threshold"
        'formatting
        Range(Cells(1, cols - 1), Cells(3, cols - 1)).HorizontalAlignment = xlRight
        Range(Cells(1, cols), Cells(2, cols)).HorizontalAlignment = xlCenter
        Cells(1, cols).Interior.Color = RGB(146, 208, 80)
        Cells(1, cols).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Low,Moderate,High"
        Cells(2, cols).NumberFormat = "0%"
        Cells(3, cols).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* "" - ""_);_(@_)"
    
        'set selected cell
        Range(Cells(1, cols - 2), Cells(1, cols - 2)).Activate
    End If

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
