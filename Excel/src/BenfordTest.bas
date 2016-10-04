Attribute VB_Name = "BenfordTest"
'Blackman & Sloop Excel Add-In, v1.2 (5/15/14)

Sub RunBenfordTest(control As IRibbonControl)
    On Error Resume Next

    Dim data As Range
    Set data = Selection
    Dim src As String
    src = "'" + ActiveSheet.name + "'"

    'create a new worksheet
    Dim Sh As Worksheet, flg As Boolean
    For Each Sh In Worksheets
        If Sh.name = "Benford Test" Then flg = True: Exit For
    Next
    If flg = True Then
        msg = MsgBox("Replace current Benford Test tab?", vbYesNo, "Replace Tab")
        If msg = vbYes Then
            Application.DisplayAlerts = False
            Sheets("Benford Test").Delete
            Application.DisplayAlerts = True
        Else: Exit Sub
        End If
    End If

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Sheets.Add.name = "Benford Test"
    Sheets("Benford Test").Select

    'set up the new sheet
    Cells.Interior.Color = RGB(216, 216, 216)
    Range("A1:D1").Merge
    Range("A1:D1").Font.Bold = True
    Range("A1:D1").HorizontalAlignment = xlCenter
    Range("A1:D1") = "Data"
    Range("F1:H1").Merge
    Range("F1:H1").Font.Bold = True
    Range("F1:H1") = "First Digit"
    Range("J1:L1").Merge
    Range("J1:L1").Font.Bold = True
    Range("J1:L1") = "Second Digit"
    Range("N1:P1").Merge
    Range("N1:P1").Font.Bold = True
    Range("N1:P1") = "First Two Digits"
    Range("A2:P2") = Array("Population", "1st Digit", "2nd Digit", "1st 2 Digits", _
        "", "Digit", "Expected", "Observed", _
        "", "Digit", "Expected", "Observed", _
        "", "Digit", "Expected", "Observed")
    Range("A1:P2").Interior.Color = RGB(216, 216, 216)
    Range("A2:D2, F2:H2, J2:L2, N2:P2").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Columns("E").ColumnWidth = 1
    Columns("I").ColumnWidth = 1
    Columns("M").ColumnWidth = 1
    Columns("Q").ColumnWidth = 1
    Range("F1:Q1").EntireColumn.hidden = True

    'set up data section
    Dim i As Integer
    i = 3
    count = 0: progress = 0: totalCount = WorksheetFunction.CountA(data)
    Application.StatusBar = progress & "% complete...."
    For Each d In data
        If d.Value <> "" And Val(d.Value) <> 0 Then
            count = count + 1
            If count / totalCount * 100 > progress + 10 Then
                progress = progress + 10
                Application.StatusBar = progress & "% complete...."
            End If

            Cells(i, 1) = Val(d.Value)
            Cells(i, 1).Hyperlinks.Add Anchor:=Cells(i, 1), Address:="", SubAddress:=src + "!" + d.Address
            i = i + 1
        End If
    Next
    Application.StatusBar = "Finalizing..."

    Range("A3:D" & i - 1).Interior.Color = xlNone
    Range("A3:A" & i - 1).name = "Data"
    Range("B3:B" & i - 1) = "=IFERROR(VALUE(LEFT(D3,1)),"""")"
    Range("B3:B" & i - 1).name = "Digit1"
    Range("C3:C" & i - 1) = "=IFERROR(VALUE(RIGHT(D3,1)),"""")"
    Range("C3:C" & i - 1).name = "Digit2"
    Range("D3:D" & i - 1) = "=VALUE(LEFT(ABS(A3*10000),2))"
    Range("D3:D" & i - 1).name = "Digit1and2"
    'set the data as a table & sort
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A2:D" & i - 1), , xlYes).name = "BenfordData"
    ActiveSheet.ListObjects("BenfordData").TableStyle = wdTableFormatNone
    Columns("A").ColumnWidth = 12.5
    Columns("B").ColumnWidth = 10
    Columns("C").ColumnWidth = 10.5
    Columns("D").ColumnWidth = 12
    'set up first digit section
    i = 1
    For Each c In Range("F3:F11").Cells
        c.Value = i
        i = i + 1
    Next
    Range("G3:G11") = "=LOG10(1+1/F3)*(COUNT(Digit1))"
    Range("H3:H11") = "=COUNTIF(Digit1,F3)"
    'set up second digit section
    i = 0
    For Each c In Range("J3:J12").Cells
        c.Value = i
        i = i + 1
    Next
    Range("K3:K12") = "=COUNT(Digit2)*(LOG10(1+1/(10*1+J3))+LOG10(1+1/(10*2+J3))+" _
        & "LOG10(1+1/(10*3+J3))+LOG10(1+1/(10*4+J3))+LOG10(1+1/(10*5+J3))+LOG10(1+1/(10*6+J3))+" _
        & "LOG10(1+1/(10*7+J3))+LOG10(1+1/(10*8+J3))+LOG10(1+1/(10*9+J3)))"
    Range("L3:L12") = "=COUNTIF(Digit2,J3)"
    'set up first two digits section
    i = 10
    For Each c In Range("N3:N92").Cells
        c.Value = i
        i = i + 1
    Next
    Range("O3:O92") = "=LOG10(1+1/N3)*COUNT(Digit2)"
    Range("P3:P92") = "=COUNTIF(Digit1and2,N3)"

    'create the charts
    ActiveSheet.ChartObjects.Add left:=262, Top:=25, Width:=300, Height:=200
    ActiveSheet.ChartObjects(1).Activate
    With ActiveChart
        .PlotVisibleOnly = False
        .HasTitle = True
        .ChartTitle.Text = "1st Digit Test"
        .HasLegend = False
        .SeriesCollection.Add source:=Range("G2:H11")
        .SeriesCollection(1).ChartType = xlLine
        .SeriesCollection(1).Border.Color = RGB(192, 80, 77)
        .SeriesCollection(2).ChartType = xlColumnClustered
        .SeriesCollection(2).Interior.Color = RGB(79, 129, 189)
        .SeriesCollection(1).XValues = Range("F3:F11")
        .SeriesCollection(2).XValues = Range("F3:F11")
        .Axes(xlCategory).AxisBetweenCategories = False
    End With
    ActiveSheet.ChartObjects.Add left:=572, Top:=25, Width:=300, Height:=200
    ActiveSheet.ChartObjects(2).Activate
    With ActiveChart
        .PlotVisibleOnly = False
        .HasTitle = True
        .ChartTitle.Text = "2nd Digit Test"
        .HasLegend = False
        .SeriesCollection.Add source:=Range("K2:L12")
        .SeriesCollection(1).ChartType = xlLine
        .SeriesCollection(1).Border.Color = RGB(192, 80, 77)
        .SeriesCollection(2).ChartType = xlColumnClustered
        .SeriesCollection(2).Interior.Color = RGB(79, 129, 189)
        .SeriesCollection(1).XValues = Range("J3:J12")
        .SeriesCollection(2).XValues = Range("J3:J12")
        .Axes(xlCategory).AxisBetweenCategories = False
    End With
    ActiveSheet.ChartObjects.Add left:=262, Top:=235, Width:=610, Height:=300
    ActiveSheet.ChartObjects(3).Activate
    With ActiveChart
        .PlotVisibleOnly = False
        .HasTitle = True
        .ChartTitle.Text = "First Two Digits Test"
        .HasLegend = False
        .SeriesCollection.Add source:=Range("O2:P92")
        .SeriesCollection(1).ChartType = xlLine
        .SeriesCollection(1).Border.Color = RGB(192, 80, 77)
        .SeriesCollection(2).ChartType = xlColumnClustered
        .SeriesCollection(2).Interior.Color = RGB(79, 129, 189)
        .SeriesCollection(1).XValues = Range("N3:N92")
        .SeriesCollection(2).XValues = Range("N3:N92")
        .Axes(xlCategory).AxisBetweenCategories = False
    End With

    'add the caveats section
    Range("R37:AD37").Merge
    Range("R38:AD38").Merge
    Range("R39:AD39").Merge
    Range("R40:AD40").Merge
    Range("R41:AD41").Merge
    Range("R42:AD42").Merge
    Range("R43:AD43").Merge
    Range("R44:AD44").Merge
    Range("R45:AD45").Merge
    Range("R47:AD47").Merge
    Range("R48:AD49").Merge
    Range("R37:AD37").Font.Bold = True
    Range("R37:AD37").Font.Size = 14
    Range("R37:AD37") = "Only use Benford's Law when..."
    Range("R38:AD38") = "    1. The population is large"
    Range("R39:AD39") = "           a. Minimum size is 100 items"
    Range("R40:AD40") = "           b. Ideal size is 500+ items"
    Range("R41:AD41") = "    2. The population is natural"
    Range("R42:AD42") = "           a. No built-in maximums/minimums (eg, no testing of only invoices from $50-$500)"
    Range("R43:AD43") = "           b. Numbers can reoccur (eg, don't use check numbers--they can't reoccur)"
    Range("R44:AD44") = "           c. The sample is random (eg, not haphazard/judgmental selection; prefer use of the entire population)"
    Range("R45:AD45") = "    Note: The Benford test ignores zeroes."
    Range("R47:AD47").Font.Bold = True
    Range("R47:AD47").Font.Underline = True
    Range("R47:AD47") = "Paste this into your testing workpaper:"
    Range("R48:AD49").WrapText = True
    Range("R48:AD49") = "B&S tested the total population using Benford's Law to identify if any disbursement amounts occurred more frequently " _
        & "than would naturally be expected. B&S used the results of that test to aid in the judgmental selection of [number] items for testing."

    Range("A1").EntireRow.hidden = True
    Range("A2").Activate

    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub
