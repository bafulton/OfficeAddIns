Attribute VB_Name = "GapAnalysis"
'Blackman & Sloop Excel Add-In, v1.2 (5/15/14)

Sub RunGapAnalysis(control As IRibbonControl)
    On Error Resume Next

    Dim data As Range
    Set data = Selection
    Dim src As String
    src = "'" + ActiveSheet.name + "'"

    'Create a new worksheet
    Dim Sh As Worksheet, flg As Boolean
    For Each Sh In Worksheets
        If Sh.name = "Gap Analysis" Then flg = True: Exit For
    Next
    If flg = True Then
        msg = MsgBox("Replace current Gap Analysis tab?", vbYesNo, "Replace Tab")
        If msg = vbYes Then
            Application.DisplayAlerts = False
            Sheets("Gap Analysis").Delete
            Application.DisplayAlerts = True
        Else: Exit Sub
        End If
    End If

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Sheets.Add.name = "Gap Analysis"
    Sheets("Gap Analysis").Select
    Range("A1") = "Values"
    Range("A1").HorizontalAlignment = xlLeft
    Range("A1").Font.Bold = True
    Range("A1").Borders(xlEdgeBottom).LineStyle = xlContinuous

    'Copy over the data
    Dim i As Integer
    i = 2
    count = 0: progress = 0: totalCount = WorksheetFunction.CountA(data)
    Application.StatusBar = progress & "% complete...."
    For Each d In data
        If d.Value <> "" And Val(d.Value) <> 0 Then
            count = count + 1
            If count / totalCount * 100 > progress + 10 Then
                progress = progress + 10
                Application.StatusBar = progress & "% complete...."
            End If

            Cells(i, 1) = "=Int(" & Val(d.Value) & ")"
            Cells(i, 1).Hyperlinks.Add Anchor:=Cells(i, 1), Address:="", SubAddress:=src + "!" + d.Address
            i = i + 1
        End If
    Next
    Application.StatusBar = "Finalizing..."

    'Set the data as a table & sort
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:A" & i - 1), , xlYes).name = "GapData"
    With ActiveSheet.ListObjects("GapData")
        .TableStyle = wdTableFormatNone
        .Sort.SortFields.Add Key:=Range("GapData[Values]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Sort.Apply
    End With

    'Add highlights
    Dim previous As Integer
    previous = Range("A2").Value - 1
    For Each c In Range("GapData")
        If c.Value <> previous + 1 Then c.Interior.Color = rgbYellow
        previous = c.Value
    Next

    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub
