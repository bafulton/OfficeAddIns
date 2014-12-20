Attribute VB_Name = "AddHEader"
'Blackman & Sloop Excel Add-In, v1.1 (7/8/13)

Sub AddAHeader(control As IRibbonControl)
    On Error Resume Next

    'blank row
    Cells(1, 1).EntireRow.Insert

    'conclusion
    Cells(1, 1).EntireRow.Insert
    Cells(1, 1) = "Conclusion:"
    Cells(1, 1).HorizontalAlignment = xlRight
    Cells(1, 1).Font.Bold = True
    Range(Cells(1, 2), Cells(1, 8)).Merge
    Range(Cells(1, 2), Cells(1, 8)).HorizontalAlignment = xlLeft
    Range(Cells(1, 2), Cells(1, 8)).VerticalAlignment = xlTop
    Range(Cells(1, 2), Cells(1, 8)).WrapText = True
    Range(Cells(1, 1), Cells(1, 8)).NumberFormat = "General"

    'procedures
    Cells(1, 1).EntireRow.Insert
    Cells(1, 1).EntireRow.Insert
    Cells(1, 1) = "Procedures:"
    Cells(1, 1).HorizontalAlignment = xlRight
    Cells(1, 1).Font.Bold = True
    Range(Cells(1, 2), Cells(2, 8)).Merge
    Range(Cells(1, 2), Cells(2, 8)).HorizontalAlignment = xlLeft
    Range(Cells(1, 2), Cells(2, 8)).VerticalAlignment = xlTop
    Range(Cells(1, 2), Cells(1, 8)).WrapText = True
    Range(Cells(1, 1), Cells(1, 8)).NumberFormat = "General"

    'purpose
    Cells(1, 1).EntireRow.Insert
    Cells(1, 1) = "Purpose:"
    Cells(1, 1).HorizontalAlignment = xlRight
    Cells(1, 1).Font.Bold = True
    Range(Cells(1, 2), Cells(1, 8)).Merge
    Range(Cells(1, 2), Cells(1, 8)).HorizontalAlignment = xlLeft
    Range(Cells(1, 2), Cells(1, 8)).VerticalAlignment = xlTop
    Range(Cells(1, 2), Cells(1, 8)).WrapText = True
    Range(Cells(1, 1), Cells(1, 8)).NumberFormat = "General"

    'blank row
    Cells(1, 1).EntireRow.Insert

    'date
    Cells(1, 1).EntireRow.Insert
    Cells(1, 1) = "=cyedate()"
    Range(Cells(1, 1), Cells(1, 8)).Merge
    Range(Cells(1, 1), Cells(1, 8)).NumberFormat = "mmmm dd, yyyy"

    'wp name & index
    Cells(1, 1).EntireRow.Insert
    Cells(1, 1) = "=wpname()&"" (""&wpindex()&"")"""
    Range(Cells(1, 1), Cells(1, 8)).Merge
    Range(Cells(1, 1), Cells(1, 8)).NumberFormat = "General"

    'client name
    Cells(1, 1).EntireRow.Insert
    Cells(1, 1) = "=pjname()"
    Range(Cells(1, 1), Cells(1, 8)).Merge
    Range(Cells(1, 1), Cells(1, 8)).Font.Bold = True
    Range(Cells(1, 1), Cells(1, 8)).NumberFormat = "General"

    'left align the top three rows
    Range(Cells(1, 1), Cells(3, 8)).HorizontalAlignment = xlLeft
End Sub
