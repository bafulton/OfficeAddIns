Attribute VB_Name = "Comparison"
'Blackman & Sloop Excel Add-In, v1.2 (5/15/14)

Sub RunComparison(control As IRibbonControl)
    On Error Resume Next

    'get the two columns of data
    Dim left, right As Range
    Dim error As Boolean
    error = False
    If Selection.Areas.count = 1 Then
        If Selection.Columns.count <> 2 Then
            error = True
        Else
            Set left = Selection.Columns(1)
            Set right = Selection.Columns(2)
        End If
    ElseIf Selection.Areas.count = 2 Then
        If Selection.Areas(1).Columns.count <> 1 Or Selection.Areas(2).Columns.count <> 1 Then
            error = True
        Else
            Set left = Selection.Areas(1)
            Set right = Selection.Areas(2)
        End If
    Else
        error = True
    End If
    If error Then
        MsgBox "You must select only two columns", vbCritical, "Column Error"
        Exit Sub
    End If

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    'identify unique values
    Dim count As Integer
    count = 0
    Dim r As Range
    'i = 1: totalCount = left.Cells.count + right.Cells.count
    For Each c In left.Cells
        'Application.StatusBar = "Comparing value " & i & " of " & totalCount
        'i = i + 1

        If c.Value <> "" Then
            Set r = right.Find(What:=c.Value, LookIn:=xlValues, LookAt:=xlWhole)
            If r Is Nothing Then
                c.Interior.Color = rgbYellow
                count = count + 1
            End If
        End If
    Next
    For Each c In right.Cells
        'Application.StatusBar = "Comparing value " & i & " of " & totalCount
        'i = i + 1

        If c.Value <> "" Then
            Set r = left.Find(What:=c.Value, LookIn:=xlValues, LookAt:=xlWhole)
            If r Is Nothing Then
                c.Interior.Color = rgbYellow
                count = count + 1
            End If
        End If
    Next

    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox count & " unique values identified.", Title:="Compare Complete"

End Sub
