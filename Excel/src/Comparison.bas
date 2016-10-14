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
    For Each c In left.Cells
        If c.Value <> "" Then
            Set r = right.Find(What:=c.Value, LookIn:=xlFormulas, LookAt:=xlWhole)
            If r Is Nothing Then
                c.Interior.Color = rgbYellow
                count = count + 1
            End If
        End If
    Next c
    For Each c In right.Cells
        If c.Value <> "" Then
            Set r = left.Find(What:=c.Value, LookIn:=xlFormulas, LookAt:=xlWhole)
            If r Is Nothing Then
                c.Interior.Color = rgbYellow
                count = count + 1
            End If
        End If
    Next c

    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox count & " unique values identified. See yellow highlights.", Title:="Compare Complete"

End Sub
