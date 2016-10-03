Attribute VB_Name = "Automerge"
'Blackman & Sloop Excel Add-In, v1.4 (10/2/16)

Sub RunAutomerge(control As IRibbonControl)
    On Error Resume Next

    Selection.VerticalAlignment = xlTop
    Selection.HorizontalAlignment = xlLeft

    If ActiveCell.MergeCells = False Then
        'Merge the selection
        Selection.Merge
        Selection.WrapText = True
    Else
        'Unmerge the selection
        Selection.UnMerge
        Selection.WrapText = False
    End If

End Sub
