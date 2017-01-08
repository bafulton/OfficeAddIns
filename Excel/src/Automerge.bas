Attribute VB_Name = "Automerge"
'Blackman & Sloop Excel Add-In, v1.4 (10/2/16)

Sub RunAutomerge(control As IRibbonControl)
    On Error Resume Next

    Selection.VerticalAlignment = xlTop
    Selection.HorizontalAlignment = xlLeft

    If Selection.MergeCells = True Then
        'Unmerge the selection
        Selection.UnMerge
        Selection.WrapText = False
    Else
        'Merge the selection
        Selection.Merge
        Selection.WrapText = True
    End If

End Sub
