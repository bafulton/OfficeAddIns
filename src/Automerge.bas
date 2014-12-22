Attribute VB_Name = "Automerge"
'Blackman & Sloop Excel Add-In, v1.2 (5/15/14)

Sub RunAutomerge(control As IRibbonControl)
    On Error Resume Next

    'Selection.ClearFormats
    Selection.Merge
    Selection.VerticalAlignment = xlTop
    Selection.HorizontalAlignment = xlLeft
    Selection.WrapText = True
End Sub
