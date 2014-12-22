Attribute VB_Name = "AddHyperlinks"
'Blackman & Sloop Excel Add-In, v1.2 (5/15/14)

Dim hyperlinking As Boolean
Dim srcSheet, srcCell As String

Sub SetSource(control As IRibbonControl)
    On Error Resume Next

    srcSheet = ActiveSheet.Name
    srcCell = ActiveCell.Address(False, False)

    'Activate hyperlinking
    hyperlinking = True
End Sub

Sub SetDestination()
    On Error Resume Next

    If hyperlinking Then
        Dim destSheet, destCell As String
        destSheet = ActiveSheet.Name
        destCell = ActiveCell.Address(False, False)
        'MsgBox srcSheet + "!" + srcCell + ", " + destSheet + "!" + destCell

        'Set the forward & reverse hyperlinks
        Sheets(srcSheet).Range(srcCell).Hyperlinks.Add _
            Anchor:=Sheets(srcSheet).Range(srcCell), _
            Address:="", _
            SubAddress:="'" + destSheet + "'!" + destCell
        Sheets(destSheet).Range(destCell).Hyperlinks.Add _
            Anchor:=Sheets(destSheet).Range(destCell), _
            Address:="", _
            SubAddress:="'" + srcSheet + "'!" + srcCell

        'Turn off hyperlinking
        hyperlinking = False
    End If
End Sub

Sub ClearLinks(control As IRibbonControl)
    On Error Resume Next

    'Delete all hyperlinks in selection (forward links only)
    Selection.Hyperlinks.Delete
End Sub

