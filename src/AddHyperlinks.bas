Attribute VB_Name = "AddHyperlinks"
'Blackman & Sloop Excel Add-In, v1.3 (12/18/14)

Dim hyperlinking As Boolean
Dim src, dest As String

Sub SetSource(control As IRibbonControl)
    On Error Resume Next

    'determine the last tickmark
    nextTickmark = 1
    For Each n In ActiveWorkbook.Names
        If left(n.Name, 9) = "tickmark_" Then
            thisTickmark = Val(Mid(n.Name, 10, Len(n.Name) - 10))
            If thisTickmark >= nextTickmark Then
                nextTickmark = thisTickmark + 1
            End If
        End If
    Next n
    src = "tickmark_" & nextTickmark & "a"
    dest = "tickmark_" & nextTickmark & "b"

    'create the new named range
    ActiveCell.Name = src

    'Activate hyperlinking
    hyperlinking = True
End Sub

Sub SetDestination()
    On Error Resume Next

    If hyperlinking Then
        'create the new named range
        ActiveCell.Name = dest

        'Set the forward & reverse hyperlinks
        Range(src).Hyperlinks.Add _
            Anchor:=Range(src), _
            Address:="", _
            SubAddress:=dest
        Range(dest).Hyperlinks.Add _
            Anchor:=Range(dest), _
            Address:="", _
            SubAddress:=src

        'Turn off hyperlinking
        hyperlinking = False
    End If
End Sub

Sub ClearLinks(control As IRibbonControl)
    On Error Resume Next

    'Delete all hyperlinks in selection (forward links only)
    For Each h In Selection.Hyperlinks
        MsgBox h.SubAddress
        Range(h.SubAddress).Name.Delete
    Next h
    Selection.Hyperlinks.Delete
End Sub

