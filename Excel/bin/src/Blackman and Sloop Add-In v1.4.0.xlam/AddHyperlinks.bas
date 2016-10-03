Attribute VB_Name = "AddHyperlinks"
'Blackman & Sloop Excel Add-In, v1.4 (10/2/16)

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

        'Make the entire cell clickable--not just the text
        Range(src).WrapText = True
        Range(dest).WrapText = True

        'Format the links
        Range(src).Font.Bold = True
        Range(dest).Font.Bold = True
        Range(src).Font.Color = vbBlue
        Range(dest).Font.Color = vbBlue

        'Turn off hyperlinking
        hyperlinking = False
    End If

End Sub

Sub ClearLinks(control As IRibbonControl)
    On Error Resume Next

    For Each h In Selection.Hyperlinks
        'Delete the hyperlink names
        Range(h.SubAddress).Name.Delete
    Next h
    'Delete all hyperlinks in selection
    Selection.Hyperlinks.Delete

End Sub


