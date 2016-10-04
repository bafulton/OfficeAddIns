Attribute VB_Name = "RemoveStyles"
'Blackman & Sloop Excel Add-In, v1.2 (5/15/14)

Sub RemoveTheStyles(control As IRibbonControl)
    On Error Resume Next

    Dim s As Style, i As Long, c As Long

    If ActiveWorkbook.MultiUserEditing Then
        If MsgBox("You cannot remove Styles in a Shared workbook." & vbCr & vbCr & _
                  "Do you want to unshare the workbook?", vbYesNo + vbInformation) = vbYes Then
            ActiveWorkbook.ExclusiveAccess
            If Err.Description = "Application-defined or object-defined error" Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If

    c = ActiveWorkbook.Styles.count
    Application.ScreenUpdating = False
    For i = c To 1 Step -1
        If i Mod 600 = 0 Then DoEvents
        Set s = ActiveWorkbook.Styles(i)
        Application.StatusBar = "Deleting " & c - i + 1 & " of " & c & " " & s.name
        If Not s.BuiltIn Then
            s.Delete
            If Err.Description = "You cannot use this command on a protected sheet. To use this command, you must first unprotect the sheet (Review tab, Changes group, Unprotect Sheet button). You may be prompted for a password." Then
                MsgBox Err.Description & vbCr & "You have to unprotect all of the sheets in the workbook to remove styles.", vbExclamation, "Remove Styles AddIn"
                Exit For
            End If
        End If
    Next
    
    Application.ScreenUpdating = True
    Application.StatusBar = False

End Sub
