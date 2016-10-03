Attribute VB_Name = "QBTBClean"
'Blackman & Sloop Excel Add-In, v1.4 (10/2/16)

Sub CleanTheTB(control As IRibbonControl)
    On Error Resume Next

    msg = MsgBox("Exclude $0 balances?", vbYesNo, "$0 Balances")

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    'Find the debit/credit cols and starting row
    found = 0
    Dim startRow, debitCol, creditCol As Integer
    For Each cell In Range("A1:K10").Cells
        If cell.Value = "Debit" Then
            found = found + 1
            startRow = cell.Row + 1
            debitCol = cell.Column
        End If
        If cell.Value = "Credit" Then
            MsgBox cell.Col
            found = found + 1
            creditCol = cell.Column
        End If

        If found = 2 Then
            Exit For
        End If
    Next cell

    'Find the column with the account number & name
    Dim accountCol As Integer
    For Each cell In Range(Cells(startRow, 1), Cells(startRow, 20)).Cells
        If cell.Value <> "" Then
            'MsgBox cell.Address
            accountCol = cell.Column
            Exit For
        End If
    Next cell

    'Find the last data row
    endRow = Cells(Rows.count, accountCol).End(xlUp).Row

    'Determine if the TB is from QB Online
    QBOnline = True
    'Search the first ten rows of accounts for a '·' character
    For Each cell In Range(Cells(startRow, accountCol), Cells(startRow + 10, accountCol))
        If InStr(Cells(startRow, accountCol).Value, "·") <> 0 Then
            QBOnline = False
            Exit For
        End If
    Next cell

    'MsgBox "startRow: " & startRow & "endRow: " & endRow & ", accountCol: " & accountCol & ", debitCol: " & debitCol & ", creditCol: " & creditCol & ", QBOnline: " & QBOnline

    For i = startRow To endRow
        'If debit or credit not $0
            'Split the account number and description
            'Add the debit/credit
    Next i


    ' If regular QB, then just split everything after the last colon on the weird dot thing
    ' If QB Online, then the account is at the front, and the name is everything after the last colon

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Private Function SplitAccount(TempAccount, WhatToReturn)
    If InStr(1, TempAccount, ":") > 0 Then
        MyPos = InStr(1, TempAccount, ":")
        MasterAccount = Mid(TempAccount, 1, (MyPos - 1))
        SubAccount = Mid(TempAccount, (MyPos + 1))

        If InStr(1, MasterAccount, " · ") > 0 Then
            MyPos = InStr(1, MasterAccount, " · ")
            MasterAccountNumber = Mid(MasterAccount, 1, (MyPos - 1))
            MasterAccountName = Mid(MasterAccount, (MyPos + 3))
        Else
            MasterAccountName = MasterAccount
        End If

        If InStr(1, SubAccount, " · ") > 0 Then
            MyPos = InStr(1, SubAccount, " · ")
            SubAccountNumber = Mid(SubAccount, 1, (MyPos - 1))
            subaccountname = Mid(SubAccount, (MyPos + 3))
        Else
            subaccountname = SubAccount
            SubAccountNumber = ""
        End If

        ProcessedName = MasterAccountName & ":" & subaccountname
        ProcessedNumber = SubAccountNumber
    Else
        SubAccount = TempAccount
        If InStr(1, SubAccount, " · ") > 0 Then
            MyPos = InStr(1, SubAccount, " · ")
            SubAccountNumber = Mid(SubAccount, 1, (MyPos - 1))
            subaccountname = Mid(SubAccount, (MyPos + 3))
        Else
            subaccountname = SubAccount
            SubAccountNumber = ""
        End If

        ProcessedName = subaccountname
        ProcessedNumber = SubAccountNumber
    End If

    If WhatToReturn = "Name" Then
        SplitAccount = ProcessedName
    ElseIf WhatToReturn = "Num" Then
        SplitAccount = ProcessedNumber
    End If

End Function
