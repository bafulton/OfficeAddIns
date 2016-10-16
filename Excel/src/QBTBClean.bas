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
        If InStr(cell.Value, "·") <> 0 Then
            QBOnline = False
            Exit For
        End If
    Next cell

    'Debug message
    'MsgBox "startRow: " & startRow & "endRow: " & endRow & ", accountCol: " & accountCol & ", debitCol: " & debitCol & ", creditCol: " & creditCol & ", QBOnline: " & QBOnline

    'Insert the new columns
    Range(Cells(1, creditCol + 2).Address, Cells(1, creditCol + 4).Address).EntireColumn.Insert
    'Add & format the header
    Cells(startRow - 2, creditCol + 2).Value = "Account"
    Cells(startRow - 2, creditCol + 3).Value = "Name"
    Cells(startRow - 2, creditCol + 4).Value = "Balance"
    Range(Cells(startRow - 2, creditCol + 2), Cells(startRow - 2, creditCol + 4)).HorizontalAlignment = xlCenter
    Range(Cells(startRow - 2, creditCol + 2), Cells(startRow - 2, creditCol + 4)).Font.Bold = True
    Range(Cells(startRow - 2, creditCol + 2), Cells(startRow - 2, creditCol + 4)).Interior.Color = RGB(217, 217, 217)
    Range(Cells(startRow - 2, creditCol + 2), Cells(startRow - 2, creditCol + 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous

    curRow = startRow
    For i = startRow To endRow
        If (Cells(i, debitCol).Value > 0 Xor Cells(i, creditCol).Value > 0) Or _
            (msg = vbNo And (IsEmpty(Cells(i, debitCol).Value) Xor IsEmpty(Cells(i, creditCol).Value))) Then

            'Split and add the account number and description
            splits = Split(Cells(i, accountCol).Value, ":")
            rightmost = splits(UBound(splits))

            Dim account As Integer
            Dim name As String
            If QBOnline = True Then
                'Account number is at beginning, name is after last colon
                If IsNumeric(Split(splits(0))(0)) = True Then
                    account = Split(splits(0))(0)
                End If

                If UBound(splits) = 0 And IsNumeric(Split(splits(0))(0)) = True Then
                    name = right(rightmost, Len(rightmost) - InStr(rightmost, " "))
                Else
                    name = rightmost
                End If
            
            Else
                'Account number and name are after the last colon (split on the '·')
                temp = Split(rightmost, " · ")
                If IsNumeric(temp(0)) = True Then
                    account = temp(0)
                    name = temp(1)
                Else
                    name = temp(0)
                End If
            End If

            'Add the values
            If account <> 0 Then
                Cells(curRow, creditCol + 2) = account
            Else
                'Highlight missing account numbers
                Cells(curRow, creditCol + 2).Interior.ColorIndex = 27
            End If
            Cells(curRow, creditCol + 3) = name
            Cells(curRow, creditCol + 4) = Cells(i, debitCol).Value - Cells(i, creditCol).Value

            curRow = curRow + 1
        End If
    Next i

    'Format the values
    Range(Cells(startRow, creditCol + 2), Cells(curRow, creditCol + 2)).HorizontalAlignment = xlCenter
    Range(Cells(startRow, creditCol + 4), Cells(curRow, creditCol + 4)).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Range(Cells(1, creditCol + 2), Cells(1, creditCol + 4)).Columns.EntireColumn.AutoFit

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

