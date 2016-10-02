Attribute VB_Name = "QBTBClean"
'Blackman & Sloop Excel Add-In, v1.2 (5/15/14)

Option Base 1    ' Set default array subscripts to 1

Sub CleanTheTB(control As IRibbonControl)
    On Error Resume Next

    msg = MsgBox("Exclude $0 balances?", vbYesNo, "$0 Balances")

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    'Find Debit, Credit, and Account columns and rows
    Dim MyAlphaArray
    MyAlphaArray = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J")
    Dim MyDebitCol As String
    Dim MyCreditCol As String
    Dim MyAcctCol As String
    Dim MyRange
    MyRange = "A2:Z2"
    Dim MyCollection
    Dim MyObject
    Set MyCollection = Range(MyRange)
    For Each MyObject In MyCollection
        If MyObject.Value = "Debit" Then
            MyPointer = MyObject.Column
            MyDebitCol = MyAlphaArray(MyObject.Column)
        End If
        If MyObject.Value = "Credit" Then
            MyPointer = MyObject.Column
            MyCreditCol = MyAlphaArray(MyObject.Column)
        End If
    Next
    MyRange = "A3:Z3"
    Set MyCollection = Range(MyRange)
    For Each MyObject In MyCollection
        MyPos = InStr(1, MyObject.Value, " · ")
        If MyPos > 0 Then
            MyPointer = MyObject.Column
            MyAcctCol = "MyAlphaArray(MyObject.Column)"
        Else
            MyAcctCol = "B"
        End If
    Next

    'Split each row into the appropriate arrays
    MyRows = Application.CountA(ActiveSheet.Range(MyAcctCol & ":" & MyAcctCol))

    Dim MyIndex As Integer
    MyIndex = MyRows + 2
    Dim AccountNumArray(1000) As Variant
    Dim AccountNameArray(1000) As Variant
    Dim AmountArray(1000) As Double
    Dim MyCnt As Integer

    For MyCnt = 3 To MyIndex
        MyPos = InStr(1, Range(MyAcctCol & MyCnt).Value, " · ")
        If MyPos > 0 Then
            StringtoCheck = Range(MyAcctCol & MyCnt).Value
            NeedLoop = "True"
            Do While NeedLoop = "True"
                MyPos2 = InStr(1, StringtoCheck, ":")
                If MyPos2 > 0 Then
                    StringtoCheck = Mid(StringtoCheck, MyPos2 + 1)
                Else
                    NeedLoop = "False"
                End If
            Loop

            AccountNumArray(MyCnt) = SplitAccount(StringtoCheck, "Num")
            AccountNameArray(MyCnt) = SplitAccount(StringtoCheck, "Name")

            If Not (Range(MyDebitCol & MyCnt).Value = 0) Then
                AmountArray(MyCnt) = Range(MyDebitCol & MyCnt).Value
            End If
            If Not (Range(MyCreditCol & MyCnt).Value) = 0 Then
                AmountArray(MyCnt) = (Range(MyCreditCol & MyCnt).Value * -1)
            End If
        Else
            AccountNameArray(MyCnt) = Range(MyAcctCol & MyCnt).Value

            If Not (Range(MyDebitCol & MyCnt).Value = 0) Then
                AmountArray(MyCnt) = Range(MyDebitCol & MyCnt).Value
            End If
            If Not (Range(MyCreditCol & MyCnt).Value) = 0 Then
                AmountArray(MyCnt) = (Range(MyCreditCol & MyCnt).Value * -1)
            End If
        End If
        
     Next MyCnt

    'Write output
    Dim Sh As Worksheet, flg As Boolean
    For Each Sh In Worksheets
        If Sh.Name = "Cleaned TB" Then flg = True: Exit For
    Next
    If flg = True Then
        Application.DisplayAlerts = False
        Sheets("Cleaned TB").Delete
        Application.DisplayAlerts = True
    End If
    Sheets.Add.Name = "Cleaned TB"
    Sheets("Cleaned TB").Select
    Range("A1").Select
    Dim NextRow As Integer
    NextRow = 1
    For MyCnt = 3 To MyIndex
        'skip $0 balances if instructed to do so
        If Not (msg = vbYes And AmountArray(MyCnt) = 0) Then
            Range("A" & NextRow).Value = AccountNumArray(MyCnt)
            Range("B" & NextRow).Value = AccountNameArray(MyCnt)
            Range("C" & NextRow).Value = AmountArray(MyCnt)
            NextRow = NextRow + 1
        End If
    Next MyCnt

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
