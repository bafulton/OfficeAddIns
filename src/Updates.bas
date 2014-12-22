Attribute VB_Name = "Updates"
'Blackman & Sloop Excel Add-In, v1.2 (5/15/14)

Public needUpdate As Boolean, checked As Boolean, nVersion, nPath, lPath

Function CheckUpdate()
    Dim lVersion, lastCheck
    nPath = "F:\IT Data\Add-In"
    lPath = Workbooks("Blackman and Sloop Add-In.xlam").Path

    'Only check if we haven't already checked
    If checked Then GoTo Ret_Value

    'Get the current version
    On Error GoTo No_Local
    Open lPath & "\Version.txt" For Input As #1
    Line Input #1, lVersion
    Line Input #1, lastCheck
    Close #1

    'Check for an update once a week
    If Date <= DateValue(lastCheck) + 7 Then
        checked = True: needUpdate = False
        GoTo Ret_Value
    End If

    'Get the latest version
    On Error GoTo No_Network
    Open nPath & "\Logs\Latest Version.txt" For Input As #1
    Line Input #1, nVersion
    Close #1

    On Error GoTo Err_Handler

    'Compare the local version to the network version
    If lVersion <> nVersion Then
        MsgBox "An update is available for the B&S add-in.", , "Update Available"
        checked = True: needUpdate = True
    Else
        checked = True: needUpdate = False

        'Update the local log
        Open lPath & "\Version.txt" For Output As #1
        Print #1, lVersion
        Print #1, Date
        Close #1
    End If

    'Update the network log
    User = Application.UserName
    logPath = nPath & "\Logs\Check Log.csv"
    Open logPath For Append As #1
    Print #1, Date & ", " & User & ", " & lVersion
    Close #1

    GoTo Ret_Value

Err_Handler:
    Close #1
    GoTo No_Network
No_Local:
    checked = True: needUpdate = True
    GoTo Ret_Value
No_Network:
    checked = True: needUpdate = False
    GoTo Ret_Value
Ret_Value:
    CheckUpdate = needUpdate
End Function

Sub GetUpdate(control As IRibbonControl)
    Shell nPath & "\Install.bat ", vbNormalFocus
End Sub

Public Function getUpdateLabel(control As IRibbonControl, ByRef label)
    If CheckUpdate() Then
        label = "Update"
    Else
        label = ""
    End If
End Function

Public Function getUpdateVisible(control As IRibbonControl, ByRef visible)
    If CheckUpdate() Then
        visible = True
    Else
        visible = False
    End If
End Function
