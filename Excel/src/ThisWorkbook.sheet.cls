'Blackman & Sloop Excel Add-In, v1.4 (10/2/16)

Private WithEvents App As Application

Private Sub Workbook_Open()
    Set App = Application
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    AddHyperlinks.SetDestination
End Sub