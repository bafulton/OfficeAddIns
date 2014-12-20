'Blackman & Sloop Excel Add-In, v1.1 (7/8/13)

Private WithEvents App As Application

Private Sub Workbook_Open()
    Set App = Application
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    AddHyperlinks.SetDestination
End Sub