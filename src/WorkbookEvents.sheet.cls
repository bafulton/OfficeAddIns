'Blackman & Sloop Excel Add-In, v1.2 (5/15/14)

Private WithEvents App As Application

Private Sub Workbook_Open()
    Set App = Application
    
    'Check for any updates
    CheckUpdate
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    AddHyperlinks.SetDestination
End Sub
