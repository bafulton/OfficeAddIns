Attribute VB_Name = "EditCategories"
Sub Reminders()
    SetCategory "Reminders"
End Sub

Sub TaxWork()
    SetCategory "Tax Work"
End Sub

Sub AssuranceWork()
    SetCategory "Assurance Work"
End Sub

Private Sub SetCategory(ByVal Category As String)
    Dim Email As Object
    Set Email = Application.ActiveExplorer.Selection.Item(1)

    CurrentCategories = Split(Email.Categories, ",")
    If UBound(CurrentCategories) >= 0 Then
        ' Check for the category
        For i = 0 To UBound(CurrentCategories)
            If Trim(CurrentCategories(i)) = Category Then
                ' Remove the category
                CurrentCategories(i) = ""
                Email.Categories = Join(CurrentCategories, ",")
                ' Category removed; exit
                Exit Sub
            End If
        Next
    End If

    ' Category not found; add it
    Email.Categories = Category & "," & Email.Categories
End Sub
