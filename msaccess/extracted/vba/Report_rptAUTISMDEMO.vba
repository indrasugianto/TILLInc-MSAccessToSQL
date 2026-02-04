' Module Name: Report_rptAUTISMDEMO
' Module Type: Document Module
' Lines of Code: 17
' Extracted: 2026-02-04 13:03:36

Option Compare Database

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    If Gender = "M" Or Gender = "F" Then GenderDisplayed = Gender Else GenderDisplayed = "?"
End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)
    If IsNull(Age) Then AgeSubfooter = "Age ? Subtotals" Else AgeSubfooter = "Age " & [Age] & " Subtotals"
End Sub

Private Sub GroupHeader0_Format(Cancel As Integer, FormatCount As Integer)
    If IsNull(Age) Then
        AgeUnknown.Visible = True: Age.Visible = False
    Else
        AgeUnknown.Visible = False: Age.Visible = True
    End If
End Sub
