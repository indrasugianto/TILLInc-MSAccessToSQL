' Module Name: Report_rptPCABILLING
' Module Type: Document Module
' Lines of Code: 19
' Extracted: 2026-02-04 13:03:36

Option Compare Database

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)
    If IsNull(Age) Then
        AgeSubfooter = "Age ? Subtotals"
    Else
        AgeSubfooter = "Age " & [Age] & " Subtotals"
    End If
End Sub

Private Sub GroupHeader0_Format(Cancel As Integer, FormatCount As Integer)
    If IsNull(Age) Then
        AgeUnknown.Visible = True
        Age.Visible = False
    Else
        AgeUnknown.Visible = False
        Age.Visible = True
    End If
End Sub
