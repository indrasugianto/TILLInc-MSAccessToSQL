' Module Name: Report_~rptTEMPLATELandscape
' Module Type: Document Module
' Lines of Code: 19
' Extracted: 1/29/2026 4:12:26 PM

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