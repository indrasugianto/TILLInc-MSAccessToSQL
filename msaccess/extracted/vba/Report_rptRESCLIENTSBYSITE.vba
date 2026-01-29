' Module Name: Report_rptRESCLIENTSBYSITE
' Module Type: Document Module
' Lines of Code: 6
' Extracted: 1/29/2026 4:12:26 PM

Option Compare Database
Option Explicit

Private Sub GroupHeader2_Format(Cancel As Integer, FormatCount As Integer)
    If NumVacancies = 0 Then NumVacancies.Visible = False Else NumVacancies.Visible = True
End Sub