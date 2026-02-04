' Module Name: Report_rptRESCLIENTSBYSITE
' Module Type: Document Module
' Lines of Code: 6
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Private Sub GroupHeader2_Format(Cancel As Integer, FormatCount As Integer)
    If NumVacancies = 0 Then NumVacancies.Visible = False Else NumVacancies.Visible = True
End Sub
