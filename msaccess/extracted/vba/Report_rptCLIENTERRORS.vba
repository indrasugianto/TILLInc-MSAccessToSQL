' Module Name: Report_rptCLIENTERRORS
' Module Type: Document Module
' Lines of Code: 6
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    DisplayErrorMessages = ErrorMessages
End Sub
