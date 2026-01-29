' Module Name: Report_rptCLIENTERRORS
' Module Type: Document Module
' Lines of Code: 6
' Extracted: 1/29/2026 4:12:27 PM

Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    DisplayErrorMessages = ErrorMessages
End Sub