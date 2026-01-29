' Module Name: Report_rptSPRINGBOARDClients
' Module Type: Document Module
' Lines of Code: 6
' Extracted: 1/29/2026 4:12:28 PM

Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    If IsDeceased Then IsDeceased.Visible = True Else IsDeceased.Visible = False
End Sub