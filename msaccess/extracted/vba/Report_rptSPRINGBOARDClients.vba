' Module Name: Report_rptSPRINGBOARDClients
' Module Type: Document Module
' Lines of Code: 6
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    If IsDeceased Then IsDeceased.Visible = True Else IsDeceased.Visible = False
End Sub
