' Module Name: Report_rptDirectoryAlphaShortResLARGE
' Module Type: Document Module
' Lines of Code: 10
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    If FullName = "*** OPEN ***" Then
        FullName.FontWeight = 800
    Else
        FullName.FontWeight = 400
    End If
End Sub
