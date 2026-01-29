' Module Name: Report_rptDirectoryAlphaShortRes
' Module Type: Document Module
' Lines of Code: 10
' Extracted: 1/29/2026 4:12:28 PM

Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    If FullName = "*** OPEN ***" Then
        FullName.FontWeight = 800
    Else
        FullName.FontWeight = 400
    End If
End Sub