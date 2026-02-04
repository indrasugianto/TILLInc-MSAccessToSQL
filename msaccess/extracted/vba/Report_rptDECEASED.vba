' Module Name: Report_rptDECEASED
' Module Type: Document Module
' Lines of Code: 10
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    If IsClient Then
        rptDECEASEDClientFamily.Visible = True: rptDECEASEDFamilyClient.Visible = False
    ElseIf IsFamilyGuardian Then
        rptDECEASEDClientFamily.Visible = False: rptDECEASEDFamilyClient.Visible = True
    End If
End Sub
