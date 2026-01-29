' Module Name: Report_ltrAUTODEPOSITAUTHTILL
' Module Type: Document Module
' Lines of Code: 6
' Extracted: 1/29/2026 4:12:27 PM

Option Compare Database
Option Explicit

Private Sub Report_Close()
    DoCmd.Close acForm, "frmRptLtrEWDLSTILL"
End Sub