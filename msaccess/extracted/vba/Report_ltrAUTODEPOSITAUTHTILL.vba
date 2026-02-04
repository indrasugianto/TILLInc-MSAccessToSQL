' Module Name: Report_ltrAUTODEPOSITAUTHTILL
' Module Type: Document Module
' Lines of Code: 6
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Private Sub Report_Close()
    DoCmd.Close acForm, "frmRptLtrEWDLSTILL"
End Sub
