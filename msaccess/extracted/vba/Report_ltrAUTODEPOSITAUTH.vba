' Module Name: Report_ltrAUTODEPOSITAUTH
' Module Type: Document Module
' Lines of Code: 6
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Private Sub Report_Close()
    DoCmd.Close acForm, "frmRptLtrEWDLS"
End Sub
