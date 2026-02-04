' Module Name: Form_frmMaintMonthlyMasterComments
' Module Type: Document Module
' Lines of Code: 6
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Private Sub OK_Click()
    Form_frmMaintMonthlyMaster.Processed = True: Form_frmMaintMonthlyMaster.Applied = True: DoCmd.Close
End Sub
