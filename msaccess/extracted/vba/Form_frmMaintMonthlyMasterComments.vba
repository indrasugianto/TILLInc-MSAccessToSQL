' Module Name: Form_frmMaintMonthlyMasterComments
' Module Type: Document Module
' Lines of Code: 6
' Extracted: 1/29/2026 4:12:23 PM

Option Compare Database
Option Explicit

Private Sub OK_Click()
    Form_frmMaintMonthlyMaster.Processed = True: Form_frmMaintMonthlyMaster.Applied = True: DoCmd.Close
End Sub