' Module Name: Form_frmDBChanges
' Module Type: Document Module
' Lines of Code: 6
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Private Sub Button_NewRecord_Click()
    DoCmd.GoToRecord acDataForm, Me.Name, acNewRec
End Sub
