' Module Name: Form_frmDBChanges
' Module Type: Document Module
' Lines of Code: 6
' Extracted: 1/29/2026 4:12:27 PM

Option Compare Database
Option Explicit

Private Sub Button_NewRecord_Click()
    DoCmd.GoToRecord acDataForm, Me.Name, acNewRec
End Sub