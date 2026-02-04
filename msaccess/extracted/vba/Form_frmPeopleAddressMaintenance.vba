' Module Name: Form_frmPeopleAddressMaintenance
' Module Type: Document Module
' Lines of Code: 15
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Private Sub Form_Load()
    Me.Requery
End Sub

Private Sub Selected_Click()
    GlobalAddress = CandidateAddress
    GlobalCity = CandidateCity
    GlobalState = CandidateState
    GlobalZIP = CandidateZIP
    GlobalValidated = True
    DoCmd.Close
End Sub
