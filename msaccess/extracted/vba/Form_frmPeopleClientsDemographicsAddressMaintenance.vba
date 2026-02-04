' Module Name: Form_frmPeopleClientsDemographicsAddressMaintenance
' Module Type: Document Module
' Lines of Code: 27
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Private Sub Form_Current()
    Me.Caption = "Select Address"
    MatchCodeDescription = DLookup("Description", "catAddressMatchCodes", "MatchCode=""" & MatchCode & """")
End Sub

Private Sub Form_Load()
    Me.Requery
End Sub

Private Sub Selected_Click()
    With Form_frmPeopleClientsDemographics
        .RepPayeeAddress = CandidateAddress
        Call UpdateChangeLog("RepPayeeAddress", CandidateAddress)
        .RepPayeeCity = CandidateCity
        Call UpdateChangeLog("RepPayeeCity", CandidateCity)
        .RepPayeeState = CandidateState
        Call UpdateChangeLog("RepPayeeState", CandidateState)
        .RepPayeeZIP = CandidateZIP
        Call UpdateChangeLog("RepPayeeZIP", CandidateZIP)
        .RepPayeeAddressValidated = True
    End With
    If IsTableQuery("temptbl") Then TILLDataBase.Execute "DROP TABLE temptbl", dbSeeChanges: Call BriefDelay
    DoCmd.Close
End Sub
