' Module Name: Form_frmPeopleClientsServiceAdultCompanion
' Module Type: Document Module
' Lines of Code: 53
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Private Sub Form_Current()
    Call ProgressMessages("Append", "   Opening adult companion form.")
    If IsNull(CityTown) Then
        CityTown = "Dedham"
        Location = "Adult Companion"
    End If
    
    If Inactive Then
        Me.Caption = "Client: Adult Companion (INACTIVE)"
    Else
        Me.Caption = "Client: Adult Companion"
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
    If Len(ContractNumber) > 0 Then ActivityCode = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber & """")
End Sub

Private Sub Inactive_Click()
    If Inactive Then
        Call CheckInactive(DateInactive, "AdultCompSetInactive", "True")
        Call GreyAndNormal(Form_frmPeople.IsClientAdultCompLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices - 1
        If Form_frmPeople.NumberOfClientServices < 0 Then Form_frmPeople.NumberOfClientServices = 0
    Else
        Call CheckInactive(DateInactive, "AdultCompSetInactive", "False")
        Call BlueAndBold(Form_frmPeople.IsClientAdultCompLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
End Sub

Private Sub SelectBroker_Click()
    DoCmd.OpenForm "frmPeopleSelectPerson", , , , , , "AdultCompCaseManager"
End Sub

Private Sub ContractNumber_AfterUpdate()
    If Len(ContractNumber) > 0 Then ActivityCode = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber & """") Else ActivityCode = ""
    Call UpdateChangeLog("AdultCompContractNumber", ContractNumber)
    Call UpdateChangeLog("AdultCompActivityCode", ActivityCode)
End Sub

Private Sub ContractNumber2_AfterUpdate()
    If Len(ContractNumber2) > 0 Then ActivityCode2 = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber2 & """") Else ActivityCode2 = ""
End Sub

Private Sub StartDate_AfterUpdate()
    If ValidateDateString(StartDate) Then Call UpdateChangeLog("ISSStartDate", [StartDate])
End Sub

Private Sub EndDate_AfterUpdate()
    If ValidateDateString(EndDate) Then Call UpdateChangeLog("ISSEndDate", EndDate)
End Sub
