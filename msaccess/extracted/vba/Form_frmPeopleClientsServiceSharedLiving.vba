' Module Name: Form_frmPeopleClientsServiceSharedLiving
' Module Type: Document Module
' Lines of Code: 57
' Extracted: 1/29/2026 4:12:26 PM

Option Compare Database
Option Explicit

Private Sub Form_Current()
    Call ProgressMessages("Append", "   Opening shared living services form.")

    If IsNull(CityTown) Then
        CityTown = "Dedham"
        Location = "Shared Living"
    End If
    
    If Inactive Then
        Me.Caption = "Client: Shared Living (INACTIVE)"
    Else
        Me.Caption = "Client: Shared Living"
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
    If Len(ContractNumber) > 0 Then ActivityCode = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber & """")
End Sub

Private Sub Inactive_Click()
    If Inactive Then
        Call CheckInactive(DateInactive, "SharedLivingSetInactive", "True")
        Call GreyAndNormal(Form_frmPeople.IsClientSharedLivingLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices - 1
        If Form_frmPeople.NumberOfClientServices < 0 Then Form_frmPeople.NumberOfClientServices = 0
    Else
        Call CheckInactive(DateInactive, "SharedLivingSetInactive", "False")
        Call BlueAndBold(Form_frmPeople.IsClientSharedLivingLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
End Sub

Private Sub ContractNumber_AfterUpdate()
    If Len(ContractNumber) > 0 Then
        ActivityCode = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber & """")
        Call UpdateChangeLog("SharedLivingContractNumber", ContractNumber)
        Call UpdateChangeLog("SharedLivingActivityCode", ActivityCode)
    Else
        ContractNumber = ""
        ActivityCode = ""
    End If
End Sub

Private Sub ContractNumber2_AfterUpdate()
    If Len(ContractNumber2) > 0 Then
        ActivityCode2 = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber2 & """")
        Call UpdateChangeLog("SharedLivingContractNumber2", ContractNumber2)
        Call UpdateChangeLog("SharedLivingActivityCode2", ActivityCode2)
    Else
        ContractNumber2 = ""
    End If
End Sub

Private Sub SelectBroker_Click()
    DoCmd.OpenForm "frmPeopleSelectPerson", , , , , , "SharedLivingCaseManager"
End Sub