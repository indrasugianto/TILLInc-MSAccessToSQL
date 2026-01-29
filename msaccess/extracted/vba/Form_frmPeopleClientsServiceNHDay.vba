' Module Name: Form_frmPeopleClientsServiceNHDay
' Module Type: Document Module
' Lines of Code: 38
' Extracted: 1/29/2026 4:12:23 PM

Option Compare Database
Option Explicit

Private Sub Form_Current()
    Call ProgressMessages("Append", "   Opening NH day services form.")
    If Inactive Then
        Me.Caption = "Client: NH Day Services (INACTIVE)"
    Else
        Me.Caption = "Client: NH Day Services"
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
    If Len(ContractNumber) > 0 Then ActivityCode = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber & """")
End Sub

Private Sub Inactive_Click()
    If Inactive Then
        Call CheckInactive(DateInactive, "NHDaySetInactive", "True")
        Call GreyAndNormal(Form_frmPeople.IsClientNHDayLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices - 1
        If Form_frmPeople.NumberOfClientServices < 0 Then Form_frmPeople.NumberOfClientServices = 0
    Else
        Call CheckInactive(DateInactive, "NHDaySetInactive", "False")
        Call BlueAndBold(Form_frmPeople.IsClientNHDayLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
End Sub

Private Sub ContractNumber_AfterUpdate()
    If Len(ContractNumber) > 0 Then ActivityCode = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber & """")
    Call UpdateChangeLog("NHDayContractNumber1", ContractNumber)
    Call UpdateChangeLog("NHDayActivityCode1", ActivityCode)
End Sub

Private Sub ContractNumber2_AfterUpdate()
    If Len(ContractNumber2) > 0 Then ActivityCode2 = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber2 & """")
    Call UpdateChangeLog("NHDayContractNumber2", ContractNumber2)
    Call UpdateChangeLog("NHDayActivityCode2", ActivityCode2)
End Sub