' Module Name: Form_frmPeopleClientsServiceTransportation
' Module Type: Document Module
' Lines of Code: 29
' Extracted: 1/29/2026 4:12:24 PM

Option Compare Database
Option Explicit

Private Sub Form_Current()
    If Inactive Then Me.Caption = "Client: Transportation Services (INACTIVE)" Else Me.Caption = "Client: Transportation Services"
    If Len(ContractNumber) > 0 Then ActivityCode = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber & """")
End Sub

Private Sub Inactive_Click()
    If Inactive Then
        Call CheckInactive(DateInactive, "TransportationSetInactive", "True")
        Call GreyAndNormal(Form_frmPeople.IsClientTransLabel)
    Else
        Call CheckInactive(DateInactive, "TransportationSetInactive", "False")
        Call BlueAndBold(Form_frmPeople.IsClientTransLabel)
    End If
End Sub

Private Sub ContractNumber_AfterUpdate()
    If Len(ContractNumber) > 0 Then ActivityCode = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber & """")
    Call UpdateChangeLog("TRANSContractNumber1", ContractNumber)
    Call UpdateChangeLog("TRANSActivityCode1", ActivityCode)
End Sub

Private Sub ContractNumber2_AfterUpdate()
    If Len(ContractNumber2) > 0 Then ActivityCode2 = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber2 & """")
    Call UpdateChangeLog("TRANSContractNumber2", ContractNumber2)
    Call UpdateChangeLog("TRANSActivityCode2", ActivityCode2)
End Sub