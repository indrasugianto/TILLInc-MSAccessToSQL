' Module Name: Form_frmPeopleClientsServiceVocational
' Module Type: Document Module
' Lines of Code: 51
' Extracted: 1/29/2026 4:12:26 PM

Option Compare Database
Option Explicit

Private Sub Form_Current()
    Call ProgressMessages("Append", "   Opening vocational services form.")

    If Inactive Then
        Me.Caption = "Client: Vocational Services (INACTIVE)"
    Else
        Me.Caption = "Client: Vocational Services"
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
    If Len(ContractNumber) > 0 Then ActivityCode = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber & """")
End Sub

Private Sub Form_GotFocus()
    If IsNull(CityTown) Then DoCmd.OpenForm "frmPeopleClientsSelectLocation", , , , , , "Voc"
End Sub

Private Sub Inactive_Click()
    If Inactive Then
        Call CheckInactive(DateInactive, "VocationalSetInactive", "True")
        Call GreyAndNormal(Form_frmPeople.IsClientVocatLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices - 1
        If Form_frmPeople.NumberOfClientServices < 0 Then Form_frmPeople.NumberOfClientServices = 0
    Else
        Call CheckInactive(DateInactive, "VocationalSetInactive", "False")
        Call BlueAndBold(Form_frmPeople.IsClientVocatLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
End Sub

Private Sub SelectResidence_Click()
    DoCmd.OpenForm "frmPeopleClientsSelectLocation", , , , , , "Voc"
End Sub

Private Sub ContractNumber_AfterUpdate()
    If Len(ContractNumber) > 0 Then
        ActivityCode = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber & """")
        Call UpdateChangeLog("VOCContractNumber1", ContractNumber)
        Call UpdateChangeLog("VOCActivityCode1", ActivityCode)
    End If
End Sub

Private Sub ContractNumber2_AfterUpdate()
    If Len(ContractNumber2) > 0 Then
        ActivityCode2 = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber2 & """")
        Call UpdateChangeLog("VOCContractNumber2", ContractNumber2)
        Call UpdateChangeLog("VOCActivityCode2", ActivityCode2)
    End If
End Sub