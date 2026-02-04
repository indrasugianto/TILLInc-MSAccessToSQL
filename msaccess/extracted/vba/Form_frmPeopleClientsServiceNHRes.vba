' Module Name: Form_frmPeopleClientsServiceNHRes
' Module Type: Document Module
' Lines of Code: 42
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Private Sub Form_Current()
    Call ProgressMessages("Append", "   Opening NH residential services form.")
    If Inactive Then
        Me.Caption = "Client: NH Residential Services (INACTIVE)"
    Else
        Me.Caption = "Client: NH Residential Services"
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
    If Len(ContractNumber) > 0 Then ActivityCode = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber & """")
    If IsNull(CaseManager) Or Len(CaseManager) = 0 Then CaseManager = DLookup("WhoGetsIt", "catFillInNames", "Category='NH Residential Case Manager'")
End Sub

Private Sub Inactive_Click()
    If Inactive Then
        Call CheckInactive(DateInactive, "NHResSetInactive", "True")
        Call GreyAndNormal(Form_frmPeople.IsClientNHResLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices - 1
        If Form_frmPeople.NumberOfClientServices < 0 Then Form_frmPeople.NumberOfClientServices = 0
    Else
        Call CheckInactive(DateInactive, "NHResSetInactive", "False")
        Call BlueAndBold(Form_frmPeople.IsClientNHResLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
End Sub

Private Sub ContractNumber_AfterUpdate()
    If Len(ContractNumber) > 0 Then
        ActivityCode = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber & """")
        Call UpdateChangeLog("NHResContract", ContractNumber)
        Call UpdateChangeLog("NHResActivityCode", ActivityCode)
    Else
        ContractNumber = ""
        ActivityCode = ""
    End If
End Sub

Private Sub SelectBroker_Click()
'   DoCmd.OpenForm "frmPeopleSelectPerson", , , , , , "NH Residential Case Manager"
End Sub
