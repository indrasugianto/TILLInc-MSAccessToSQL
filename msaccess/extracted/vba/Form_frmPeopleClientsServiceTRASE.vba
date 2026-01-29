' Module Name: Form_frmPeopleClientsServiceTRASE
' Module Type: Document Module
' Lines of Code: 24
' Extracted: 1/29/2026 4:12:24 PM

Option Compare Database
Option Explicit

Private Sub Form_Current()
    If Inactive Then
        Me.Caption = "Client: TRASE (INACTIVE)"
    Else
        Me.Caption = "Client: TRASE"
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
End Sub

Private Sub Inactive_Click()
    If Inactive Then
        Call CheckInactive(DateInactive, "TRASESetInactive", "True")
        Call GreyAndNormal(Form_frmPeople.IsClientTRASELabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices - 1
        If Form_frmPeople.NumberOfClientServices < 0 Then Form_frmPeople.NumberOfClientServices = 0
    Else
        Call CheckInactive(DateInactive, "TRASESetInactive", "False")
        Call BlueAndBold(Form_frmPeople.IsClientTRASELabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
End Sub