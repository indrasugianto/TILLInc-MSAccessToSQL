' Module Name: Form_frmPeopleClientsServiceCommunityConnections
' Module Type: Document Module
' Lines of Code: 44
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Private Sub Bowlers_AfterUpdate()
    If Bowlers Then Call UpdateChangeLog("CCBowler", "TRUE") Else Call UpdateChangeLog("CCBowler", "FALSE")
End Sub

Private Sub Bowlers_Click()
    BowlingTeam.Visible = Bowlers
End Sub

Private Sub Form_Current()
    Call ProgressMessages("Append", "   Opening community connections form.")
    If IsNull(Inactive) Then
        If Form_frmMainMenu.NewRecordInProgress Then
            RecordAddedDate = Format(Now(), "mm/dd/yyyy")
            RecordAddedBy = Form_frmMainMenu.UserName
            Bowlers = False
            BowlingTeam = Null
            Inactive = False
        End If
    End If
    If Inactive Then
        Me.Caption = "Client: Recreation (INACTIVE)"
    Else
        Me.Caption = "Client: Recreation"
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
    BowlingTeam.Visible = Bowlers
    Exit Sub
End Sub

Private Sub Inactive_Click()
    If Inactive Then
        Call CheckInactive(DateInactive, "CommunityConnectionsSetInactive", "True")
        Call GreyAndNormal(Form_frmPeople.IsClientCommunityConnectionsLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices - 1
        If Form_frmPeople.NumberOfClientServices < 0 Then Form_frmPeople.NumberOfClientServices = 0
    Else
        Call CheckInactive(DateInactive, "CommunityConnectionsSetInactive", "False")
        Call BlueAndBold(Form_frmPeople.IsClientCommunityConnectionsLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
End Sub
