' Module Name: Form_frmPeopleConsultants
' Module Type: Document Module
' Lines of Code: 62
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Dim SPRFullName As String

Private Sub Form_Current()
    Me.Caption = "Consultants"
    Call Department_AfterUpdate
End Sub

Private Sub Department_AfterUpdate()
    Select Case Department
        Case "Springboard"
            SpringboardGroupCode1.Visible = True
            SpringboardGroupCode2.Visible = True
            SpringboardGroupCode3.Visible = True
        Case Else
            SpringboardGroupCode1.Visible = False
            SpringboardGroupCode2.Visible = False
            SpringboardGroupCode3.Visible = False
    End Select
    Call UpdateChangeLog("ConsultantDepartment", Department)
End Sub

Private Sub Inactive_Click()
    Call CheckPersonCompletelyInactive
End Sub

Private Sub SpringboardGroupCode1_AfterUpdate()
    If IsNull(SpringboardGroupCode1) Then
        TILLDataBase.Execute "UPDATE qrytblPeopleClientsSpringboardServices SET LeaderIndexedName = '' WHERE (GroupCode is Null) OR (Len(GroupCode) <= 0);", dbSeeChanges: Call BriefDelay
    Else
        SPRFullName = Form_frmPeople.FirstName & " " & Form_frmPeople.MiddleInitial & " " & Form_frmPeople.LastName
        TILLDataBase.Execute "UPDATE qrytblPeopleClientsSpringboardServices SET qrytblPeopleClientsSpringboardServices.LeaderIndexedName = """ & Form_frmPeople.IndexedName & """, qrytblPeopleClientsSpringboardServices.Leader = '" & SPRFullName & "' " & _
            "WHERE qrytblPeopleClientsSpringboardServices.GroupCode='" & Form_frmPeopleConsultants.SpringboardGroupCode1 & "';", dbSeeChanges: Call BriefDelay
    End If
    Call UpdateChangeLog("ConsultantSpringboardGroupCode1", SpringboardGroupCode1)
End Sub

Private Sub SpringboardGroupCode2_AfterUpdate()
    If IsNull(SpringboardGroupCode2) Then
        TILLDataBase.Execute "UPDATE qrytblPeopleClientsSpringboardServices SET LeaderIndexedName = '' WHERE (GroupCode is Null) OR (Len(GroupCode) <= 0);", dbSeeChanges: Call BriefDelay
'           "WHERE qrytblPeopleClientsSpringboardServices.GroupCode=" & Form_frmPeopleConsultants.SpringboardGroupCode2 & ";", dbseechanges: call briefdelay
    Else
        SPRFullName = Form_frmPeople.FirstName & " " & Form_frmPeople.MiddleInitial & " " & Form_frmPeople.LastName
        TILLDataBase.Execute "UPDATE qrytblPeopleClientsSpringboardServices SET qrytblPeopleClientsSpringboardServices.LeaderIndexedName = """ & Form_frmPeople.IndexedName & """, qrytblPeopleClientsSpringboardServices.Leader = '" & SPRFullName & "' " & _
            "WHERE qrytblPeopleClientsSpringboardServices.GroupCode='" & Form_frmPeopleConsultants.SpringboardGroupCode2 & "';", dbSeeChanges: Call BriefDelay
    End If
    Call UpdateChangeLog("ConsultantSpringboardGroupCode2", SpringboardGroupCode2)
End Sub

Private Sub SpringboardGroupCode3_AfterUpdate()
    If IsNull(SpringboardGroupCode3) Then
        TILLDataBase.Execute "UPDATE qrytblPeopleClientsSpringboardServices SET LeaderIndexedName = '' WHERE (GroupCode is Null) OR (Len(GroupCode) <= 0);", dbSeeChanges: Call BriefDelay
'           "WHERE qrytblPeopleClientsSpringboardServices.GroupCode=" & Form_frmPeopleConsultants.SpringboardGroupCode3 & ";", dbseechanges: call briefdelay
    Else
        TILLDataBase.Execute "UPDATE qrytblPeopleClientsSpringboardServices SET qrytblPeopleClientsSpringboardServices.LeaderIndexedName = """ & Form_frmPeople.IndexedName & """, qrytblPeopleClientsSpringboardServices.Leader = '" & SPRFullName & "' " & _
            "WHERE qrytblPeopleClientsSpringboardServices.GroupCode='" & Form_frmPeopleConsultants.SpringboardGroupCode3 & "';", dbSeeChanges: Call BriefDelay
    End If
    Call UpdateChangeLog("ConsultantSpringboardGroupCode3", SpringboardGroupCode3)
End Sub

