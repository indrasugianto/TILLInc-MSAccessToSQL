' Module Name: Form_frmPeopleClientsServiceSpringboard
' Module Type: Document Module
' Lines of Code: 123
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Dim Subject As String, TextBody As String

Private Sub BeginBillingDate_AfterUpdate()
    Subject = "TILLDB: Springboard Member " & Form_frmPeople.FirstName & " " & Form_frmPeople.LastName & ": Begin Billing Date Set/Changed"
    TextBody = "Begin billing date for " & Form_frmPeople.FirstName & " " & Form_frmPeople.LastName & " was set or changed to " & BeginBillingDate & " by " & Form_frmMainMenu.UserName & "." & vbCrLf & "    " & Form_frmPeople.PhysicalAddress & vbCrLf & "    " & Form_frmPeople.PhysicalCity & " " & Form_frmPeople.PhysicalState & " " & Form_frmPeople.PhysicalZIP & vbCrLf
    Call SendEmailMessage("tilldbnotifications@tillinc.org", DLookup("ParameterValue", "appParameters", "ID=22"), "db.springboard@tillinc.org", Null, Subject, TextBody, Null)
End Sub

Private Sub CustomerID_AfterUpdate()
    If IsNull(CustomerID) Then CustomerID.BackColor = RGB(255, 0, 0) Else CustomerID.BackColor = RGB(255, 255, 255)
End Sub

Private Sub Form_Current()
    Call ProgressMessages("Append", "   Opening Springobard form.")

    CustomerID.SetFocus
    
    If Inactive Then
        Me.Caption = "Client: Springboard (INACTIVE)"
    Else
        Me.Caption = "Client: Springboard"
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If

    If IsNull(CustomerID) Then CustomerID.BackColor = RGB(255, 0, 0) Else CustomerID.BackColor = RGB(255, 255, 255)

    If Form_frmMainMenu.NewRecordInProgress Then
        RecordAddedDate = Format(Now(), "mm/dd/yyyy")
        RecordAddedBy = Form_frmMainMenu.UserName
    End If

    If Len(GroupCode) > 0 Then
        If DCount("IndexedName", "tblPeopleConsultants", "SpringboardGroupCode1=""" & GroupCode & """") > 0 Then
            LeaderIndexedName = DLookup("IndexedName", "tblPeopleConsultants", "SpringboardGroupCode1='" & GroupCode & "' AND Inactive=False")
        Else
            If DCount("IndexedName", "tblPeopleConsultants", "SpringboardGroupCode2=""" & GroupCode & """") > 0 Then
                LeaderIndexedName = DLookup("IndexedName", "tblPeopleConsultants", "SpringboardGroupCode2='" & GroupCode & "' AND Inactive=False")
            Else
                If DCount("IndexedName", "tblPeopleConsultants", "SpringboardGroupCode3=""" & GroupCode & """") > 0 Then
                    LeaderIndexedName = DLookup("IndexedName", "tblPeopleConsultants", "SpringboardGroupCode3='" & GroupCode & "' AND Inactive=False")
                Else
                    LeaderIndexedName = ""
                End If
            End If
        End If
    End If

    If LeaderIndexedName <> "" _
    Then _
        Leader = DLookup("FirstName", "tblPeople", "IndexedName=""" & LeaderIndexedName & """") & " " & _
            DLookup("MiddleInitial", "tblPeople", "IndexedName=""" & LeaderIndexedName & """") & " " & _
            DLookup("LastName", "tblPeople", "IndexedName=""" & LeaderIndexedName & """") _
    Else _
        Leader = ""
    
    If Len(Form_frmPeopleClientsDemographics.DateOfBirth) = 10 Then Age = CalcAge(CDate(Form_frmPeopleClientsDemographics.DateOfBirth)) Else Age = Null
End Sub

Private Sub GroupCode_AfterUpdate()
    If Len(GroupCode) > 0 Then
        If DCount("IndexedName", "tblPeopleConsultants", "SpringboardGroupCode1=""" & GroupCode & """") > 0 Then
            LeaderIndexedName = DLookup("IndexedName", "tblPeopleConsultants", "SpringboardGroupCode1=""" & GroupCode & """")
        Else
            If DCount("IndexedName", "tblPeopleConsultants", "SpringboardGroupCode2=""" & GroupCode & """") > 0 Then
                LeaderIndexedName = DLookup("IndexedName", "tblPeopleConsultants", "SpringboardGroupCode2=""" & GroupCode & """")
            Else
                If DCount("IndexedName", "tblPeopleConsultants", "SpringboardGroupCode3=""" & GroupCode & """") > 0 Then
                    LeaderIndexedName = DLookup("IndexedName", "tblPeopleConsultants", "SpringboardGroupCode3=""" & GroupCode & """")
                Else
                    LeaderIndexedName = ""
                End If
            End If
        End If
    End If

    If LeaderIndexedName <> "" Then Leader = DLookup("FirstName", "tblPeople", "IndexedName=""" & LeaderIndexedName & """") & " " & _
            DLookup("MiddleInitial", "tblPeople", "IndexedName=""" & LeaderIndexedName & """") & " " & _
            DLookup("LastName", "tblPeople", "IndexedName=""" & LeaderIndexedName & """") _
    Else _
        Leader = ""
    
    Call UpdateChangeLog("SPRINGBOARDGroupCode", GroupCode)
End Sub

Private Sub Inactive_Click()
    Dim LN As Variant, FN As Variant, PA As Variant, PC As Variant, PS As Variant, PZ As Variant
    
    LN = DLookup("LastName", "tblPeople", "IndexedName='" & IndexedName & "'")
    FN = DLookup("FirstName", "tblPeople", "IndexedName='" & IndexedName & "'")
    PA = DLookup("PhysicalAddress", "tblPeople", "IndexedName='" & IndexedName & "'")
    PC = DLookup("PhysicalCity", "tblPeople", "IndexedName='" & IndexedName & "'")
    PS = DLookup("PhysicalState", "tblPeople", "IndexedName='" & IndexedName & "'")
    PZ = DLookup("PhysicalZIP", "tblPeople", "IndexedName='" & IndexedName & "'")
    
    If Inactive Then
        Call CheckInactive(DateInactive, "SpringboardSetInactive", "True")
        Subject = "TILLDB: Springboard Member " & FN & " " & LN & " Inactive"
        TextBody = FN & " " & LN & " was flagged as INACTIVE by " & Form_frmMainMenu.UserName & "." & vbCrLf & "    " & PA & vbCrLf & "    " & PC & " " & PS & " " & PZ & vbCrLf
        Call SendEmailMessage("tilldbnotifications@tillinc.org", DLookup("ParameterValue", "appParameters", "ID=22"), "db.springboard@tillinc.org", Null, Subject, TextBody, Null)
        Call GreyAndNormal(Form_frmPeople.IsClientSpringLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices - 1
        If Form_frmPeople.NumberOfClientServices < 0 Then Form_frmPeople.NumberOfClientServices = 0
    Else
        Call CheckInactive(DateInactive, "SpringboardSetInactive", "False")
        Subject = "TILLDB: Springboard Member " & FN & " " & LN & " Re-activated"
        TextBody = FN & " " & LN & " was flagged as ACTIVE by " & Form_frmMainMenu.UserName & "." & vbCrLf & "    " & PA & vbCrLf & "    " & PC & " " & PS & " " & PZ & vbCrLf
        Call SendEmailMessage("tilldbnotifications@tillinc.org", DLookup("ParameterValue", "appParameters", "ID=22"), "db.springboard@tillinc.org", Null, Subject, TextBody, Null)
        Call BlueAndBold(Form_frmPeople.IsClientSpringLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
End Sub

Private Sub ReasonForTermination_Exit(Cancel As Integer)
    CustomerID.SetFocus
End Sub

Private Sub ReasonForTermination_LostFocus()
    CustomerID.SetFocus
End Sub

