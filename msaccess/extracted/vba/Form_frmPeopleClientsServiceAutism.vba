' Module Name: Form_frmPeopleClientsServiceAutism
' Module Type: Document Module
' Lines of Code: 100
' Extracted: 1/29/2026 4:12:23 PM

Option Compare Database
Option Explicit

Private Sub Comments_AfterUpdate()
    Call UpdateChangeLog("AutismComments", Comments)
End Sub

Private Sub CurrentAutismWaiverClient_Click()
    Dim TempDate As Date

    If CurrentAutismWaiverClient Then ' Child is a waiver client.
        If IsNull(Age) Then ' Check to make sure there's an age for this person.
            Form_frmPeople.ErrorMessages = Form_frmPeople.ErrorMessages & "[Sev High] There is no age set for this individual." & vbCrLf & "Please go to the Demographics page and enter a Date of Birth." & vbCrLf & " " & vbCrLf
            Form_frmPeople.ErrorMessages.BackStyle = 1
            CurrentAutismWaiverClient = False
            Exit Sub
        End If
        If Age >= 9 And CurrentAutismWaiverClient = True Then ' If child is 9 or more, cannot be a waiver client.
            MsgBox "Individual is 9 years of age or older." & vbCrLf & "Not eligible for Waiver Client status.  Resetting.", , "Note"
            CurrentAutismWaiverClient = False
            CurrentAutismWaiverClient.Visible = False
            Exit Sub
        End If
        ' Calculate an end date.
        TempDate = DateValue(Form_frmPeopleClientsDemographics.DateOfBirth) + DateAdd("yyyy", 9, TempDate)
        AutismWaiverEndDate.Visible = True
        AutismWaiverEndDate = Format(TempDate, "mm/dd/yyyy")
        ' Assume waiver start date is today.
        AutismWaiverStartDate = Format(Now(), "mm/dd/yyyy")
    End If
End Sub

Private Sub DDSArea_AfterUpdate()
    DDSArea = Trim(DDSArea)
End Sub

Private Sub Diagnosis_AfterUpdate()
    Call UpdateChangeLog("AutismDiagnosis", Diagnosis)
End Sub

Private Sub Form_Current()
    Call ProgressMessages("Append", "   Opening autism services form.")
    If Inactive Then
        Me.Caption = "Client: Autism Services (INACTIVE)"
    Else
        Me.Caption = "Client: Autism Services"
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If

    If CurrentAutismWaiverClient = False And FormerWaiverClient = False Then
        AutismWaiverEndDate = Null
        AutismWaiverStartDate = Null
    End If

    SupportBroker = SupportBrokerFirstName & " " & SupportBrokerLastName

    Age = Form_frmPeopleClientsDemographics.Age
    If Not IsNull(Age) And Age >= 9 And CurrentAutismWaiverClient = True Then
        CurrentAutismWaiverClient.Visible = False
        CurrentAutismWaiverClient = False
        FormerWaiverClient = True
        Form_frmPeople.ErrorMessages = Form_frmPeople.ErrorMessages & "[Sev Med] Individual is no longer autism waiver eligible." & vbCrLf & " " & vbCrLf
        Form_frmPeople.ErrorMessages.BackStyle = 1
    End If
    
    If (IsNull(DDSArea) Or Len(DDSArea) <= 0) And Form_frmPeople.PhysicalState = "MA" Then DDSArea = DLookup("Area", "catDDSAreasAndRegions", "CityTown='" & Form_frmPeople.PhysicalCity & "'")
    If Form_frmPeople.PhysicalState <> "MA" Then DDSArea = "*** Out of State *** "
End Sub

Private Sub FormerWaiverClient_Click()
    Dim TempDate As Date
    ' Calculate end date if individual is a former waiver client.
    If FormerWaiverClient Then
        TempDate = DateValue(Form_frmPeopleClientsDemographics.DateOfBirth) + DateAdd("yyyy", 9, TempDate)
        AutismWaiverEndDate = Format(TempDate, "mm/dd/yyyy")
    Else
        AutismWaiverEndDate = Null
    End If
End Sub

Private Sub Inactive_Click()
    If Inactive Then
        Call CheckInactive(DateInactive, "AutismSetInactive", "True")
        Call GreyAndNormal(Form_frmPeople.IsClientAutismLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices - 1
        If Form_frmPeople.NumberOfClientServices < 0 Then Form_frmPeople.NumberOfClientServices = 0
    Else
        Call CheckInactive(DateInactive, "AutismSetInactive", "False")
        Call BlueAndBold(Form_frmPeople.IsClientAutismLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
End Sub

Private Sub ReferralSource_AfterUpdate()
    Call UpdateChangeLog("AutismReferralSource", ReferralSource)
End Sub

Private Sub SelectBroker_Click()
    DoCmd.OpenForm "frmPeopleSelectPerson", , , , , , "AutismBroker"
End Sub