' Module Name: Form_frmPeopleClientsServiceDay
' Module Type: Document Module
' Lines of Code: 112
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Dim DontUpdateScores As Boolean

Private Sub Form_Current()
    Call ProgressMessages("Append", "   Opening day services form.")
    If IsNull(Form_frmPeopleClientsVendors.DayVendor) Then Call FillWithTILL("Day", CityTown & "-" & LocationName)
    MedicaidRate = DLookup("MedicaidRate", "catSeverityRates", "Severity = '" & Severity & "'")
    MedicaidRate.Requery
    If Inactive Then
        Me.Caption = "Client: Day Services (INACTIVE)"
    Else
        Me.Caption = "Client: Day Services"
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
    Annual = DMRAnnual + MedicaidAnnual
    DontUpdateScores = True
    Call TotalScore
    DontUpdateScores = False
    If Len(ContractNumber) > 0 Then ActivityCode = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber & """")
End Sub

Private Sub Form_GotFocus()
    If IsNull(CityTown) Then DoCmd.OpenForm "frmPeopleClientsSelectLocation", , , , , , "Day"
End Sub

Private Function TotalScore(Optional ChangeLogLabel As String, Optional ChangeLogValue As Variant) As Boolean
    Dim Criteria As Variant

    ScoreTotal = ScoreAggression + ScoreAuditory + ScoreCommunication + ScoreEating + ScoreHyperactivity + _
        ScoreLearning + ScoreMedical + ScoreMobility + ScoreNoncompliance + ScoreSelfInjury + ScoreSocial + _
        ScoreToileting + ScoreVision
    Profile = ScoreTotal + 13
    
    If Profile >= DLookup("LowScore", "catSeverityRates", "Severity='Low'") And Profile <= DLookup("HighScore", "catSeverityRates", "Severity='Low'") Then
        Severity = "Low"
    ElseIf Profile >= DLookup("LowScore", "catSeverityRates", "Severity='Mod'") And Profile <= DLookup("HighScore", "catSeverityRates", "Severity='Mod'") Then
        Severity = "Mod"
    ElseIf Profile >= DLookup("LowScore", "catSeverityRates", "Severity='High'") And Profile <= DLookup("HighScore", "catSeverityRates", "Severity='High'") Then
        Severity = "High"
    Else
        Severity = Null
    End If
    
    Criteria = "Severity = '" & Severity & "'"
    MedicaidRate = DLookup("MedicaidRate", "catSeverityRates", Criteria)
    MedicaidRate.Requery
    If Not IsMissing(ChangeLogLabel) Then Call UpdateChangeLog(ChangeLogLabel, ChangeLogValue)
    TotalScore = True
    ' Score updated.  Make note.
    If DontUpdateScores Then
    Else
        ScoresUpdatedWhen = Format(Now(), "mm/dd/yyyy")
        ScoresUpdatedWho = Form_frmMainMenu.UserName
    End If
End Function

Private Sub Inactive_Click()
    If Inactive Then
        Call CheckInactive(DateInactive, "DayServicesSetInactive", "True")
        TILLDataBase.Execute "UPDATE qrytblPeopleClientsVendors SET qrytblPeopleClientsVendors.ResidentialVendor = Null, qrytblPeopleClientsVendors.ResVendorAddress = Null, qrytblPeopleClientsVendors.ResVendorCity = Null, qrytblPeopleClientsVendors.ResVendorState = Null, qrytblPeopleClientsVendors.ResVendorZIP = Null, qrytblPeopleClientsVendors.ResidentialVendorPhoneNumber = Null, qrytblPeopleClientsVendors.ResVendorLocation = Null " & _
            "WHERE (((qrytblPeopleClientsVendors.IndexedName)='" & IndexedName & "'));", dbSeeChanges: Call BriefDelay
        Call GreyAndNormal(Form_frmPeople.IsClientDayLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices - 1
        If Form_frmPeople.NumberOfClientServices < 0 Then Form_frmPeople.NumberOfClientServices = 0
    Else
        Call CheckInactive(DateInactive, "DayServicesSetInactive", "False")
        Call BlueAndBold(Form_frmPeople.IsClientDayLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
End Sub

Private Sub LocationName_AfterUpdate()
    Call UpdateChangeLog("DayCityTown", CityTown)
    Call UpdateChangeLog("DayLocation", LocationName)
    Form_frmPeopleClientsVendors.DayVendorLocation = CityTown & "-" & LocationName
End Sub

Private Sub MassHealthServiceLevel_Change()
    LevelUpdatedWhen = Format(Date, "mm/dd/yyyy")
    LevelUpdatedWho = Form_frmMainMenu.UserName
End Sub

Private Sub SelectDay_Click()
    DoCmd.OpenForm "frmPeopleClientsSelectLocation", , , , , , "Day"
End Sub

Private Sub Severity_AfterUpdate()
    MedicaidRate = DLookup("MedicaidRate", "catSeverityRates", "Severity = '" & Severity & "'")
    MedicaidRate.Requery
    Call UpdateChangeLog("DaySeverity", Severity)
    Call UpdateChangeLog("DayMedicaidRate", MedicaidRate)
End Sub

Private Sub ContractNumber_AfterUpdate()
    If Len(ContractNumber) > 0 Then ActivityCode = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber & """")
    Call UpdateChangeLog("DayActivityCode1", ActivityCode)
End Sub

Private Sub ContractNumber2_AfterUpdate()
    If Len(ContractNumber2) > 0 Then ActivityCode2 = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber2 & """")
    Call UpdateChangeLog("DayActivityCode2", ActivityCode2)
End Sub

Private Sub StartDate_AfterUpdate()
    If ValidateDateString(StartDate) Then Call UpdateChangeLog("DayStartDate", StartDate)
End Sub

Private Sub EndDate_AfterUpdate()
    If ValidateDateString(EndDate) Then Call UpdateChangeLog("DayEndDate", EndDate)
End Sub
