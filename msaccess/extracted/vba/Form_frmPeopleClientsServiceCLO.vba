' Module Name: Form_frmPeopleClientsServiceCLO
' Module Type: Document Module
' Lines of Code: 221
' Extracted: 1/29/2026 4:12:26 PM

Option Compare Database
Option Explicit

Private Sub Form_Current()
    Call ProgressMessages("Append", "   Opening CLO form.")
    If IsNull(RecordAddedDate) Then
        RecordAddedDate = Format(Now(), "mm/dd/yyyy")
        RecordAddedBy = Form_frmMainMenu.UserName
        DoCmd.OpenForm "frmPeopleClientsSelectLocation", , , , , , "CLO"
    End If
    If IsNull(Form_frmPeopleClientsVendors.ResidentialVendor) Then Call FillWithTILL("Res", CityTown & "-" & Location)
    If Inactive Then
        Me.Caption = "Client: Creative Living Options (INACTIVE)"
    Else
        Me.Caption = "Client: Creative Living Options"
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
    If Len(ContractNumber) > 0 Then ActivityCode = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber & """")
    If Section8 Then
        WaitListSection8.Visible = True
        DateInspected.Visible = True
        PassFail.Visible = True
        Section8Review.Visible = True
        If WaitListSection8 Then
            WaitListSection8Date.Visible = True
            If IsNull(WaitListSection8Date) Then
                MsgBox "Section 8 Wait List is selected.  You must enter a Wait List date.", vbOKOnly
                WaitListSection8Date.SetFocus
            End If
            If Not IsDate(WaitListSection8Date) Then
                MsgBox "You must enter a valid date.", vbOKOnly
                WaitListSection8Date.SetFocus
            End If
        Else
            WaitListSection8Date.Visible = False
        End If
    Else
        WaitListSection8.Visible = False
        WaitListSection8Date.Visible = False
        DateInspected.Visible = False
        PassFail.Visible = False
        Section8Review.Visible = False
    End If
End Sub

Private Sub Form_GotFocus()
    If IsNull(CityTown) Then DoCmd.OpenForm "frmPeopleClientsSelectLocation", , , , , , "CLO"
End Sub

Private Sub ActivityCode_AfterUpdate()
    Call UpdateChangeLog("CLOActivityCode", ActivityCode)
End Sub

Private Sub ActivityCode2_AfterUpdate()
    Call UpdateChangeLog("CLOActivityCode2", ActivityCode2)
End Sub

Private Sub ClientContribution_AfterUpdate()
    Call UpdateChangeLog("CLOChargesForCare", ClientContribution)
End Sub

Private Sub ContractNumber_AfterUpdate()
    If Len(ContractNumber) > 0 Then ActivityCode = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber & """")
    Call UpdateChangeLog("CLOContractNumber", ContractNumber)
End Sub

Private Sub ContractNumber2_AfterUpdate()
    If Len(ContractNumber2) > 0 Then ActivityCode2 = DLookup("ActivityCode", "tblContracts", "ContractID = """ & ContractNumber2 & """")
    Call UpdateChangeLog("CLOContractNumber2", ContractNumber2)
End Sub

Private Sub EndDate_AfterUpdate()
    If ValidateDateString(EndDate) Then Call UpdateChangeLog("CLOEndDate", EndDate)
End Sub

Private Sub Inactive_Click()
    If Inactive Then
        Call CheckInactive(DateInactive, "CLOSetInactive", "True")
        TILLDataBase.Execute "UPDATE qrytblPeopleClientsVendors SET qrytblPeopleClientsVendors.ResidentialVendor = Null, qrytblPeopleClientsVendors.ResVendorAddress = Null, qrytblPeopleClientsVendors.ResVendorCity = Null, qrytblPeopleClientsVendors.ResVendorState = Null, qrytblPeopleClientsVendors.ResVendorZIP = Null, qrytblPeopleClientsVendors.ResidentialVendorPhoneNumber = Null, qrytblPeopleClientsVendors.ResVendorLocation = Null " & _
            "WHERE (((qrytblPeopleClientsVendors.IndexedName)='" & IndexedName & "'));", dbSeeChanges: Call BriefDelay
        Call GreyAndNormal(Form_frmPeople.IsCilentCLOLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices - 1
        If Form_frmPeople.NumberOfClientServices < 0 Then Form_frmPeople.NumberOfClientServices = 0
    Else
        Call CheckInactive(DateInactive, "CLOSetInactive", "False")
        Call BlueAndBold(Form_frmPeople.IsCilentCLOLabel)
        Form_frmPeople.NumberOfClientServices = Form_frmPeople.NumberOfClientServices + 1
    End If
End Sub

Private Sub Funding_AfterUpdate()
    Call UpdateChangeLog("CLOFundingSource", Funding)
End Sub

Private Sub HousingAuthorityAddress_AfterUpdate()
    Call UpdateChangeLog("CLOHousingAuthorityAddress", HousingAuthorityAddress)
End Sub

Private Sub HousingAuthorityCaseManager_AfterUpdate()
    Call UpdateChangeLog("CLOHousingAuthorityCaseManager", HousingAuthorityCaseManager)
End Sub

Private Sub HousingAuthorityCity_AfterUpdate()
    Call UpdateChangeLog("CLOHousingAuthorityCity", HousingAuthorityCity)
End Sub

Private Sub HousingAuthorityCurrentRent_AfterUpdate()
    Call UpdateChangeLog("CLOHousingAuthorityCurrentRent", HousingAuthorityCurrentRent)
End Sub

Private Sub HousingAuthorityFunds_AfterUpdate()
    Call UpdateChangeLog("CLOHousingAuthorityFunds", HousingAuthorityFunds)
End Sub

Private Sub HousingAuthorityOffice_AfterUpdate()
    Call UpdateChangeLog("CLOHousingAuthorityOffice", HousingAuthorityOffice)
End Sub

Private Sub HousingAuthorityPermissionLetter_AfterUpdate()
    Call UpdateChangeLog("CLOHousingAuthorityPermissionLetter", HousingAuthorityPermissionLetter)
End Sub

Private Sub HousingAuthorityPhone_AfterUpdate()
    Call UpdateChangeLog("CLOHousingAuthorityPhone", HousingAuthorityPhone)
End Sub

Private Sub HousingAuthorityState_AfterUpdate()
    Call UpdateChangeLog("CLOHousingAuthorityState", HousingAuthorityState)
End Sub

Private Sub HousingAuthorityWorkInProgress_AfterUpdate()
    Call UpdateChangeLog("CLOHousingAuthorityWorkInProgress", HousingAuthorityWorkInProgress)
End Sub

Private Sub HousingAuthorityZIP_AfterUpdate()
    Call UpdateChangeLog("CLOHousingAuthorityZIP", HousingAuthorityZIP)
End Sub

Private Sub LeaseBegins_AfterUpdate()
    If ValidateDateString(LeaseBegins) Then Call UpdateChangeLog("CLOLeaseBegins", LeaseBegins)
End Sub

Private Sub LeaseEnds_AfterUpdate()
    If ValidateDateString(LeaseEnds) Then Call UpdateChangeLog("CLOLeaseEnds", LeaseEnds)
End Sub

Private Sub LPNHoursAtResidence_AfterUpdate()
    Call UpdateChangeLog("CLOLPNHoursAtResidence", LPNHoursAtResidence)
End Sub

Private Sub Portion_AfterUpdate()
    Call UpdateChangeLog("CLOAllocation", Portion)
End Sub

Private Sub RecertificationMonth_AfterUpdate()
    Call UpdateChangeLog("CLORecertificationMonth", RecertificationMonth)
End Sub

Private Sub RNHoursAtResidence_AfterUpdate()
    Call UpdateChangeLog("CLORNHoursAtResidence", RNHoursAtResidence)
End Sub

Private Sub RoomAndBoard_AfterUpdate()
    Call UpdateChangeLog("CLORoomAndBoard", RoomAndBoard)
End Sub

Private Sub Section8_AfterUpdate()
    If Section8 Then
        WaitListSection8.Visible = True
        DateInspected.Visible = True
        PassFail.Visible = True
        Section8Review.Visible = True
        WaitListSection8Date.Visible = False
        Call UpdateChangeLog("CLOSection8", "True")
    Else
        WaitListSection8.Visible = False
        WaitListSection8 = False
        WaitListSection8Date.Visible = False
        WaitListSection8Date = Null
        DateInspected.Visible = False
        DateInspected = Null
        PassFail.Visible = False
        PassFail = False
        Section8Review.Visible = False
        Section8Review = Null
        Call UpdateChangeLog("CLOSection8", "False")
    End If
End Sub

Private Sub SelectResidence_Click()
    DoCmd.OpenForm "frmPeopleClientsSelectLocation", , , , , , "CLO"
End Sub

Private Sub StartDate_AfterUpdate()
    If ValidateDateString(StartDate) Then Call UpdateChangeLog("CLOStartDate", StartDate)
End Sub

Private Sub TerminationReason_AfterUpdate()
    Call UpdateChangeLog("CLOTerminationReason", TerminationReason)
End Sub

Private Sub WaitListSection8_Click()
    If WaitListSection8 Then
        WaitListSection8Date.Visible = True
        WaitListSection8Date.SetFocus
    Else
        WaitListSection8Date = Null
        WaitListSection8Date.Visible = False
    End If
End Sub

Private Sub WaitListSection8Date_AfterUpdate()
    If IsNull(WaitListSection8Date) Then
        MsgBox "Section 8 Wait List is selected.  You must enter a Wait List date.", vbOKOnly
        WaitListSection8Date.SetFocus
    ElseIf Not IsDate(WaitListSection8Date) Then
        MsgBox "Section 8 Wait List is selected.  You must enter a valid Wait List date.", vbOKOnly
        WaitListSection8Date.SetFocus
    End If
End Sub
