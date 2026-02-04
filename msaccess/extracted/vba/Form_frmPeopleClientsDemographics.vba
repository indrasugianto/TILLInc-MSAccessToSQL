' Module Name: Form_frmPeopleClientsDemographics
' Module Type: Document Module
' Lines of Code: 1520
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Dim DOB As Date, ErrMsgs As String, MissingFields As Boolean, AddFamilyOpenArgs As Variant, FirstErrMessage As Boolean

Private Sub Form_Load()
    ' Set user permissions.
    BlockFinancials.Visible = DCount("Action", "catUserPermissions", "User='" & Form_frmMainMenu.UserName & "' AND Action='Can access Demographics Financials'") = 0
    BlockEmploymentRepPayee.Visible = DCount("Action", "catUserPermissions", "User='" & Form_frmMainMenu.UserName & "' AND Action='Block Employment and Rep Payees'") > 0
    BlockFamily.Visible = DCount("Action", "catUserPermissions", "User='" & Form_frmMainMenu.UserName & "' AND Action='Cannot access Family'") > 0
    BlockExpirations.Visible = DCount("Action", "catUserPermissions", "User='" & Form_frmMainMenu.UserName & "' AND Action='Cannot access Expirations'") > 0
    AutismServices.Visible = False  ' Autism Services closed 11/10/21.
End Sub

Private Sub Form_Current()
    ' Close all open subforms, initialize all jump buttons, and initialize all service toggles.
    DoCmd.SetWarnings False
'   DoCmd.OpenForm "frmPleaseWaitBrief"
    Call ProgressMessages("Append", "   Opening demographics form.")
    SysCmdResult = SysCmd(4, "Initializing client subforms.")
    Me.Caption = "Client: Demographics"
    Call CloseClientServicesSubForms
    Call Initialize
    ' Open vendors page.
    SysCmdResult = SysCmd(4, "Opening Vendors form.")
    DoCmd.OpenForm "frmPeopleClientsVendors", , , "IndexedName=""" & IndexedName & """"
    Me.SetFocus
    JumpClient2.Visible = True
    ' Check required fields.
    Call CheckRequiredFields
    ' Open Service Forms.
    Call OpenServiceForms
    ' Continue initialization of demographics.
    SysCmdResult = SysCmd(4, "Continue demographics initialization.")
    If MotherStatus Then MotherDateOfDeath.Visible = True Else MotherDateOfDeath.Visible = False
    If FatherStatus Then FatherDateOfDeath.Visible = True Else FatherDateOfDeath.Visible = False
    If WaiverClient Then
        ResidentialWaiver.Visible = True
        CommunityLivingWaiver.Visible = True
        AdultSupportsWaiver.Visible = True
        If (Not BlockFinancials.Visible) And (Not Form_frmPeople.IsDeceased) And (Not ResidentialWaiver) And (Not CommunityLivingWaiver) And (Not AdultSupportsWaiver) Then
            Form_frmPeople.ErrorMessages = Form_frmPeople.ErrorMessages & "[Sev Med]   Waiver Client selected but Waiver Type not selected.  Select a Waiver Type or check the Waiver Client box off." & vbCrLf & " " & vbCrLf
            Form_frmPeople.ErrorMessages.BackStyle = 1
            WaiverClient.SetFocus
        End If
    Else
        ResidentialWaiver.Visible = False
        CommunityLivingWaiver.Visible = False
        AdultSupportsWaiver.Visible = False
    End If
    If FoodStampsEligible Then
        FoodStamps.Visible = True
        If FoodStamps Then
            FoodStampsCardNumber.Visible = True
            FoodStampsAmount.Visible = True
'           FoodStampsOffice.Visible = True
            SNAPAgencyID.Visible = True
            FoodStampsLastCertDate.Visible = True
            FoodStampsNextCertDate.Visible = True
        Else
            FoodStampsCardNumber.Visible = False
            FoodStampsAmount.Visible = False
'           FoodStampsOffice.Visible = False
            SNAPAgencyID.Visible = False
            FoodStampsLastCertDate.Visible = False
            FoodStampsNextCertDate.Visible = False
        End If
    Else
        FoodStamps.Visible = False
        FoodStampsCardNumber.Visible = False
        FoodStampsAmount.Visible = False
'       FoodStampsOffice.Visible = False
        SNAPAgencyID.Visible = False
        FoodStampsLastCertDate.Visible = False
        FoodStampsNextCertDate.Visible = False
    End If
    If DayServices Then ClientID.Visible = True Else ClientID.Visible = False
    If RepPayeeIsTILL Or RepPayeeIsClient Then BlockRepPayee.Visible = True Else BlockRepPayee.Visible = False
    If RepPayeeAddressPopulated And RepPayeeCityPopulated And RepPayeeStatePopulated And (Not RepPayeeAddressValidated) Then Call ValidateAddress
    ' Check Expiration Dates.
    Call ExpDatesAfterUpdate(1, True)
    Call ExpDatesAfterUpdate(2, True)
    Call ExpDatesAfterUpdate(3, True)
    Call ExpDatesAfterUpdate(4, True)
    Call ExpDatesAfterUpdate(5, True)
    Call ExpDatesAfterUpdate(7, True)
    Call ExpDatesAfterUpdate(8, True)
    ' Check legal name.  Pre-fill if necessary.
    If Len(LegalName) <= 0 Or IsNull(LegalName) Then If Len(MiddleInitial) > 0 Then LegalName = FirstName & " " & MiddleInitial & " " & LastName Else LegalName = FirstName & " " & LastName
    ' Fill county of residence.
    If Len(CountyOfResidence) <= 0 Or IsNull(CountyOfResidence) Then CountyOfResidence = DLookup("County", "catCounties", "State='" & Form_frmPeople.PhysicalState & "' AND CityTown='" & Form_frmPeople.PhysicalCity & "'")
    SysCmdResult = SysCmd(5)
    DoCmd.SetWarnings True
'   DoCmd.Close acForm, "frmPleaseWaitBrief"
End Sub

Private Sub OpenServiceForms()
    If DayServices Then Call OpenServiceForm_Day Else Call YellowAndNormal(DayLabel)
    If ResidentialServices Then Call OpenServiceForm_Res Else Call YellowAndNormal(ResidentialLabel)
    If TransportationServices Then Call OpenServiceForm_Trans Else Call YellowAndNormal(TransLabel)
    If CLO Then Call OpenServiceForm_CLO Else Call YellowAndNormal(CLOLabel)
    If VocationalServices Then Call OpenServiceForm_Voc Else Call YellowAndNormal(VocLabel)
    If SharedLiving Then Call OpenServiceForm_SharedLiving Else Call YellowAndNormal(SharedLivingLabel)
    If NHDay Then Call OpenServiceForm_NHDay Else Call YellowAndNormal(NHDayLabel)
    If NHRes Then Call OpenServiceForm_NHRes Else Call YellowAndNormal(NHResLabel)
    If IndivSupport Then Call OpenServiceForm_ISS Else Call YellowAndNormal(ISLabel)
    If AdultComp Then Call OpenServiceForm_AdultCompanion Else Call YellowAndNormal(ACompLabel)
    If AdultCoach Then Call OpenServiceForm_AdultCoaching Else Call YellowAndNormal(ACoachLabel)
'   If PCA Then Call OpenServiceForm_PCA Else Call YellowAndNormal(PCALabel)
    If SPRINGBOARD Then Call OpenServiceForm_Springboard Else Call YellowAndNormal(SpringboardLabel)
'   If CommunityConnections Then Call OpenServiceForm_CC Else Call YellowAndNormal(CommunityConnectionsLabel)
'   If TRASE Then Call OpenServiceForm_TRASE Else Call YellowAndNormal(TRASELabel)
End Sub

Private Sub CheckRequiredFields()
    ' Skip this if individual is deceased.
    If Form_frmPeople.IsDeceased Then Exit Sub
    ' Check Medicare status.
    If BlockFinancials.Visible = False Then
        If Age >= 65 And IsNull(MedicareNumber) Then
            If Not Form_frmPeople.IsDeceased Then
                Form_frmPeople.ErrorMessages = Form_frmPeople.ErrorMessages & "[Sev Med]   Client is 65 years of age or older.  Please obtain Medicare Number." & vbCrLf & " " & vbCrLf
                Form_frmPeople.ErrorMessages.BackStyle = 1
            End If
        End If
    End If
    ' Check Date of Birth.  Calculate if necessary.
    If Not Form_frmPeople.NewClient Then
        If IsNull(DateOfBirth) Then
            MissingFields = True
            ErrMsgs = ErrMsgs & "[Sev High] The following mandatory field is not entered: Date of Birth" & vbCrLf
        Else
            If IsDate(DateOfBirth) Then
                DOB = CDate(DateOfBirth)
                Age = CalcAge(DOB)
            End If
        End If
    End If
    ' Check gender.
    If Not Form_frmPeople.NewClient Then
        If IsNull(Gender) Then
            MissingFields = True
            ErrMsgs = ErrMsgs & "[Sev High] The following mandatory field is not entered: Gender" & vbCrLf
        End If
    End If
    ' Check SSN.  Not required for PCA, Springboard, or Autism clients.
    If Not Form_frmPeople.NewClient Then
        If (IsNull(SocialSecurityNumber) Or Len(SocialSecurityNumber) <> 11) And Not PCA_Autism_Or_CommunityConnections_Springboard_TRASE Then
            MissingFields = True
            ErrMsgs = ErrMsgs & "[Sev High] The following mandatory field is not entered: Social Security Number" & vbCrLf
        End If
    End If
    ' Check Date of Death.
    If Not Form_frmPeople.NewClient Then
        If Form_frmPeople.IsDeceased And IsNull(Form_frmPeople.DeceasedDate) Or Len(Form_frmPeople.DeceasedDate) <> 10 Then
            MissingFields = True
            ErrMsgs = ErrMsgs & "[Sev High] The following mandatory field is not entered: Date of Death" & vbCrLf
        End If
    End If
    ' Set legal status if not otherwise set.
    If IsNull(LegalStatus) Then If Age < 18 Then LegalStatus = "Minor" Else LegalStatus = "Presumed Competent"
    If Age < 18 Then LegalStatus = "Minor"
    If LegalStatus = "Minor" And Age >= 18 Then LegalStatus = "Presumed Competent"
    ' Display missing field message.
    If MissingFields Then
        If Form_frmPeople.ErrorMessages = "" Or IsNull(Form_frmPeople.ErrorMessages) Then
            Form_frmPeople.ErrorMessages = ErrMsgs & vbCrLf & " " & vbCrLf
        Else
            Form_frmPeople.ErrorMessages = Form_frmPeople.ErrorMessages & ErrMsgs & vbCrLf & " " & vbCrLf
        End If
        Form_frmPeople.ErrorMessages.BackStyle = 1
    End If
End Sub

Private Sub OpenServiceForm(FieldLabel As Label, FormName As String, JumpButton As CommandButton)
    Call BlueAndBold(FieldLabel)
    DoCmd.OpenForm FormName, , , "IndexedName=""" & IndexedName & """"
    JumpButton.Visible = True
End Sub

Private Sub OpenServiceForm_Day()
    If DCount("IndexedName", "tblPeopleClientsDayServices", "IndexedName = """ & IndexedName & """") <= 0 Then Call DayServices_Click
    SysCmdResult = SysCmd(4, "Opening Day Services form.")
    Call OpenServiceForm(DayLabel, "frmPeopleClientsServiceDay", JumpDay)
    If Form_frmPeopleClientsServiceDay.Inactive Then DayLabel.ForeColor = RGB(165, 165, 165)
    Form_frmPeopleClientsVendors.DayVendorLocation = Form_frmPeopleClientsServiceDay.CityTown & "-" & Form_frmPeopleClientsServiceDay.LocationName
End Sub

Private Sub OpenServiceForm_Res()
    If DCount("IndexedName", "tblPeopleClientsResidentialServices", "IndexedName = """ & IndexedName & """") <= 0 Then Call ResidentialServices_Click
    SysCmdResult = SysCmd(4, "Opening Residential Services form.")
    Call OpenServiceForm(ResidentialLabel, "frmPeopleClientsServiceResidential", JumpResidential)
    If Form_frmPeopleClientsServiceResidential.Inactive Then ResidentialLabel.ForeColor = RGB(165, 165, 165)
    Form_frmPeopleClientsVendors.ResVendorLocation = Form_frmPeopleClientsServiceResidential.CityTown & "-" & Form_frmPeopleClientsServiceResidential.Location
    Form_frmPeopleClientsServiceResidential.ClientContribution = 0
    If Not IsNull(SSAMonthlyAmount) Then Form_frmPeopleClientsServiceResidential.ClientContribution = Form_frmPeopleClientsServiceResidential.ClientContribution + SSAMonthlyAmount
    If Not IsNull(SSIMonthlyAmount) Then Form_frmPeopleClientsServiceResidential.ClientContribution = Form_frmPeopleClientsServiceResidential.ClientContribution + SSIMonthlyAmount
    If Not IsNull(SSPMonthlyAmount) Then Form_frmPeopleClientsServiceResidential.ClientContribution = Form_frmPeopleClientsServiceResidential.ClientContribution + SSPMonthlyAmount
    If Not IsNull(Pension) Then Form_frmPeopleClientsServiceResidential.ClientContribution = Form_frmPeopleClientsServiceResidential.ClientContribution + Pension
    Form_frmPeopleClientsServiceResidential.ClientContribution = Form_frmPeopleClientsServiceResidential.ClientContribution * 0.75
    Form_frmPeopleClientsServiceResidential.HousingAuthorityCurrentRent = Form_frmPeopleClientsServiceResidential.ClientContribution
End Sub

Private Sub OpenServiceForm_Trans()
    If DCount("IndexedName", "tblPeopleClientsTransportationServices", "IndexedName = """ & IndexedName & """") <= 0 Then Call TransportationServices_Click
    SysCmdResult = SysCmd(4, "Opening Transportation Services form.")
    Call OpenServiceForm(TransLabel, "frmPeopleClientsServiceTransportation", JumpTrans)
    If Form_frmPeopleClientsServiceTransportation.Inactive Then TransLabel.ForeColor = RGB(165, 165, 165)
End Sub

Private Sub OpenServiceForm_CLO()
    If DCount("IndexedName", "tblPeopleClientsCLOServices", "IndexedName = """ & IndexedName & """") <= 0 Then Call CLO_Click
    SysCmdResult = SysCmd(4, "Opening CLO Services form.")
    Call OpenServiceForm(CLOLabel, "frmPeopleClientsServiceCLO", JumpCLO)
    If Form_frmPeopleClientsServiceCLO.Inactive Then CLOLabel.ForeColor = RGB(165, 165, 165)
    Form_frmPeopleClientsVendors.ResVendorLocation = Form_frmPeopleClientsServiceCLO.CityTown & "-" & Form_frmPeopleClientsServiceCLO.Location
    Form_frmPeopleClientsServiceCLO.ClientContribution = 0
    If Not IsNull(SSAMonthlyAmount) Then Form_frmPeopleClientsServiceCLO.ClientContribution = Form_frmPeopleClientsServiceCLO.ClientContribution + SSAMonthlyAmount
    If Not IsNull(SSIMonthlyAmount) Then Form_frmPeopleClientsServiceCLO.ClientContribution = Form_frmPeopleClientsServiceCLO.ClientContribution + SSIMonthlyAmount
    If Not IsNull(SSPMonthlyAmount) Then Form_frmPeopleClientsServiceCLO.ClientContribution = Form_frmPeopleClientsServiceCLO.ClientContribution + SSPMonthlyAmount
    If Not IsNull(Pension) Then Form_frmPeopleClientsServiceCLO.ClientContribution = Form_frmPeopleClientsServiceCLO.ClientContribution + Pension
    Form_frmPeopleClientsServiceCLO.ClientContribution = Form_frmPeopleClientsServiceCLO.ClientContribution * 0.75
    Form_frmPeopleClientsServiceCLO.HousingAuthorityCurrentRent = Form_frmPeopleClientsServiceCLO.ClientContribution
End Sub

Private Sub OpenServiceForm_Voc()
    If DCount("IndexedName", "tblPeopleClientsVocationalServices", "IndexedName = """ & IndexedName & """") <= 0 Then Call VocationalServices_Click
    SysCmdResult = SysCmd(4, "Opening Vocational Services form.")
    Call OpenServiceForm(VocLabel, "frmPeopleClientsServiceVocational", JumpVoc)
    If Form_frmPeopleClientsServiceVocational.Inactive Then VocLabel.ForeColor = RGB(165, 165, 165)
    Form_frmPeopleClientsVendors.DayVendorLocation = Form_frmPeopleClientsServiceVocational.CityTown & "-" & Form_frmPeopleClientsServiceVocational.Location
    If IsNull(Employer) Or Len(Employer) = 0 Then Call FillWithTILL("Emp", "")
End Sub

Private Sub OpenServiceForm_SharedLiving()
    If DCount("IndexedName", "tblPeopleClientsSharedLivingServices", "IndexedName = """ & IndexedName & """") <= 0 Then Call SharedLiving_Click
    SysCmdResult = SysCmd(4, "Opening Shared Living Services form.")
    Call OpenServiceForm(SharedLivingLabel, "frmPeopleClientsServiceSharedLiving", JumpSharedLiving)
    If Form_frmPeopleClientsServiceSharedLiving.Inactive Then SharedLivingLabel.ForeColor = RGB(165, 165, 165)
End Sub

Private Sub OpenServiceForm_NHDay()
    If DCount("IndexedName", "tblPeopleClientsNHDay", "IndexedName = """ & IndexedName & """") <= 0 Then Call NHDay_Click
    SysCmdResult = SysCmd(4, "Opening NH Day Services form.")
    Call OpenServiceForm(NHDayLabel, "frmPeopleClientsServiceNHDay", JumpNHDay)
    If Form_frmPeopleClientsServiceNHDay.Inactive Then NHDayLabel.ForeColor = RGB(165, 165, 165)
End Sub

Private Sub OpenServiceForm_NHRes()
    If DCount("IndexedName", "tblPeopleClientsNHRes", "IndexedName = """ & IndexedName & """") <= 0 Then Call NHRes_Click
    SysCmdResult = SysCmd(4, "Opening NH Residential Services form.")
    Call OpenServiceForm(NHResLabel, "frmPeopleClientsServiceNHRes", JumpNHRes)
    If Form_frmPeopleClientsServiceNHRes.Inactive Then NHResLabel.ForeColor = RGB(165, 165, 165)
End Sub

Private Sub OpenServiceForm_ISS()
    If DCount("IndexedName", "tblPeopleClientsIndividualSupportServices", "IndexedName = """ & IndexedName & """") <= 0 Then Call IndivSupport_Click
    SysCmdResult = SysCmd(4, "Opening Individual Support Services form.")
    Call OpenServiceForm(ISLabel, "frmPeopleClientsServiceIndividualSupport", JumpISS)
    If Form_frmPeopleClientsServiceIndividualSupport.Inactive Then ISLabel.ForeColor = RGB(165, 165, 165)
End Sub

Private Sub OpenServiceForm_AdultCompanion()
    If DCount("IndexedName", "tblPeopleClientsAdultCompanion", "IndexedName = """ & IndexedName & """") <= 0 Then Call AdultComp_Click
    SysCmdResult = SysCmd(4, "Opening Adult Companion services form.")
    Call OpenServiceForm(ACompLabel, "frmPeopleClientsServiceAdultCompanion", JumpAComp)
    If Form_frmPeopleClientsServiceAdultCompanion.Inactive Then ACompLabel.ForeColor = RGB(165, 165, 165)
End Sub

Private Sub OpenServiceForm_AdultCoaching()
    If DCount("IndexedName", "tblPeopleClientsAdultCoaching", "IndexedName = """ & IndexedName & """") <= 0 Then Call AdultCoach_Click
    SysCmdResult = SysCmd(4, "Opening Adult Coaching services form.")
    Call OpenServiceForm(ACoachLabel, "frmPeopleClientsServiceAdultCoaching", JumpACoach)
    If Form_frmPeopleClientsServiceAdultCoaching.Inactive Then ACoachLabel.ForeColor = RGB(165, 165, 165)
End Sub

Private Sub OpenServiceForm_Autism()
    If DCount("IndexedName", "tblPeopleClientsAutismServices", "IndexedName = """ & IndexedName & """") <= 0 Then Call AutismServices_Click
    SysCmdResult = SysCmd(4, "Opening Autism Services form.")
    Call OpenServiceForm(AutismLabel, "frmPeopleClientsServiceAutism", JumpAutism)
    If Form_frmPeopleClientsServiceAutism.Inactive Then AutismLabel.ForeColor = RGB(165, 165, 165)
End Sub

Private Sub OpenServiceForm_PCA()
    If DCount("IndexedName", "tblPeopleClientsPCAServices", "IndexedName = """ & IndexedName & """") <= 0 Then Call PCA_Click
    SysCmdResult = SysCmd(4, "Opening PCA Services form.")
    Call OpenServiceForm(PCALabel, "frmPeopleClientsServicePCA", JumpPCA)
    If Form_frmPeopleClientsServicePCA.Inactive Then PCALabel.ForeColor = RGB(165, 165, 165)
End Sub

Private Sub OpenServiceForm_Springboard()
    If DCount("IndexedName", "tblPeopleClientsSpringboardServices", "IndexedName = """ & IndexedName & """") <= 0 Then Call SPRINGBOARD_Click
    SysCmdResult = SysCmd(4, "Opening Springboard form.")
    Call OpenServiceForm(SpringboardLabel, "frmPeopleClientsServiceSpringboard", JumpSpringboard)
    If Form_frmPeopleClientsServiceSpringboard.Inactive Then SpringboardLabel.ForeColor = RGB(165, 165, 165)
End Sub

Private Sub OpenServiceForm_CC()
    If DCount("IndexedName", "tblPeopleClientsCommunityConnectionsServices", "IndexedName = """ & IndexedName & """") <= 0 Then Call CommunityConnections_Click
    SysCmdResult = SysCmd(4, "Opening Community Connections form.")
    Call OpenServiceForm(CommunityConnectionsLabel, "frmPeopleClientsServiceCommunityConnections", JumpCC)
    If Form_frmPeopleClientsServiceCommunityConnections.Inactive Then CommunityConnectionsLabel.ForeColor = RGB(165, 165, 165)
End Sub

Private Sub OpenServiceForm_TRASE()
    If DCount("IndexedName", "tblPeopleClientsTRASEServices", "IndexedName = """ & IndexedName & """") <= 0 Then Call TRASE_Click
    SysCmdResult = SysCmd(4, "Opening TRASE form.")
    Call OpenServiceForm(TRASELabel, "frmPeopleClientsServiceTRASE", JumpTRASE)
    If Form_frmPeopleClientsServiceTRASE.Inactive Then TRASELabel.ForeColor = RGB(165, 165, 165)
End Sub

Private Sub Initialize()
    JumpPeople.Visible = True:          JumpClient2.Visible = True
    JumpAutism.Visible = False:         JumpCLO.Visible = False:                 JumpCC.Visible = False:       JumpDay.Visible = False
    JumpISS.Visible = False:            JumpAComp.Visible = False:               JumpACoach.Visible = False:   JumpPCA.Visible = False
    JumpResidential.Visible = False:    JumpSharedLiving.Visible = False:        JumpNHDay.Visible = False:    JumpNHRes.Visible = False
    JumpSpringboard.Visible = False:    JumpTrans.Visible = False:               JumpTRASE.Visible = False:    JumpVoc.Visible = False
    
    With Form_frmPeople
        .IsClientAutism = AutismServices
        .IsCilentCLO = CLO
        .IsClientCommunityConnections = CommunityConnections
        .IsClientDay = DayServices
        .IsClientIndiv = IndivSupport
        .IsClientNHDay = NHDay
        .IsClientNHRes = NHRes
        .IsClientPCA = PCA
        .IsClientRes = ResidentialServices
        .IsClientSharedLiving = SharedLiving
        .IsClientSpring = SPRINGBOARD
        .IsClientTrans = TransportationServices
        .IsClientTRASE = TRASE
        .IsClientVocat = VocationalServices
    End With
    
    MissingFields = False
    ErrMsgs = ""
    FirstErrMessage = True
    AddFamilyOpenArgs = IndexedName & "," & LastName & "," & MiddleInitial & "," & LastName
End Sub

Private Sub BCStatus_AfterUpdate()
    If BCStatus Then Call UpdateChangeLog("BCStatus", "True") Else Call UpdateChangeLog("BCStatus", "False")
End Sub

Private Function ValidateExpirationDate(DateVal As Variant, ChangeLog As String, WhenCurrent As Boolean) As Boolean
    If IsDate(DateVal) Then
        Call UpdateChangeLog(ChangeLog, DateVal)
        ValidateExpirationDate = True
    Else
        If IsNull(DateVal) Then
            ValidateExpirationDate = True
            Exit Function
        End If
        If Not WhenCurrent Then
            MsgBox "The value you entered is not recognized as a legitimate date.", vbOKOnly, "ERROR!"
        Else
            MsgBox "The " & ChangeLog & " field value is not recognized as a legitimate date.", vbOKOnly, "ERROR!"
        End If
        DateVal = Null
        ValidateExpirationDate = False
    End If
End Function

Private Function ExpDatesAfterUpdate(WhichField As Integer, WhenCurrent As Boolean) As Boolean
    ExpDatesAfterUpdate = True
    Select Case WhichField
        Case 1
            If Not ValidateExpirationDate(DateISP, "ISP Date", WhenCurrent) Then
                DateISP.ForeColor = RGB(255, 0, 0)
                DateISP.SetFocus
                DateISP = RememberPreviousDate
            Else
                DateISP.ForeColor = RGB(0, 0, 0)
            End If
        Case 2
            If Not ValidateExpirationDate(DateConsentFormsSigned, "Date Consent Forms Signed", WhenCurrent) Then
                DateConsentFormsSigned.ForeColor = RGB(255, 0, 0)
                DateConsentFormsSigned.SetFocus
                DateConsentFormsSigned = RememberPreviousDate
            Else
                DateConsentFormsSigned.ForeColor = RGB(0, 0, 0)
            End If
        Case 3
            If Not ValidateExpirationDate(DateBMMExpires, "Date BMM Expires", WhenCurrent) Then
                DateBMMExpires.ForeColor = RGB(255, 0, 0)
                DateBMMExpires.SetFocus
                DateBMMExpires = RememberPreviousDate
            Else
                DateBMMExpires.ForeColor = RGB(0, 0, 0)
            End If
        Case 4
            If Not ValidateExpirationDate(DateBMMAccessSigned, "Date BMM Access Signed", WhenCurrent) Then
                DateBMMAccessSigned.ForeColor = RGB(255, 0, 0)
                DateBMMAccessSigned.SetFocus
                DateBMMAccessSigned = RememberPreviousDate
            Else
                DateBMMAccessSigned.ForeColor = RGB(0, 0, 0)
            End If
        Case 5
            If Not ValidateExpirationDate(DateBMMAccessSignedHRC, "Date BMM Access Signed HRC", WhenCurrent) Then
                DateBMMAccessSignedHRC.ForeColor = RGB(255, 0, 0)
                DateBMMAccessSignedHRC.SetFocus
                DateBMMAccessSignedHRC = RememberPreviousDate
            Else
                DateBMMAccessSignedHRC.ForeColor = RGB(0, 0, 0)
            End If
        Case 7
            If Not ValidateExpirationDate(DateSignaturesDueBy, "Date Signatures Due By", WhenCurrent) Then
                DateSignaturesDueBy.ForeColor = RGB(255, 0, 0)
                DateSignaturesDueBy.SetFocus
                DateSignaturesDueBy = RememberPreviousDate
            Else
                DateSignaturesDueBy.ForeColor = RGB(0, 0, 0)
            End If
        Case 8
            If Not ValidateExpirationDate(DateSPDAuthExpires, "Date SPD Auth Expires", WhenCurrent) Then
                DateSPDAuthExpires.ForeColor = RGB(255, 0, 0)
                DateSPDAuthExpires.SetFocus
                DateSPDAuthExpires = RememberPreviousDate
            Else
                DateSPDAuthExpires.ForeColor = RGB(0, 0, 0)
            End If
    End Select
End Function

Private Sub DateOfBirth_AfterUpdate()
    If IsDate(DateOfBirth) Then
        Call UpdateChangeLog("DateOfBirth", DateOfBirth)
        Age = CalcAge(CDate(DateOfBirth))
    Else
        MsgBox "The value you entered is not recognized as a legitimate date.", vbOKOnly, "ERROR!"
        DateOfBirth = Null
        DateOfBirth.SetFocus
    End If
End Sub

Private Sub ServiceClick(PeopleToggle As Boolean, ServiceToggle As Boolean, AddQuery As String, ServiceForm As String, Dept As String, InactiveFlag As Boolean)
    PeopleToggle = ServiceToggle
    If ServiceToggle Then
        Call BriefDelay
        TILLDataBase.Execute "INSERT INTO " & AddQuery & " ( IndexedName, RecordAddedDate, RecordAddedBy, Inactive ) " & _
            "SELECT """ & Left(Form_frmPeople.IndexedName, 160) & """ AS IndexedName, Now() AS RecordAddedDate, """ & Form_frmMainMenu.UserName & """ AS RecordAddedBy, False AS Inactive;", dbSeeChanges: Call BriefDelay
        DoCmd.OpenForm ServiceForm, , , "IndexedName=""" & Left(Form_frmPeople.IndexedName, 160) & """"
        Form_frmPeople.DeptCriteria = Dept
    Else
        Call BriefDelay
        DoCmd.Close acForm, ServiceForm
        Form_frmPeople.DeptCriteria = ""
    End If
    Form_frmPeople.SetFocus
End Sub

Private Sub AutismServices_Click()
    If AutismServices Then
        If MsgBox("You are about to add this individual as an Autism Services Client.  Proceed?", vbYesNo, "Verify") = vbNo Then
            AutismServices = False
            Exit Sub
        End If
    Else
        If MsgBox("You are about to remove this individual as an Autism Services Client.  Proceed?" & vbCrLf & "NOTE: You may want to make this individual inactive instead of removing from Autism Services.", vbYesNo, "Verify") = vbNo Then AutismServices = True
    End If

    Call ServiceClick(Form_frmPeople.IsClientAutism, AutismServices, "tblPeopleClientsAutismServices", "frmPeopleClientsServiceAutism", "Autism Services", Form_frmPeopleClientsServiceAutism.Inactive)
    If AutismServices Then
        Call BlueAndBold(AutismLabel)
        Call UpdateChangeLog("IsClientAutism", "True")
    Else
        Call YellowAndNormal(AutismLabel)
        Call UpdateChangeLog("IsClientAutism", "False")
    End If
    
    Form_frmPeople.IsClientAutism = AutismServices

    If MsgBox("Would you like to add family member(s)/contact(s)?", vbYesNo, "Add family?") = vbYes Then
        AddFamilyMemberToClient = True
        DoCmd.OpenForm "frmPeopleFamilyEnterAndValidatePerson", , , , , , AddFamilyOpenArgs
    Else
        AddFamilyMemberToClient = False
    End If
End Sub

Private Sub CLO_Click()
    If CLO Then
        If MsgBox("You are about to add this individual as a CLO Client.  Proceed?", vbYesNo, "Verify") = vbNo Then
            CLO = False
            Exit Sub
        End If
    Else
        If MsgBox("You are about to remove this individual as a CLO Client.  Proceed?" & vbCrLf & "NOTE: You may want to make this individual inactive instead of removing from CLO Services.", vbYesNo, "Verify") = vbNo Then
            CLO = True
        Else
            With Form_frmPeopleClientsVendors
                .ResidentialVendor = ""
                .ResVendorAddress = ""
                .ResVendorCity = ""
                .ResVendorState = ""
                .ResVendorZIP = ""
                .ResidentialVendorPhoneNumber = ""
                .ResVendorLocation = ""
                .ResVendorLocation.Visible = False
'               .Dirty = False
            End With
        End If
    End If
    
    Call ServiceClick(Form_frmPeople.IsClientCLO, CLO, "tblPeopleClientsCLOServices", "frmPeopleClientsServiceCLO", "Creative Living Options", Form_frmPeopleClientsServiceCLO.Inactive)
    
    If CLO Then
        DoCmd.OpenForm "frmPeopleClientsSelectLocation", , , , , , "CLO"
        Call FillWithTILL("Res", Form_frmPeopleClientsServiceCLO.CityTown & "-" & Form_frmPeopleClientsServiceCLO.Location)
    End If
    
    If CLO Then
        Call BlueAndBold(CLOLabel)
        Call UpdateChangeLog("IsClientCLO", "True")
    Else
        Call YellowAndNormal(CLOLabel)
        Call UpdateChangeLog("IsClientCLO", "False")
    End If
    
    Form_frmPeople.IsCilentCLO = CLO

    If MsgBox("Would you like to add family member(s)/contact(s)?", vbYesNo, "Add family?") = vbYes Then
        AddFamilyMemberToClient = True
        DoCmd.OpenForm "frmPeopleFamilyEnterAndValidatePerson", , , , , , AddFamilyOpenArgs
    Else
        AddFamilyMemberToClient = False
    End If
End Sub

Private Sub CommunityConnections_Click()
    If CommunityConnections Then
        If MsgBox("You are about to add this individual as a Recreation Client.  Proceed?", vbYesNo, "Verify") = vbNo Then
            CommunityConnections = False
            Exit Sub
        End If
    Else
        If MsgBox("You are about to remove this individual as a Recreation Client.  Proceed" & vbCrLf & "NOTE: You may want to make this individual inactive instead of removing from Recreation Services.?", vbYesNo, "Verify") = vbNo Then CommunityConnections = True
    End If

    Call ServiceClick(Form_frmPeople.IsClientCommunityConnections, CommunityConnections, "tblPeopleClientsCommunityConnectionsServices", "frmPeopleClientsServiceCommunityConnections", "Community Connections", Form_frmPeopleClientsServiceCommunityConnections.Inactive)
    If CommunityConnections Then
        Call BlueAndBold(CommunityConnectionsLabel)
        Call UpdateChangeLog("IsClientCommunityConnections", "True")
    Else
        Call YellowAndNormal(CommunityConnectionsLabel)
        Call UpdateChangeLog("IsClientCommunityConnections", "False")
    End If
    
    Form_frmPeople.IsClientCommunityConnections = CommunityConnections

    If MsgBox("Would you like to add family member(s)/contact(s)?", vbYesNo, "Add family?") = vbYes Then
        AddFamilyMemberToClient = True
        DoCmd.OpenForm "frmPeopleFamilyEnterAndValidatePerson", , , , , , AddFamilyOpenArgs
    Else
        AddFamilyMemberToClient = False
    End If
End Sub

Private Sub DayServices_Click()
    If DayServices Then
        If MsgBox("You are about to add this individual as a Day Services Client.  Proceed?", vbYesNo, "Verify") = vbNo Then
            DayServices = False
            ClientID.Visible = True
            Exit Sub
        End If
    Else
        If MsgBox("You are about to remove this individual as a Day Services Client.  Proceed?" & vbCrLf & "NOTE: You may want to make this individual inactive instead of removing from Day Services.", vbYesNo, "Verify") = vbNo Then
            DayServices = True
            ClientID.Visible = True
        Else
            ClientID.Visible = False
            With Form_frmPeopleClientsVendors
                .DayVendor = ""
                .DayVendorAddress = ""
                .DayVendorCity = ""
                .DayVendorState = ""
                .DayVendorZIP = ""
                .DayVendorPhoneNumber = ""
                .DayVendorLocation = ""
                .DayVendorLocation.Visible = False
            End With
        End If
    End If

    Call ServiceClick(Form_frmPeople.IsClientDay, DayServices, "tblPeopleClientsDayServices", "frmPeopleClientsServiceDay", "Day Habilitation", Form_frmPeopleClientsServiceDay.Inactive)
    
    If DayServices Then
        DoCmd.OpenForm "frmPeopleClientsSelectLocation", , , , , , "Day"
        Call FillWithTILL("Day", Form_frmPeopleClientsServiceDay.CityTown & "-" & Form_frmPeopleClientsServiceDay.LocationName)
    End If
    
    If DayServices Then
        Call BlueAndBold(DayLabel)
        Call UpdateChangeLog("IsClientDayServices", "True")
    Else
        Call YellowAndNormal(DayLabel)
        Call UpdateChangeLog("IsClientDayServices", "False")
    End If
    
    Form_frmPeople.IsClientDay = DayServices

    If MsgBox("Would you like to add family member(s)/contact(s)?", vbYesNo, "Add family?") = vbYes Then
        AddFamilyMemberToClient = True
        DoCmd.OpenForm "frmPeopleFamilyEnterAndValidatePerson", , , , , , AddFamilyOpenArgs
    Else
        AddFamilyMemberToClient = False
    End If
End Sub

Private Sub Gender_AfterUpdate()
    Call UpdateChangeLog("Gender", Gender)
    If (Not IsNull(Form_frmPeople.Salutation)) Or Len(Form_frmPeople.Salutation) > 0 Then _
        If Gender = "M" Then _
            Form_frmPeople.Salutation = "Mr" _
        Else _
            If Gender = "F" Then _
                Form_frmPeople.Salutation = "Ms"
End Sub

Private Sub GuardianshipPapersOnFile_AfterUpdate()
    If GuardianshipPapersOnFile Then Call UpdateChangeLog("GuardianshipPapersOnFile", "True") Else Call UpdateChangeLog("GuardianshipPapersOnFile", "False")
End Sub

Private Sub Hearing_Impaired_AfterUpdate()
    If HearingImpaired Then Call UpdateChangeLog("HearingImpaired", "True") Else Call UpdateChangeLog("HearingImpaired", "False")
End Sub

Private Sub VisuallyImpaired_AfterUpdate()
    If VisuallyImpaired Then Call UpdateChangeLog("VisuallyImpaired", "True") Else Call UpdateChangeLog("VisuallyImpaired", "False")
End Sub

Private Sub IndivSupport_Click()
    If IndivSupport Then
        If MsgBox("You are about to add this individual as an Individual Support Client.  Proceed?", vbYesNo, "Verify") = vbNo Then
            IndivSupport = False
            Exit Sub
        End If
    Else
        If MsgBox("You are about to remove this individual as an Individual Support Client.  Proceed?" & vbCrLf & "NOTE: You may want to make this individual inactive instead of removing from Individual Support Services.", vbYesNo, "Verify") = vbNo Then IndivSupport = True
    End If

    Call ServiceClick(Form_frmPeople.IsClientIndiv, IndivSupport, "tblPeopleClientsIndividualSupportServices", "frmPeopleClientsServiceIndividualSupport", "IndivSupport", Form_frmPeopleClientsServiceIndividualSupport.Inactive)
    If IndivSupport Then
        Call BlueAndBold(ISLabel)
        Call UpdateChangeLog("IsIndivSupport", "True")
        
    Else
        Call YellowAndNormal(ISLabel)
        Call UpdateChangeLog("IsIndivSupport", "False")
    End If
    
    Form_frmPeople.IsClientIndiv = IndivSupport

    If MsgBox("Would you like to add family member(s)/contact(s)?", vbYesNo, "Add family?") = vbYes Then
        AddFamilyMemberToClient = True
        DoCmd.OpenForm "frmPeopleFamilyEnterAndValidatePerson", , , , , , AddFamilyOpenArgs
    Else
        AddFamilyMemberToClient = False
    End If
End Sub

Private Sub AdultComp_Click()
    If AdultComp Then
        If MsgBox("You are about to add this individual as an Adult Companion client.  Proceed?", vbYesNo, "Verify") = vbNo Then
            AdultComp = False
            Exit Sub
        End If
    Else
        If MsgBox("You are about to remove this individual as an Adult Companion client.  Proceed?" & vbCrLf & "NOTE: You may want to make this individual inactive instead of removing from Individual Support Services.", vbYesNo, "Verify") = vbNo Then AdultComp = True
    End If
    
    Form_frmPeople.IsClientAdultComp = AdultComp
    If AdultComp Then
        TILLDataBase.Execute "INSERT INTO tblPeopleClientsAdultCompanion ( IndexedName, RecordAddedDate, RecordAddedBy, Inactive ) " & _
            "SELECT """ & Left(Form_frmPeople.IndexedName, 160) & """ AS IndexedName, Now() AS RecordAddedDate, """ & Form_frmMainMenu.UserName & """ AS RecordAddedBy, False AS Inactive;", dbSeeChanges: Call BriefDelay
        DoCmd.OpenForm "frmPeopleClientsServiceAdultCompanion", , , "IndexedName=""" & IndexedName & """"
        Form_frmPeople.DeptCriteria = "AdultComp"
    Else
        DoCmd.Close acForm, "frmPeopleClientsServiceAdultCompanion"
        Form_frmPeople.DeptCriteria = ""
    End If
    
    Form_frmPeople.SetFocus
    
    If AdultComp Then
        Call BlueAndBold(ACompLabel)
        Call UpdateChangeLog("IsAdultCompanion", "True")
    Else
        Call YellowAndNormal(ACompLabel)
        Call UpdateChangeLog("IsAdultCompanion", "False")
    End If
    
    Form_frmPeople.IsClientAdultComp = AdultComp

    If MsgBox("Would you like to add family member(s)/contact(s)?", vbYesNo, "Add family?") = vbYes Then
        AddFamilyMemberToClient = True
        DoCmd.OpenForm "frmPeopleFamilyEnterAndValidatePerson", , , , , , AddFamilyOpenArgs
    Else
        AddFamilyMemberToClient = False
    End If
End Sub

Private Sub AdultCoach_Click()
    If AdultCoach Then
        If MsgBox("You are about to add this individual as an Adult Coaching client.  Proceed?", vbYesNo, "Verify") = vbNo Then
            AdultCoach = False
            Exit Sub
        End If
    Else
        If MsgBox("You are about to remove this individual as an Adult Coaching client.  Proceed?" & vbCrLf & "NOTE: You may want to make this individual inactive instead of removing from Individual Support Services.", vbYesNo, "Verify") = vbNo Then AdultCoach = True
    End If
    
    Form_frmPeople.IsClientAdultCoach = AdultCoach
    If AdultCoach Then
        TILLDataBase.Execute "INSERT INTO tblPeopleClientsAdultCoaching ( IndexedName, RecordAddedDate, RecordAddedBy, Inactive ) " & _
            "SELECT """ & Left(Form_frmPeople.IndexedName, 160) & """ AS IndexedName, Now() AS RecordAddedDate, """ & Form_frmMainMenu.UserName & """ AS RecordAddedBy, False AS Inactive;", dbSeeChanges: Call BriefDelay
        DoCmd.OpenForm "frmPeopleClientsServiceAdultCoaching", , , "IndexedName=""" & IndexedName & """"
        Form_frmPeople.DeptCriteria = "AdultCoach"
    Else
        DoCmd.Close acForm, "frmPeopleClientsServiceAdultCoaching"
        Form_frmPeople.DeptCriteria = ""
    End If
    
    Form_frmPeople.SetFocus
    
    If AdultCoach Then
        Call BlueAndBold(ACoachLabel)
        Call UpdateChangeLog("IsAdultCoaching", "True")
    Else
        Call YellowAndNormal(ACoachLabel)
        Call UpdateChangeLog("IsAdultCoaching", "False")
    End If
    
    Form_frmPeople.IsClientAdultCoach = AdultCoach

    If MsgBox("Would you like to add family member(s)/contact(s)?", vbYesNo, "Add family?") = vbYes Then
        AddFamilyMemberToClient = True
        DoCmd.OpenForm "frmPeopleFamilyEnterAndValidatePerson", , , , , , AddFamilyOpenArgs
    Else
        AddFamilyMemberToClient = False
    End If
End Sub

Private Sub MotherStatus_AfterUpdate()
    If MotherStatus Then Call UpdateChangeLog("MotherStatus", "True") Else Call UpdateChangeLog("MotherStatus", "False")
End Sub

Private Sub FatherStatus_AfterUpdate()
    If FatherStatus Then Call UpdateChangeLog("FatherStatus", "True") Else Call UpdateChangeLog("FatherStatus", "False")
End Sub

Private Sub MotherStatus_Click()
    MotherDateOfDeath.Visible = True
End Sub

Private Sub FatherStatus_Click()
    FatherDateOfDeath.Visible = True
End Sub

Private Sub NHDay_Click()
    If NHDay Then
        If MsgBox("You are about to add this individual as a NH Day Client.  Proceed?", vbYesNo, "Verify") = vbNo Then
            NHDay = False
            Exit Sub
        End If
    Else
        If MsgBox("You are about to remove this individual as a NH Day Client.  Proceed?" & vbCrLf & "NOTE: You may want to make this individual inactive instead of removing from NH Day Services.", vbYesNo, "Verify") = vbNo Then NHDay = True
    End If

    Call ServiceClick(Form_frmPeople.IsClientNHDay, NHDay, "tblPeopleClientsNHDay", "frmPeopleClientsServiceNHDay", "NHDay", Form_frmPeopleClientsServiceNHDay.Inactive)
    If NHDay Then
        Call BlueAndBold(NHDayLabel)
        Call UpdateChangeLog("IsClientNHDay", "True")
    Else
        Call YellowAndNormal(NHDayLabel)
        Call UpdateChangeLog("IsClientNHDay", "False")
    End If
    
    Form_frmPeople.IsClientNHDay = NHDay

    If MsgBox("Would you like to add family member(s)/contact(s)?", vbYesNo, "Add family?") = vbYes Then
        AddFamilyMemberToClient = True
        DoCmd.OpenForm "frmPeopleFamilyEnterAndValidatePerson", , , , , , AddFamilyOpenArgs
    Else
        AddFamilyMemberToClient = False
    End If
End Sub

Private Sub NHRes_Click()
    If NHRes Then
        If MsgBox("You are about to add this individual as a NH Residential Client.  Proceed?", vbYesNo, "Verify") = vbNo Then
            NHRes = False
            Exit Sub
        End If
    Else
        If MsgBox("You are about to remove this individual as a NH Residential Client.  Proceed?" & vbCrLf & "NOTE: You may want to make this individual inactive instead of removing from NH Residential Services.", vbYesNo, "Verify") = vbNo Then NHRes = True
    End If

    Call ServiceClick(Form_frmPeople.IsClientNHRes, NHRes, "tblPeopleClientsNHRes", "frmPeopleClientsServiceNHRes", "NHRes", Form_frmPeopleClientsServiceNHRes.Inactive)
    If NHRes Then
        Call BlueAndBold(NHResLabel)
        Call UpdateChangeLog("IsClientNHRes", "True")
    Else
        Call YellowAndNormal(NHResLabel)
        Call UpdateChangeLog("IsClientNHRes", "False")
    End If
    
    Form_frmPeople.IsClientNHRes = NHRes

    If MsgBox("Would you like to add family member(s)/contact(s)?", vbYesNo, "Add family?") = vbYes Then
        AddFamilyMemberToClient = True
        DoCmd.OpenForm "frmPeopleFamilyEnterAndValidatePerson", , , , , , AddFamilyOpenArgs
    Else
        AddFamilyMemberToClient = False
    End If
End Sub

Private Sub PCA_Click()
    If PCA Then
        If MsgBox("You are about to add this individual as a PCA Client.  Proceed?", vbYesNo, "Verify") = vbNo Then
            PCA = False
            Exit Sub
        End If
    Else
        If MsgBox("You are about to remove this individual as a PCA Client.  Proceed?" & vbCrLf & "NOTE: You may want to make this individual inactive instead of removing from PCA Services.", vbYesNo, "Verify") = vbNo Then PCA = True
    End If

    Call ServiceClick(Form_frmPeople.IsClientPCA, PCA, "tblPeopleClientsPCAServices", "frmPeopleClientsServicePCA", "PCA", Form_frmPeopleClientsServicePCA.Inactive)
    If PCA Then
        Call BlueAndBold(PCALabel)
        Call UpdateChangeLog("IsClientPCA", "True")
    Else
        Call YellowAndNormal(PCALabel)
        Call UpdateChangeLog("IsClientPCA", "False")
    End If
    
    Form_frmPeople.IsClientPCA = PCA

    If MsgBox("Would you like to add family member(s)/contact(s)?", vbYesNo, "Add family?") = vbYes Then
        AddFamilyMemberToClient = True
        DoCmd.OpenForm "frmPeopleFamilyEnterAndValidatePerson", , , , , , AddFamilyOpenArgs
    Else
        AddFamilyMemberToClient = False
    End If
End Sub

Private Sub RepPayeeIsClient_Click()
    BlockRepPayee.Visible = RepPayeeIsTILL Or RepPayeeIsClient
    
    If RepPayeeIsClient Then
        RepPayeeIsTILL = False
        RepresentativePayee = [Forms]![frmPeople]![FirstName] & " " & [Forms]![frmPeople]![MiddleInitial] & " " & [Forms]![frmPeople]![LastName]
        RepPayeeAddress = [Forms]![frmPeople]![MailingAddress]
        RepPayeeCity = [Forms]![frmPeople]![MailingCity]
        RepPayeeState = [Forms]![frmPeople]![MailingState]
        RepPayeeZIP = [Forms]![frmPeople]![MailingZIP]
        RepPayeePhone = [Forms]![frmPeople]![HomePhone]
        RepPayeeAddressValidated = [Forms]![frmPeople]![MailingAddressValidated]
        BlockRepPayee.Visible = True
    Else
        RepPayeeIsTILL = False
        RepresentativePayee = ""
        RepPayeeAddress = ""
        RepPayeeCity = ""
        RepPayeeState = ""
        RepPayeeZIP = ""
        RepPayeePhone = ""
        RepPayeeAddressValidated = False
        MsgBox "Representative Payee information has been cleared.  Be sure to enter Representative Payee Information", vbOKOnly, "Note"
        BlockRepPayee.Visible = False
    End If
End Sub

Private Sub RepPayeeIsTILL_Click()
    BlockRepPayee.Visible = RepPayeeIsTILL Or RepPayeeIsClient
    
    If RepPayeeIsTILL Then
        RepPayeeIsClient = False
        RepresentativePayee = "TILL, Inc."
        RepPayeeAddress = "20 Eastbrook Rd Ste 201"
        RepPayeeCity = "Dedham"
        RepPayeeState = "MA"
        RepPayeeZIP = "02026-2087"
        RepPayeePhone = "(781) 302-4600"
        RepPayeeAddressValidated = True
        BlockRepPayee.Visible = True
    Else
        RepPayeeIsClient = False
        RepresentativePayee = ""
        RepPayeeAddress = ""
        RepPayeeCity = ""
        RepPayeeState = ""
        RepPayeeZIP = ""
        RepPayeePhone = ""
        RepPayeeAddressValidated = False
        MsgBox "Representative Payee information has been cleared.  Be sure to enter Representative Payee Information", vbOKOnly, "Note"
        BlockRepPayee.Visible = False
    End If
End Sub

Private Sub ResidentialServices_Click()
    If ResidentialServices Then
        If MsgBox("You are about to add this individual as a Residential Services Client.  Proceed?", vbYesNo, "Verify") = vbNo Then
            ResidentialServices = False
            Exit Sub
        End If
    Else
        If MsgBox("You are about to remove this individual as a Residential Services Client.  Proceed?" & vbCrLf & "NOTE: You may want to make this individual inactive instead of removing from Residential Services.", vbYesNo, "Verify") = vbNo Then
            ResidentialServices = True
        Else
            With Form_frmPeopleClientsVendors
                .ResidentialVendor = ""
                .ResVendorAddress = ""
                .ResVendorCity = ""
                .ResVendorState = ""
                .ResVendorZIP = ""
                .ResidentialVendorPhoneNumber = ""
                .ResVendorLocation = ""
                .ResVendorLocation.Visible = False
            End With
        End If
    End If

    Call ServiceClick(Form_frmPeople.IsClientRes, ResidentialServices, "tblPeopleClientsResidentialServices", "frmPeopleClientsServiceResidential", "Residential Services", Form_frmPeopleClientsServiceResidential.Inactive)
    
    If ResidentialServices Then
        DoCmd.OpenForm "frmPeopleClientsSelectLocation", , , , , , "Res"
        Call FillWithTILL("Res", Form_frmPeopleClientsServiceResidential.CityTown & "-" & Form_frmPeopleClientsServiceResidential.Location)
    End If
    
    If ResidentialServices Then
        Call BlueAndBold(ResidentialLabel)
        Call UpdateChangeLog("IsResidentialServices", "True")
    Else
        Call YellowAndNormal(ResidentialLabel)
        Call UpdateChangeLog("IsResidentialServices", "False")
    End If
    
    Form_frmPeople.IsClientRes = ResidentialServices

    If MsgBox("Would you like to add family member(s)/contact(s)?", vbYesNo, "Add family?") = vbYes Then
        AddFamilyMemberToClient = True
        DoCmd.OpenForm "frmPeopleFamilyEnterAndValidatePerson", , , , , , AddFamilyOpenArgs
    Else
        AddFamilyMemberToClient = False
    End If
End Sub

Private Sub SharedLiving_Click()
    If SharedLiving Then
        If MsgBox("You are about to add this individual as a Shared Living Client.  Proceed?", vbYesNo, "Verify") = vbNo Then
            SharedLiving = False
            Exit Sub
        End If
    Else
        If MsgBox("You are about to remove this individual as a Shared Living Client.  Proceed?" & vbCrLf & "NOTE:You may want to make this individual inactive instead of removing from Shared Living Services.", vbYesNo, "Verify") = vbNo Then SharedLiving = True
    End If

    Call ServiceClick(Form_frmPeople.IsClientSharedLiving, SharedLiving, "tblPeopleClientsSharedLivingServices", "frmPeopleClientsServiceSharedLiving", "Shared Living", Form_frmPeopleClientsServiceSharedLiving.Inactive)
    If SharedLiving Then
        Call BlueAndBold(SharedLivingLabel)
        Call UpdateChangeLog("IsClientSharedLiving", "True")
    Else
        Call YellowAndNormal(SharedLivingLabel)
        Call UpdateChangeLog("IsClientSharedLiving", "False")
    End If
    
    Form_frmPeople.IsClientSharedLiving = SharedLiving

    If MsgBox("Would you like to add family member(s)/contact(s)?", vbYesNo, "Add family?") = vbYes Then
        AddFamilyMemberToClient = True
        DoCmd.OpenForm "frmPeopleFamilyEnterAndValidatePerson", , , , , , AddFamilyOpenArgs
    Else
        AddFamilyMemberToClient = False
    End If
End Sub

Private Sub SPRINGBOARD_Click()
    Dim Subject As String, TextBody As String
    
    If SPRINGBOARD Then
        If MsgBox("You are about to add this individual as a Springboard Client.  Proceed?", vbYesNo, "Verify") = vbNo Then
            SPRINGBOARD = False
            Exit Sub
        End If
    Else
        If MsgBox("You are about to remove this individual as a Springboard Client.  Proceed?" & vbCrLf & "NOTE: You may want to make this individual inactive instead of removing from Springboard.", vbYesNo, "Verify") = vbNo Then SPRINGBOARD = True
    End If

    Call ServiceClick(Form_frmPeople.IsClientSpring, SPRINGBOARD, "tblPeopleClientsSpringboardServices", "frmPeopleClientsServiceSpringboard", "Springboard", Form_frmPeopleClientsServiceSpringboard.Inactive)
    If SPRINGBOARD Then
        Call BlueAndBold(SpringboardLabel)
        Call UpdateChangeLog("IsClientSPRINGBOARD", "True")
        Subject = "TILLDB: Springboard Member Added to the Database"
        TextBody = Form_frmPeople.FirstName & " " & Form_frmPeople.LastName & " was added to the database by " & Form_frmMainMenu.UserName & "." & vbCrLf & "    " & Form_frmPeople.PhysicalAddress & vbCrLf & "    " & Form_frmPeople.PhysicalCity & " " & Form_frmPeople.PhysicalState & " " & Form_frmPeople.PhysicalZIP & vbCrLf
        Call SendEmailMessage("tilldbnotifications@tillinc.org", DLookup("ParameterValue", "appParameters", "ID=22"), "db.springboard@tillinc.org", Null, Subject, TextBody, Null)
    Else
        Call YellowAndNormal(SpringboardLabel)
        Call UpdateChangeLog("IsClientSPRINGBOARD", "False")
        Subject = "TILLDB: Springboard Member Deleted from the Database"
        TextBody = Form_frmPeople.FirstName & " " & Form_frmPeople.LastName & " was deleted from the database by " & Form_frmMainMenu.UserName & "." & vbCrLf & "    " & Form_frmPeople.PhysicalAddress & vbCrLf & "    " & Form_frmPeople.PhysicalCity & " " & Form_frmPeople.PhysicalState & " " & Form_frmPeople.PhysicalZIP & vbCrLf
        Call SendEmailMessage("tilldbnotifications@tillinc.org", DLookup("ParameterValue", "appParameters", "ID=22"), "db.springboard@tillinc.org", Null, Subject, TextBody, Null)
    End If
    
    Form_frmPeople.IsClientSpring = SPRINGBOARD

    If MsgBox("Would you like to add family member(s)/contact(s)?", vbYesNo, "Add family?") = vbYes Then
        AddFamilyMemberToClient = True
        DoCmd.OpenForm "frmPeopleFamilyEnterAndValidatePerson", , , , , , AddFamilyOpenArgs
    Else
        AddFamilyMemberToClient = False
    End If
End Sub

Private Sub TransportationServices_Click()
    If TransportationServices Then
        If MsgBox("You are about to add this individual as a Transportation Services Client.  Proceed?", vbYesNo, "Verify") = vbNo Then
            TransportationServices = False
            Exit Sub
        End If
    Else
        If MsgBox("You are about to remove this individual as a Transportation Services Client.  Proceed?" & vbCrLf & "NOTE: You may want to make this individual inactive instead of removing from Transportation Services.", vbYesNo, "Verify") = vbNo Then TransportationServices = True
    End If

    Call ServiceClick(Form_frmPeople.IsClientTrans, TransportationServices, "tblPeopleClientsTransportationServices", "frmPeopleClientsServiceTransportation", "Transportation Services", Form_frmPeopleClientsServiceTransportation.Inactive)
    If TransportationServices Then
        Call BlueAndBold(TransLabel)
        Call UpdateChangeLog("IsClientTransportationServices", "True")
    Else
        Call YellowAndNormal(TransLabel)
        Call UpdateChangeLog("IsClientTransportationServices", "False")
    End If
    
    Form_frmPeople.IsClientTrans = TransportationServices

    If MsgBox("Would you like to add family member(s)/contact(s)?", vbYesNo, "Add family?") = vbYes Then
        AddFamilyMemberToClient = True
        DoCmd.OpenForm "frmPeopleFamilyEnterAndValidatePerson", , , , , , AddFamilyOpenArgs
    Else
        AddFamilyMemberToClient = False
    End If
End Sub

Private Sub TRASE_Click()
    If TRASE Then
        If MsgBox("You are about to add this individual as a TRASE Client.  Proceed?", vbYesNo, "Verify") = vbNo Then
            TRASE = False
            Exit Sub
        End If
    Else
        If MsgBox("You are about to remove this individual as a TRASE Client.  Proceed?" & vbCrLf & "NOTE: You may want to make this individual inactive instead of removing from TRASE.", vbYesNo, "Verify") = vbNo Then TRASE = True
    End If

    Call ServiceClick(Form_frmPeople.IsClientTRASE, TRASE, "tblPeopleClientsTRASEServices", "frmPeopleClientsServiceTRASE", "TRASE", Form_frmPeopleClientsServiceTRASE.Inactive)
    If TRASE Then
        Call BlueAndBold(TRASELabel)
        Call UpdateChangeLog("IsClientTRASE", "True")
    Else
        Call YellowAndNormal(TRASELabel)
        Call UpdateChangeLog("IsClientTRASE", "False")
    End If
    
    Form_frmPeople.IsClientTRASE = TRASE

    If MsgBox("Would you like to add family member(s)/contact(s)?", vbYesNo, "Add family?") = vbYes Then
        AddFamilyMemberToClient = True
        DoCmd.OpenForm "frmPeopleFamilyEnterAndValidatePerson", , , , , , AddFamilyOpenArgs
    Else
        AddFamilyMemberToClient = False
    End If
End Sub

Private Sub VocationalServices_Click()
    If VocationalServices Then
        If MsgBox("You are about to add this individual as a Vocational Services Client.  Proceed?", vbYesNo, "Verify") = vbNo Then
            VocationalServices = False
            Employer = ""
            EmployerAddress = ""
            EmployerCity = ""
            EmployerState = ""
            EmployerZIP = ""
            EmployerPhone = ""
            EmployerSupervisorName = ""
            Exit Sub
        End If
    Else
        If MsgBox("You are about to remove this individual as a Vocational Services Client.  Proceed?" & vbCrLf & "NOTE: You may want to make this individual inactive instead of removing from Vocational Services.", vbYesNo, "Verify") = vbNo Then
            VocationalServices = True
        Else
            Employer = ""
            EmployerAddress = ""
            EmployerCity = ""
            EmployerState = ""
            EmployerZIP = ""
            EmployerPhone = ""
            EmployerSupervisorName = ""
            Form_frmPeopleClientsServiceVocational.Inactive = True
            Call CheckClientCompletelyInactive
            DoCmd.Close acForm, "frmPeopleClientsServiceVocational"
            Exit Sub
        End If
    End If
    
    Call ServiceClick(Form_frmPeople.IsClientVocat, VocationalServices, "tblPeopleClientsVocationalServices", "frmPeopleClientsServiceVocational", "Vocational Services", Form_frmPeopleClientsServiceVocational.Inactive)
    Call FillWithTILL("Voc", "")
    If VocationalServices Then
        Call BlueAndBold(VocLabel)
        Call UpdateChangeLog("IsClientVocationalServices", "True")
    Else
        Call YellowAndNormal(VocLabel)
        Call UpdateChangeLog("IsClientVocationalServices", "False")
    End If
    
    Form_frmPeople.IsClientVocat = VocationalServices

    If MsgBox("Would you like to add family member(s)/contact(s)?", vbYesNo, "Add family?") = vbYes Then
        AddFamilyMemberToClient = True
        DoCmd.OpenForm "frmPeopleFamilyEnterAndValidatePerson", , , , , , AddFamilyOpenArgs
    Else
        AddFamilyMemberToClient = False
    End If
End Sub

Private Sub FieldGotFocus(Optional DataField As Variant, Optional Border As Variant, Optional FieldLabel As Label)
    If Not (IsMissing(DataField)) Then DataField.BackColor = RGB(249, 205, 170)
    Border.BorderColor = RGB(255, 0, 0)
    FieldLabel.ForeColor = RGB(255, 255, 0)
End Sub

Private Sub DateBMMAccessSigned_GotFocus()
    RememberPreviousDate = DateBMMAccessSigned
    Call Highlight([DateBMMAccessSigned], True)
End Sub

Private Sub DateBMMAccessSignedHRC_GotFocus()
    RememberPreviousDate = DateBMMAccessSignedHRC
    Call Highlight([DateBMMAccessSignedHRC], True)
End Sub

Private Sub DateBMMExpires_GotFocus()
    RememberPreviousDate = DateBMMExpires
    Call Highlight([DateBMMExpires], True)
End Sub

Private Sub DateConsentFormsSigned_GotFocus()
    RememberPreviousDate = DateConsentFormsSigned
    Call Highlight([DateConsentFormsSigned], True)
End Sub

Private Sub DateISP_GotFocus()
    RememberPreviousDate = DateISP
    Call Highlight(DateISP, True)
End Sub

Private Sub DateSignaturesDueBy_GotFocus()
    RememberPreviousDate = DateSignaturesDueBy
    Call Highlight([DateSignaturesDueBy], True)
End Sub

Private Sub DateSPDAuthExpires_GotFocus()
    RememberPreviousDate = DateSPDAuthExpires
    Call Highlight([DateSPDAuthExpires], True)
End Sub

Private Sub DateOfBirth_LostFocus()
    If IsNull(DateOfBirth) Or Len(DateOfBirth) < 10 Then
        MsgBox "Date of birth is a mandatory field.  Please provide a valid date of birth.", vbOKOnly, "Error!"
        DateOfBirth.SetFocus
    End If
    Call Highlight(DateOfBirth, False)
End Sub

Private Sub Gender_LostFocus()
    If IsNull(Gender) Or Len(Gender) <> 1 Then
        MsgBox "Gender is a mandatory field.  Please provide a valid gender.", vbOKOnly, "Error!"
        Gender.SetFocus
        Exit Sub
    End If
    Call Highlight(Gender, False)
End Sub

Private Sub LegalStatus_LostFocus()
    If IsNull(LegalStatus) Or Len(LegalStatus) = 0 Then
        MsgBox "Legal status is a mandatory field.  Please provide a valid legal status.", vbOKOnly, "Error!"
        LegalStatus.SetFocus
        Exit Sub
    End If
    Call Highlight(LegalStatus, False)
End Sub

' AfterUpdate actions: Log field changes into the Change Log.

Private Sub AdultSupportsWaiver_AfterUpdate()
    If AdultSupportsWaiver Then Call UpdateChangeLog("AdultSupportsWaiver", "True") Else Call UpdateChangeLog("AdultSupportsWaiver", "False")
End Sub

Private Sub BankRoutingNumber_AfterUpdate()
    If Len(BankRoutingNumber) <> 9 Then
        MsgBox "Bank routing number must be nine digits in length.  Be sure to include all leading zeroes.  Please re-enter", , "ERROR"
        BankRoutingNumber = Null
        Me.Refresh
    Else
        Call UpdateChangeLog("BankRoutingNumber", BankRoutingNumber)
    End If
End Sub

Private Sub CommunityLivingWaiver_AfterUpdate()
    If CommunityLivingWaiver Then Call UpdateChangeLog("CommunityLivingWaiver", "True") Else Call UpdateChangeLog("CommunityLivingWaiver", "False")
End Sub

Private Sub DualEligible_AfterUpdate()
    If DualEligible Then Call UpdateChangeLog("DualEligible", "True") Else Call UpdateChangeLog("DualEligible", "False")
End Sub

Private Sub FoodStampsLastCertDate_AfterUpdate()
    If IsDate(FoodStampsLastCertDate) Then
        Call UpdateChangeLog("FoodStampsLastCertDate", FoodStampsLastCertDate)
    Else
        MsgBox "Value is not identified as a legitimate date.", vbOKOnly, "ERROR!"
        FoodStampsLastCertDate = Null
        FoodStampsLastCertDate.SetFocus
    End If
End Sub

Private Sub FoodStampsNextCertDate_AfterUpdate()
    If IsDate(FoodStampsNextCertDate) Then
        Call UpdateChangeLog("FoodStampsNextCertDate", FoodStampsNextCertDate)
    Else
        MsgBox "Value is not identified as a legitimate date.", vbOKOnly, "ERROR!"
        FoodStampsNextCertDate = Null
        FoodStampsNextCertDate.SetFocus
    End If
End Sub

Private Sub MedicaidNumber_AfterUpdate()
    Call UpdateChangeLog("MedicaidNumber", MedicaidNumber)
    If Not IsNull(MedicaidNumber) And Not IsNull(MedicareNumber) Then
        DualEligible = True
        Call UpdateChangeLog("DualEligible", "True")
    End If
End Sub

Private Sub MedicareNumber_AfterUpdate()
    Call UpdateChangeLog("MedicareNumber", MedicareNumber)
    If Not IsNull(MedicaidNumber) And Not IsNull(MedicareNumber) Then
        DualEligible = True
        Call UpdateChangeLog("DualEligible", "True")
    End If
End Sub

Private Sub ResidentialWaiver_AfterUpdate()
    If ResidentialWaiver Then Call UpdateChangeLog("ResidentialWaiver", "True") Else Call UpdateChangeLog("ResidentialWaiver", "False")
End Sub

Private Sub SocialSecurityNumber_AfterUpdate()
    Dim SSNCount As Integer, ELastName As String, EFirstName As String, EIndexedName As String
    ' Make sure this SSN is not being used.
    If Not IsNull(SocialSecurityNumber) Or Len(SocialSecurityNumber) > 0 Then
        SSNCount = DCount("IndexedName", "tblPeopleClientsDemographics", "SocialSecurityNumber=""" & SocialSecurityNumber & """")
        If SSNCount > 0 Then ' It's being used somewhere.
            EIndexedName = DLookup("IndexedName", "tblPeopleClientsDemographics", "SocialSecurityNumber=""" & SocialSecurityNumber & """")
            ELastName = DLookup("LastName", "tblPeople", "IndexedName=""" & EIndexedName & """")
            EFirstName = DLookup("FirstName", "tblPeople", "IndexedName=""" & EIndexedName & """")
            MsgBox "The SSN you entered belongs to " & EFirstName & " " & ELastName & ".  Please correct this.", vbOKOnly, "ERROR!"
            SocialSecurityNumber = ""
            SocialSecurityNumber.SetFocus
'           If Me.Dirty Then Me.Dirty = False
        Else
            Call UpdateChangeLog("SocialSecurityNumber", SocialSecurityNumber)
        End If
    End If
End Sub

Private Sub SSCardStatus_AfterUpdate()
    If SSCardStatus Then Call UpdateChangeLog("SSCardStatus", "True") Else Call UpdateChangeLog("SSCardStatus", "False")
End Sub

Private Sub Bank_AfterUpdate()
    Bank = CorrectProperNames(StrConv(Bank, vbProperCase))
    Call UpdateChangeLog("Bank", Bank)
End Sub

Private Sub BankAddress_AfterUpdate()
    BankAddress = CorrectProperNames(StrConv(BankAddress, vbProperCase))
    Call UpdateChangeLog("BankAddress", BankAddress)
End Sub

Private Sub BankCity_AfterUpdate()
    BankCity = CorrectProperNames(StrConv(BankCity, vbProperCase))
    Call UpdateChangeLog("BankCity", BankCity)
End Sub

Private Sub OtherInsurance_AfterUpdate()
    OtherInsurance = CorrectProperNames(StrConv(OtherInsurance, vbProperCase))
    Call UpdateChangeLog("OtherInsurance", OtherInsurance)
End Sub

Private Sub RepresentativePayee_AfterUpdate()
    If RepresentativePayee = "TILL" Then
        Call FillWithTILL("RPY", "")
        RepPayeeAddressValidated = True
    End If
    
    If RepresentativePayee = "TILL-NH" Or RepresentativePayee = "TILL NH" Then
        RepPayeeAddress = "154 Broad St Ste 1521"
        RepPayeeCity = "Nashua"
        RepPayeeState = "NH"
        RepPayeeZIP = "03063-3205"
        RepPayeeAddressValidated = True
    End If
    
    If RepresentativePayee = "Self" Then
        RepresentativePayee = FullName
        RepPayeeAddress = Form_frmPeople.MailingAddress
        RepPayeeCity = Form_frmPeople.MailingCity
        RepPayeeState = Form_frmPeople.MailingState
        RepPayeeZIP = Form_frmPeople.MailingZIP
        RepPayeeAddressValidated = Form_frmPeople.MailingAddressValidated
    End If
    
    Call UpdateChangeLog("RepresentativePayee", RepresentativePayee)
End Sub

Private Sub RepPayeeAddress_AfterUpdate()
    RepPayeeAddress = CorrectProperNames(StrConv(RepPayeeAddress, vbProperCase))
    If (Not IsNull(RepPayeeAddress)) Or Len(RepPayeeAddress) > 0 Then RepPayeeAddressPopulated = True
    Call UpdateChangeLog("RepPayeeAddress", RepPayeeAddress)
End Sub

Private Sub RepPayeeCity_AfterUpdate()
    RepPayeeCity = CorrectProperNames(StrConv(RepPayeeCity, vbProperCase))
    If (Not IsNull(RepPayeeCity)) Or Len(RepPayeeCity) > 0 Then RepPayeeCityPopulated = True
    Call UpdateChangeLog("RepPayeeCity", RepPayeeCity)
End Sub

Private Sub RepPayeeState_AfterUpdate()
    RepPayeeState = StrConv(RepPayeeState, vbUpperCase)
    If (Not IsNull(RepPayeeState)) Or Len(RepPayeeState) > 0 Then RepPayeeStatePopulated = True
    Call UpdateChangeLog("RepPayeeState", RepPayeeState)
End Sub

Private Sub RepPayeeZIP_AfterUpdate()
    RepPayeeZIP = CorrectProperNames(StrConv(RepPayeeZIP, vbProperCase))
    If (Not IsNull(RepPayeeZIP)) Or Len(RepPayeeZIP) > 0 Then RepPayeeZIPPopulated = True
    Call UpdateChangeLog("RepPayeeZIP", RepPayeeZIP)
End Sub

Private Sub WaiverClient_AfterUpdate()
    If WaiverClient Then Call UpdateChangeLog("WaiverClient", "True") Else Call UpdateChangeLog("WaiverClient", "False")
End Sub

' Click actions: Generally will set a flag and/or change some value then log the change to the Change log.

Private Sub FoodStampsEligible_Click()
    If FoodStampsEligible Then FoodStamps.Visible = True Else FoodStamps.Visible = False
    Call FoodStamps_Click
End Sub

Private Sub FoodStamps_Click()
    If FoodStamps Then
        FoodStampsCardNumber.Visible = True
        FoodStampsAmount.Visible = True
'       FoodStampsOffice.Visible = True
        SNAPAgencyID.Visible = True
        FoodStampsLastCertDate.Visible = True
        FoodStampsNextCertDate.Visible = True
        Call UpdateChangeLog("FoodStamps", "True")
    Else
        FoodStampsCardNumber.Visible = False
        FoodStampsAmount.Visible = False
        FoodStampsOffice.Visible = False
        SNAPAgencyID.Visible = False
        FoodStampsLastCertDate.Visible = False
        FoodStampsNextCertDate.Visible = False
        FoodStampsCardNumber = Null
        FoodStampsAmount = Null
        FoodStampsOffice = Null
        SNAPAgencyID = Null
        FoodStampsLastCertDate = Null
        FoodStampsNextCertDate = Null
        Call UpdateChangeLog("FoodStamps", "False")
    End If
End Sub

Private Sub GetFamily_Click()
    DoCmd.OpenForm "frmPeopleSelectPerson", , , , , , "ClientSelectFamily"
End Sub

Private Sub WaiverClient_Click()
    If WaiverClient Then
        ResidentialWaiver.Visible = True
        CommunityLivingWaiver.Visible = True
        AdultSupportsWaiver.Visible = True
        If Form_frmPeopleClientsDemographics.AutismServices Then Form_frmPeopleClientsServiceAutism.FormerWaiverClient = False
    Else
        ResidentialWaiver.Visible = False
        CommunityLivingWaiver.Visible = False
        AdultSupportsWaiver.Visible = False
        ResidentialWaiver = False
        CommunityLivingWaiver = False
        AdultSupportsWaiver = False
        If Form_frmPeopleClientsDemographics.AutismServices Then Form_frmPeopleClientsServiceAutism.FormerWaiverClient = True
    End If
'   Me.Dirty = False
End Sub

Private Sub ValidateAddress()
On Error GoTo ShowMeError
    Dim strUrl As String    ' Our URL which will include the authentication info
    Dim strReq As String    ' The body of the POST request
    Dim xmlHttp As New MSXML2.XMLHTTP60
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim candidates, candidate, components, metadata, analysis As MSXML2.IXMLDOMNode
    Dim AddressToCheck As Variant, CityToCheck As Variant, StateToCheck As Variant, ZIPToCheck As Variant, Validated As Boolean, District As Long, MatchCode As String, Footnotes As String
    Dim candidate_count As Long, Start, Finish
    
' **************************************************************************************************************************
' Address validation is provided under a free non-profit license by Smarty (www.smarty.com).  Also known as SmartyStreets. *
' **************************************************************************************************************************
    
' This URL will execute the search request and return the resulting matches to the search in XML.
    SysCmdResult = SysCmd(4, "Validating address.")
    strUrl = "https://api.smartystreets.com/street-address?auth-id=e9ae7820-92a0-4f43-9968-050355d25c21&auth-token=0hODirlqxklciKo5Hz8kTlRrG304A9JCq7SrycLfwRslllxSx%2B9y5dsbKUMswEyLFWF0taWP81XXWqBWBgMuCA%3D%3D"
    AddressToCheck = RepPayeeAddress
    CityToCheck = RepPayeeCity
    StateToCheck = RepPayeeState
    If Len(RepPayeeZIP) = 6 Then ZIPToCheck = Left(RepPayeeZIP, 5) Else ZIPToCheck = RepPayeeZIP
' Body of the POST request
    strReq = "<?xml version=""1.0"" encoding=""utf-8""?>" & "<request>" & "<address>" & _
                "   <street>" & AddressToCheck & "</street>" & "   <city>" & CityToCheck & "</city>" & _
                "   <state>" & StateToCheck & "</state>" & "   <zipcode>" & ZIPToCheck & "</zipcode>" & _
                "   <candidates>10</candidates>" & "</address>" & "</request>"
    
    With xmlHttp
        .Open "POST", strUrl, False                     ' Prepare POST request
        .setRequestHeader "Content-Type", "text/xml"    ' Sending XML ...
        .setRequestHeader "Accept", "text/xml"          ' ... expect XML in return.
        .send strReq                                    ' Send request body
    End With
    
' The request has been saved into xmlHttp.responseText and is
' now ready to be parsed. Remember that fields in our XML response may
' change or be added to later, so make sure your method of parsing accepts that.
' Google and Stack Overflow are replete with helpful examples.
    Set xmlDoc = New MSXML2.DOMDocument60
    If Not xmlDoc.loadXML(xmlHttp.responseText) Then
        Err.Raise xmlDoc.parseError.errorCode, , xmlDoc.parseError.reason
        SysCmdResult = SysCmd(5)
        Exit Sub
    End If
' According to the schema (http://smartystreets.com/kb/liveaddress-api/parsing-the-response#xml),
' <candidates> is a top-level node with each <candidate> below it. Let's obtain each one.
    Set candidates = xmlDoc.documentElement
    candidate_count = 0
' First, get a count of all the search results.
    For Each candidate In candidates.childNodes
        candidate_count = candidate_count + 1
    Next
    
    Select Case candidate_count
        Case 0 ' Bad address cannot be corrected.  Try again.
            MsgBox "The address supplied does not match a valid address in the USPS database.  Please correct this.", vbOKOnly, "Warning"
            SysCmdResult = SysCmd(5)
            Exit Sub
        Case 1 ' Only one candidate address...use it and return.
            For Each candidate In candidates.childNodes
                RepPayeeAddress = candidate.selectSingleNode("delivery_line_1").nodeTypedValue
                Set components = candidate.selectSingleNode("components")
                If Not components Is Nothing Then
                    RepPayeeCity = components.selectSingleNode("city_name").nodeTypedValue
                    RepPayeeState = components.selectSingleNode("state_abbreviation").nodeTypedValue
                    RepPayeeZIP = components.selectSingleNode("zipcode").nodeTypedValue & "-" & components.selectSingleNode("plus4_code").nodeTypedValue
                    RepPayeeAddressValidated = True
                End If
            Next
            Exit Sub
        Case Else ' Multiple candidate addresses...post them and allow the user to select.
            If IsTableQuery("temptbl") Then TILLDataBase.Execute "DROP TABLE temptbl", dbSeeChanges: Call BriefDelay
            TILLDataBase.Execute "CREATE TABLE temptbl (Selected BIT, CandidateAddress CHAR(50), CandidateCity CHAR(25), CandidateState CHAR(2), CandidateZIP CHAR(10), CandidateCongressionalDistrict INTEGER, MatchCode CHAR(1), Footnotes CHAR(30));", dbSeeChanges: Call BriefDelay
            Call BriefDelay
            For Each candidate In candidates.childNodes
                AddressToCheck = candidate.selectSingleNode("delivery_line_1").nodeTypedValue
                Set components = candidate.selectSingleNode("components")
                If Not components Is Nothing Then
                    CityToCheck = components.selectSingleNode("city_name").nodeTypedValue
                    StateToCheck = components.selectSingleNode("state_abbreviation").nodeTypedValue
                    ZIPToCheck = components.selectSingleNode("zipcode").nodeTypedValue & "-" & components.selectSingleNode("plus4_code").nodeTypedValue
                End If
                District = 0
                Set analysis = candidate.selectSingleNode("analysis")
                If Not analysis Is Nothing Then
                    MatchCode = analysis.selectSingleNode("dpv_match_code").nodeTypedValue
                    Footnotes = analysis.selectSingleNode("dpv_footnotes").nodeTypedValue
                End If
                TILLDataBase.Execute "INSERT INTO temptbl (Selected, CandidateAddress, CandidateCity, CandidateState, CandidateZIP, CandidateCongressionalDistrict, MatchCode, Footnotes)" & _
                    vbCrLf & "SELECT False AS Expr0, """ & AddressToCheck & """ AS Expr1, """ & CityToCheck & """ AS Expr2, """ & StateToCheck & _
                    """ AS Expr3, """ & ZIPToCheck & """ AS Expr4, " & District & " AS Expr5, """ & MatchCode & """ AS Expr6, """ & _
                    Footnotes & """ AS Expr7;", dbSeeChanges: Call BriefDelay
            Next
            DoCmd.OpenForm "frmPeopleClientsDemographicsAddressMaintenance"
            Me.SetFocus
    End Select
    SysCmdResult = SysCmd(5)
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Function PCA_Autism_Or_CommunityConnections_Springboard_TRASE()
    PCA_Autism_Or_CommunityConnections_Springboard_TRASE = (PCA Or AutismServices Or SPRINGBOARD Or CommunityConnections Or TRASE)
    If PCA_Autism_Or_CommunityConnections_Springboard_TRASE Then
        If DayServices Or ResidentialServices Or TransportationServices Or CLO Or VocationalServices Or SharedLiving Or NHDay Or NHRes Or IndivSupport Or AdultComp Or AdultCoach Then
            PCA_Autism_Or_CommunityConnections_Springboard_TRASE = False
            Exit Function
        End If
    End If
End Function

