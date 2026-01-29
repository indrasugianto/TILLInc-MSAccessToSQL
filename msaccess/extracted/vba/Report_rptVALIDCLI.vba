' Module Name: Report_rptVALIDCLI
' Module Type: Document Module
' Lines of Code: 146
' Extracted: 1/29/2026 4:12:25 PM

Option Compare Database

Private Sub GroupHeader0_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo ShowMeError
    Dim WhichLocation As Variant, WhichProgram As Variant, Criteria As String
    
    With Report_rptVALIDCLI
        Criteria = "IndexedName = """ & .IndexedName & """"
        .Service.ForeColor = RGB(0, 0, 0)
        .IsDayLabel.ForeColor = RGB(146, 146, 146): .IsDayLabel.FontBold = False
        .IsResLabel.ForeColor = RGB(146, 146, 146): .IsResLabel.FontBold = False
        .IsTransLabel.ForeColor = RGB(146, 146, 146): .IsTransLabel.FontBold = False
        .IsVocLabel.ForeColor = RGB(146, 146, 146): .IsVocLabel.FontBold = False
        .IsSHCLabel.ForeColor = RGB(146, 146, 146): .IsSHCLabel.FontBold = False
        .IsNHLabel.ForeColor = RGB(146, 146, 146): .IsNHLabel.FontBold = False
        .IsNHDayLabel.ForeColor = RGB(146, 146, 146): .IsNHDayLabel.FontBold = False
        .IsCLOLabel.ForeColor = RGB(146, 146, 146): .IsCLOLabel.FontBold = False
        .IsISSLabel.ForeColor = RGB(146, 146, 146): .IsISSLabel.FontBold = False
        .IsAutLabel.ForeColor = RGB(146, 146, 146): .IsAutLabel.FontBold = False
        .IsPCALabel.ForeColor = RGB(146, 146, 146): .IsPCALabel.FontBold = False
        .IsSpringLabel.ForeColor = RGB(146, 146, 146): .IsSpringLabel.FontBold = False
        .IsTRASELabel.ForeColor = RGB(146, 146, 146): .IsTRASELabel.FontBold = False
        .IsRecLabel.ForeColor = RGB(146, 146, 146): .IsRecLabel.FontBold = False
    End With

    If IsISS Then
        Service = "Individual Support": IsISSLabel.ForeColor = RGB(0, 128, 0): IsISSLabel.FontBold = True
    ElseIf IsNHRes Then
        Service = "NH Residential"
        IsNHLabel.ForeColor = RGB(0, 128, 0): IsNHLabel.FontBold = True
    ElseIf IsNHDay Then
        Service = "NH Day Services"
        IsNHDayLabel.ForeColor = RGB(0, 128, 0): IsNHDayLabel.FontBold = True
    ElseIf IsTrans Then
        Service = "Transportation"
        IsTransLabel.ForeColor = RGB(0, 128, 0): IsTransLabel.FontBold = True
    ElseIf IsRes Then
        IsResLabel.ForeColor = RGB(0, 128, 0): IsResLabel.FontBold = True
        WhichLocation = DLookup("CityTown", "tblPeopleClientsResidentialServices", Criteria)
        If IsNull(WhichLocation) Then
            Service = "Residential NOT SPECIFIED": Service.ForeColor = RGB(255, 0, 0)
            GoTo Continue
        End If
        WhichProgram = DLookup("Location", "tblPeopleClientsResidentialServices", Criteria)
        Service = "RESIDENTIAL: " & WhichLocation & " - " & WhichProgram
        GoTo Continue
    ElseIf IsDay Then
        IsDayLabel.ForeColor = RGB(0, 128, 0): IsDayLabel.FontBold = True
        WhichLocation = DLookup("CityTown", "tblPeopleClientsDayServices", Criteria)
        If IsNull(WhichLocation) Then
            Service = "Day Program NOT SPECIFIED": Service.ForeColor = RGB(255, 0, 0)
            GoTo Continue
        End If
        WhichProgram = DLookup("LocationName", "tblPeopleClientsDayServices", Criteria)
        Service = "DAY SERVICES: " & WhichLocation & " - " & WhichProgram
        GoTo Continue
    ElseIf IsCLO Then
        IsCLOLabel.ForeColor = RGB(0, 128, 0): IsCLOLabel.FontBold = True
        WhichLocation = DLookup("CityTown", "tblPeopleClientsCLOServices", Criteria)
        If IsNull(WhichLocation) Then
            Service = "CLO NOT SPECIFIED": Service.ForeColor = RGB(255, 0, 0)
            GoTo Continue
        End If
        WhichProgram = DLookup("Location", "tblPeopleClientsCLOServices", Criteria)
        Service = "CLO: " & WhichLocation & " - " & WhichProgram
        GoTo Continue
    ElseIf IsVoc Then
        IsVocLabel.ForeColor = RGB(0, 128, 0): IsVocLabel.FontBold = True
        WhichLocation = DLookup("CityTown", "tblPeopleClientsVocationalServices", Criteria)
        If IsNull(WhichLocation) Then
            Service = "Vocational Program NOT SPECIFIED": Service.ForeColor = RGB(255, 0, 0)
            GoTo Continue
        End If
        WhichProgram = DLookup("Location", "tblPeopleClientsVocationalServices", Criteria)
        Service = "VOCATIONAL: " & WhichLocation & " - " & WhichProgram
        GoTo Continue
    Else
        Service.Visible = False
    End If
    
Continue:
    
    If Not PhysicalAddressValidated Then
        PhysicalAddress.BackColor = RGB(255, 255, 0): PhysicalCityStateZIP.BackColor = RGB(255, 255, 0)
    Else
        PhysicalAddress.BackColor = RGB(255, 255, 255): PhysicalCityStateZIP.BackColor = RGB(255, 255, 255)
    End If
    
    If Not MailingAddressValidated Then
        MailingAddress.BackColor = RGB(255, 255, 0): MailingCityStateZIP.BackColor = RGB(255, 255, 0)
    Else
        MailingAddress.BackColor = RGB(255, 255, 255): MailingCityStateZIP.BackColor = RGB(255, 255, 255)
    End If

    If IsNull(HomePhone) Then HomePhone.BackColor = RGB(255, 255, 0) Else HomePhone.BackColor = RGB(255, 255, 255)
    If IsNull(WorkPhone) Then WorkPhone.BackColor = RGB(255, 255, 0) Else WorkPhone.BackColor = RGB(255, 255, 255)
    If IsNull(MobilePhone) Then MobilePhone.BackColor = RGB(255, 255, 0) Else MobilePhone.BackColor = RGB(255, 255, 255)
    If IsNull(EmailAddress) Then EmailAddress.BackColor = RGB(255, 255, 0) Else EmailAddress.BackColor = RGB(255, 255, 255)
    If IsNull(Gender) Then Gender.BackColor = RGB(255, 255, 0) Else Gender.BackColor = RGB(255, 255, 255)
    If IsNull(Race) Then Race.BackColor = RGB(255, 255, 0) Else Race.BackColor = RGB(255, 255, 255)
    If IsNull(DateOfBirth) Then DateOfBirth.BackColor = RGB(255, 255, 0) Else DateOfBirth.BackColor = RGB(255, 255, 255)
    If IsNull(LegalName) Then LegalName.BackColor = RGB(255, 255, 0) Else LegalName.BackColor = RGB(255, 255, 255)
    If IsNull(LegalStatus) Then LegalStatus.BackColor = RGB(255, 255, 0) Else LegalStatus.BackColor = RGB(255, 255, 255)
    If IsNull(PlaceOfBirth) Then PlaceOfBirth.BackColor = RGB(255, 255, 0) Else PlaceOfBirth.BackColor = RGB(255, 255, 255)
    If IsNull(SocialSecurityNumber) Then SocialSecurityNumber.BackColor = RGB(255, 255, 0) Else SocialSecurityNumber.BackColor = RGB(255, 255, 255)
    If IsNull(MedicaidNumber) Then MedicaidNumber.BackColor = RGB(255, 255, 0) Else MedicaidNumber.BackColor = RGB(255, 255, 255)
    If IsNull(MedicaidLastCertDate) Then MedicaidLastCertDate.BackColor = RGB(255, 255, 0) Else MedicaidLastCertDate.BackColor = RGB(255, 255, 255)
    If IsNull(MedicareNumber) Then MedicareNumber.BackColor = RGB(255, 255, 0) Else MedicareNumber.BackColor = RGB(255, 255, 255)
    If IsNull(SSAMonthlyAmount) Then SSAMonthlyAmount.BackColor = RGB(255, 255, 0) Else SSAMonthlyAmount.BackColor = RGB(255, 255, 255)
    If IsNull(SSABenefitsLetterDate) Then SSABenefitsLetterDate.BackColor = RGB(255, 255, 0) Else SSABenefitsLetterDate.BackColor = RGB(255, 255, 255)
    If IsNull(SSIMonthlyAmount) Then SSIMonthlyAmount.BackColor = RGB(255, 255, 0) Else SSIMonthlyAmount.BackColor = RGB(255, 255, 255)
    If IsNull(SSPMonthlyAmount) Then SSPMonthlyAmount.BackColor = RGB(255, 255, 0) Else SSPMonthlyAmount.BackColor = RGB(255, 255, 255)
    If IsNull(WaiverType) Then WaiverType.BackColor = RGB(255, 255, 0) Else WaiverType.BackColor = RGB(255, 255, 255)
    If IsNull(FoodStampsCardNumber) Then FoodStampsCardNumber.BackColor = RGB(255, 255, 0) Else FoodStampsCardNumber.BackColor = RGB(255, 255, 255)
    If IsNull(FoodStampsAmount) Then FoodStampsAmount.BackColor = RGB(255, 255, 0) Else FoodStampsAmount.BackColor = RGB(255, 255, 255)
    If IsNull(FoodStampsLastCertDate) Then FoodStampsLastCertDate.BackColor = RGB(255, 255, 0) Else FoodStampsLastCertDate.BackColor = RGB(255, 255, 255)
    If IsNull(FoodStampsOffice) Then FoodStampsOffice.BackColor = RGB(255, 255, 0) Else FoodStampsOffice.BackColor = RGB(255, 255, 255)
    If IsNull(OtherBenefits) Then OtherBenefits.BackColor = RGB(255, 255, 0) Else OtherBenefits.BackColor = RGB(255, 255, 255)
    If IsNull(OtherMonthlyAmount) Then OtherMonthlyAmount.BackColor = RGB(255, 255, 0) Else OtherMonthlyAmount.BackColor = RGB(255, 255, 255)
    If IsNull(OtherInsurance) Then OtherInsurance.BackColor = RGB(255, 255, 0) Else OtherInsurance.BackColor = RGB(255, 255, 255)
    If IsNull(Mother) Then Mother.BackColor = RGB(255, 255, 0) Else Mother.BackColor = RGB(255, 255, 255)
    If IsNull(MotherDateOfBirth) Then MotherDateOfBirth.BackColor = RGB(255, 255, 0) Else MotherDateOfBirth.BackColor = RGB(255, 255, 255)
    If IsNull(MotherDateOfDeath) Then MotherDateOfDeath.BackColor = RGB(255, 255, 0) Else MotherDateOfDeath.BackColor = RGB(255, 255, 255)
    If IsNull(Father) Then Father.BackColor = RGB(255, 255, 0) Else Father.BackColor = RGB(255, 255, 255)
    If IsNull(FatherDateOfBirth) Then FatherDateOfBirth.BackColor = RGB(255, 255, 0) Else FatherDateOfBirth.BackColor = RGB(255, 255, 255)
    If IsNull(FatherDateOfDeath) Then FatherDateOfDeath.BackColor = RGB(255, 255, 0) Else FatherDateOfDeath.BackColor = RGB(255, 255, 255)
    If IsNull(Siblings) Then Siblings.BackColor = RGB(255, 255, 0) Else Siblings.BackColor = RGB(255, 255, 255)
    If IsNull(Bank) Then Bank.BackColor = RGB(255, 255, 0) Else Bank.BackColor = RGB(255, 255, 255)
    If IsNull(BankAccountNumber) Then BankAccountNumber.BackColor = RGB(255, 255, 0) Else BankAccountNumber.BackColor = RGB(255, 255, 255)
    If IsNull(BankAddress) Then BankAddress.BackColor = RGB(255, 255, 0) Else BankAddress.BackColor = RGB(255, 255, 255)
    If IsNull(BankCityStateZIP) Then BankCityStateZIP.BackColor = RGB(255, 255, 0) Else BankCityStateZIP.BackColor = RGB(255, 255, 255)
    If IsNull(BankTypeOfAccount) Then BankTypeOfAccount.BackColor = RGB(255, 255, 0) Else BankTypeOfAccount.BackColor = RGB(255, 255, 255)
    If IsNull(BankRoutingNumber) Then BankRoutingNumber.BackColor = RGB(255, 255, 0) Else BankRoutingNumber.BackColor = RGB(255, 255, 255)
    If IsNull(BankPhoneNumber) Then BankPhoneNumber.BackColor = RGB(255, 255, 0) Else BankPhoneNumber.BackColor = RGB(255, 255, 255)
    If IsNull(RepresentativePayee) Then RepresentativePayee.BackColor = RGB(255, 255, 0) Else RepresentativePayee.BackColor = RGB(255, 255, 255)
    If IsNull(RepPayeeReportDate) Then RepPayeeReportDate.BackColor = RGB(255, 255, 0) Else RepPayeeReportDate.BackColor = RGB(255, 255, 255)
    If IsNull(RepPayeeAddress) Then RepPayeeAddress.BackColor = RGB(255, 255, 0) Else RepPayeeAddress.BackColor = RGB(255, 255, 255)
    If IsNull(RepPayeeCityStateZIP) Then RepPayeeCityStateZIP.BackColor = RGB(255, 255, 0) Else RepPayeeCityStateZIP.BackColor = RGB(255, 255, 255)
    If IsNull(ResidentialVendor) Then ResidentialVendor.BackColor = RGB(255, 255, 0) Else ResidentialVendor.BackColor = RGB(255, 255, 255)
    If IsNull(ResidentialVendorPhoneNumber) Then ResidentialVendorPhoneNumber.BackColor = RGB(255, 255, 0) Else ResidentialVendorPhoneNumber.BackColor = RGB(255, 255, 255)
    If IsNull(DayVendor) Then DayVendor.BackColor = RGB(255, 255, 0) Else DayVendor.BackColor = RGB(255, 255, 255)
    If IsNull(DayVendorPhoneNumber) Then DayVendorPhoneNumber.BackColor = RGB(255, 255, 0) Else DayVendorPhoneNumber.BackColor = RGB(255, 255, 255)
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub