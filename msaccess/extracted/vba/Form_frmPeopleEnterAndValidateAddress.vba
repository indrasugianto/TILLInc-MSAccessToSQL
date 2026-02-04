' Module Name: Form_frmPeopleEnterAndValidateAddress
' Module Type: Document Module
' Lines of Code: 140
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Private Sub PhysicalAddress_AfterUpdate()
    If Len(PhysicalAddress) > 0 And Not IsNull(PhysicalAddress) Then
        PhysicalAddressPopulated = True
        If IsNull(MailingAddress) Or Len(MailingAddress) = 0 Then
            MailingAddress = PhysicalAddress
            MailingAddressPopulated = True
        End If
        Call UpdateCounty
    End If
End Sub

Private Sub PhysicalCity_AfterUpdate()
    If Len(PhysicalCity) > 0 And Not IsNull(PhysicalCity) Then
        PhysicalCityPopulated = True
        If IsNull(MailingCity) Or Len(MailingCity) = 0 Then
            MailingCity = PhysicalCity
            MailingCityPopulated = True
        End If
        Call UpdateCounty
    End If
End Sub

Private Sub PhysicalState_AfterUpdate()
    If Len(PhysicalState) > 0 And Not IsNull(PhysicalState) Then
        PhysicalStatePopulated = True
        If IsNull(MailingState) Or Len(MailingState) = 0 Then
            MailingState = PhysicalState
            MailingStatePopulated = True
        End If
        Call UpdateCounty
    End If
End Sub

Private Sub PhysicalZIP_AfterUpdate()
    If Len(PhysicalZIP) > 0 And Not IsNull(PhysicalZIP) Then
        PhysicalZIPPopulated = True
        If IsNull(MailingZIP) Or Len(MailingZIP) = 0 Then
            MailingZIP = PhysicalZIP
            MailingZIPPopulated = True
        End If
    End If
End Sub

Private Sub MailingAddress_AfterUpdate()
    If Len(MailingAddress) > 0 And Not IsNull(MailingAddress) Then MailingAddressPopulated = True
End Sub

Private Sub MailingCity_AfterUpdate()
    If Len(MailingCity) > 0 And Not IsNull(MailingCity) Then MailingCityPopulated = True
End Sub

Private Sub MailingState_AfterUpdate()
    If Len(MailingState) > 0 And Not IsNull(MailingState) Then MailingStatePopulated = True
End Sub

Private Sub MailingZIP_AfterUpdate()
    If Len(MailingZIP) > 0 And Not IsNull(MailingZIP) Then MailingZIPPopulated = True
End Sub

Private Sub ValidateAndContinue_Click()
    Dim ProceedPhysical As Boolean, ProceedMailing As Boolean
    Dim EnteredPhysicalAddress As Boolean, EnteredPhysicalCity As Boolean, EnteredPhysicalState As Boolean, EnteredPhysicalZIP As Boolean
    Dim EnteredMailingAddress As Boolean, EnteredMailingCity As Boolean, EnteredMailingState As Boolean, EnteredMailingZIP As Boolean

    EnteredPhysicalAddress = Len(PhysicalAddress) > 0 And Not IsNull(PhysicalAddress)
    EnteredPhysicalCity = Len(PhysicalCity) > 0 And Not IsNull(PhysicalCity)
    EnteredPhysicalState = Len(PhysicalState) > 0 And Not IsNull(PhysicalState)
    EnteredPhysicalZIP = Len(PhysicalZIP) > 0 And Not IsNull(PhysicalZIP)
    EnteredMailingAddress = Len(MailingAddress) > 0 And Not IsNull(MailingAddress)
    EnteredMailingCity = Len(MailingCity) > 0 And Not IsNull(MailingCity)
    EnteredMailingState = Len(MailingState) > 0 And Not IsNull(MailingState)
    EnteredMailingZIP = Len(MailingZIP) > 0 And Not IsNull(MailingZIP)

    PhysicalAddress.BackColor = RGB(255, 255, 255)
    PhysicalCity.BackColor = RGB(255, 255, 255)
    PhysicalState.BackColor = RGB(255, 255, 255)
    PhysicalZIP.BackColor = RGB(255, 255, 255)
    MailingAddress.BackColor = RGB(255, 255, 255)
    MailingCity.BackColor = RGB(255, 255, 255)
    MailingState.BackColor = RGB(255, 255, 255)
    MailingZIP.BackColor = RGB(255, 255, 255)

    ProceedPhysical = EnteredPhysicalAddress And EnteredPhysicalCity And EnteredPhysicalState
    ProceedMailing = EnteredMailingAddress And EnteredMailingCity And EnteredMailingState
    
    If Not ProceedPhysical Or Not ProceedMailing Then   ' One of the addresses is not complete.
        MsgBox "One of the address fields is missing data.  Please complete the blank fields before proceeding", , "Error"
        AddressValidationFailed = True
        PhysicalAddress.SetFocus
    Else
        AddressValidationFailed = False
        With Form_frmPeople
            .PhysicalAddress = PhysicalAddress
            .PhysicalCity = PhysicalCity
            .PhysicalState = PhysicalState
            .PhysicalZIP = PhysicalZIP
            .CongressionalDistrict = Null
            If ValidateAddress(.PhysicalAddress, .PhysicalCity, .PhysicalState, .PhysicalZIP, .PhysicalAddressValidated, Me, .CongressionalDistrict) Then
                DoCmd.OpenForm "frmPeopleAddressMaintenance"
                Do Until CurrentProject.AllForms("frmPeopleAddressMaintenance").IsLoaded = False
                    DoEvents
                Loop
                .PhysicalAddress = GlobalAddress: Call UpdateChangeLog("PhysicalAddress", .PhysicalAddress)
                .PhysicalCity = GlobalCity:       Call UpdateChangeLog("PhysicalCity", .PhysicalCity)
                .PhysicalState = GlobalState:     Call UpdateChangeLog("PhysicalState", .PhysicalState)
                .PhysicalZIP = GlobalZIP:         Call UpdateChangeLog("PhysicalZIP", .PhysicalZIP)
                .PhysicalAddressValidated = GlobalValidated
                .CongressionalDistrict = GlobalCongressionalDistrict
            End If
            
            .MailingAddress = MailingAddress
            .MailingCity = MailingCity
            .MailingState = MailingState
            .MailingZIP = MailingZIP
            If ValidateAddress(.MailingAddress, .MailingCity, .MailingState, .MailingZIP, .MailingAddressValidated, Me, .MailingCongressionalDistrict) Then
                DoCmd.OpenForm "frmPeopleAddressMaintenance"
                Do Until CurrentProject.AllForms("frmPeopleAddressMaintenance").IsLoaded = False
                    DoEvents
                Loop
                .MailingAddress = GlobalAddress: Call UpdateChangeLog("MailingAddress", .MailingAddress)
                .MailingCity = GlobalCity:       Call UpdateChangeLog("MailingCity", .MailingCity)
                .MailingState = GlobalState:     Call UpdateChangeLog("MailingState", .MailingState)
                .MailingZIP = GlobalZIP:         Call UpdateChangeLog("MailingZIP", .MailingZIP)
                .MailingAddressValidated = GlobalValidated
            End If

            .IsClient.Locked = False
            .IsConsultant.Locked = False
            .IsDonor.Locked = False
            .IsFamilyGuardian.Locked = False
            .IsInterestedParty.Locked = False
            .IsStaff.Locked = False
        End With

        DoCmd.Close acForm, "frmPeopleEnterAndValidateAddress"
    End If
End Sub
