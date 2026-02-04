' Module Name: Form_frmPeopleFamilyEnterAndValidatePerson
' Module Type: Document Module
' Lines of Code: 280
' Extracted: 2026-02-04 13:03:36

Option Compare Database
Option Explicit

Private Sub CancelAndReturn_Click()
    Form_frmMainMenu.NewPersonIndexedName = Null
    DoCmd.Close
End Sub

Private Sub Form_Current()
    NewPersonSalutation = Null: NewPersonFirstName = Null: NewPersonMiddleInitial = Null: NewPersonLastName = Null: NewPersonCompanyOrganization = Null
    NewPersonTitle = Null: NewPhysicalAddress = Null: NewPhysicalCity = Null: NewPhysicalState = Null: NewPhysicalZIP = Null: NewMailingAddress = Null
    NewMailingCity = Null: NewMailingState = Null: NewMailingZIP = Null: NewRelationship = Null: NewContact = False: NewGuardian = False
    NewSurrogate = False: NewRepPayee = False
End Sub

Private Sub Form_Load()
    Dim CommaPointer As Integer
    
    SalutationEntered = False: LastNameEntered = False: FirstNameEntered = False: CompanyEntered = False: MiddleInitialEntered = False
    NewClientIndexedName = Form_frmPeopleClientsDemographics.IndexedName
    NewClientLastName = Form_frmPeopleClientsDemographics.LastName
    NewClientMiddleInitial = Form_frmPeopleClientsDemographics.MiddleInitial
    NewClientFirstName = Form_frmPeopleClientsDemographics.FirstName
End Sub

Private Sub NewPersonCompanyOrganization_AfterUpdate()
    If Len(NewPersonCompanyOrganization) = 0 Then CompanyEntered = False Else CompanyEntered = True
End Sub

Private Sub NewPersonFirstName_AfterUpdate()
    If Len(NewPersonFirstName) = 0 Then FirstNameEntered = False Else FirstNameEntered = True
End Sub

Private Sub NewPersonLastName_AfterUpdate()
    If Len(NewPersonLastName) = 0 Then LastNameEntered = False Else LastNameEntered = True
End Sub

Private Sub NewPersonMiddleInitial_AfterUpdate()
    If Len(NewPersonMiddleInitial) = 0 Then MiddleInitialEntered = False Else MiddleInitialEntered = True
End Sub

Private Sub NewPersonSalutation_AfterUpdate()
    If Len(NewPersonSalutation) = 0 Then SalutationEntered = False Else SalutationEntered = True
End Sub

Private Sub ValidateAndContinue_Click()
On Error GoTo ShowMeError
    Dim NewPersonIndexedName As String, NewPersonFamiliarGreeting As String, LastNameComp As Variant, FirstNameComp As Variant, CompanyComp As Variant, RememberIndexedName As Variant
    Dim ProceedPhysical As Boolean, ProceedMailing As Boolean
    Dim EnteredPhysicalAddress As Boolean, EnteredPhysicalCity As Boolean, EnteredPhysicalState As Boolean, EnteredPhysicalZIP As Boolean
    Dim EnteredMailingAddress As Boolean, EnteredMailingCity As Boolean, EnteredMailingState As Boolean, EnteredMailingZIP As Boolean
    ' Initialize.
    NewPersonLastName.BackColor = RGB(255, 255, 255)
    NewPersonFirstName.BackColor = RGB(255, 255, 255)
    NewPersonCompanyOrganization.BackColor = RGB(255, 255, 255)
    ' Validate.  A company/organization can be entered without a last/first name.
    If CompanyEntered Then
        NewPersonCompanyOrganization.BackColor = RGB(255, 255, 255)
        ' There could be a name.  Validate that it's completely provided.
        NewPersonLastName.BackColor = RGB(255, 255, 255)
        NewPersonFirstName.BackColor = RGB(255, 255, 255)
        If FirstNameEntered And Not LastNameEntered Then
            NewPersonLastName.BackColor = RGB(255, 0, 0)
            Exit Sub
        End If
        If LastNameEntered And Not FirstNameEntered Then
            NewPersonFirstName.BackColor = RGB(255, 0, 0)
            Exit Sub
        End If
    ' No company.  Last and first names must be provided.
    Else
        If LastNameEntered And FirstNameEntered Then
            NewPersonLastName.BackColor = RGB(255, 255, 255)
            NewPersonFirstName.BackColor = RGB(255, 255, 255)
            NewPersonCompanyOrganization.BackColor = RGB(255, 255, 255)
        Else
            If Not LastNameEntered Then
                NewPersonLastName.BackColor = RGB(255, 0, 0)
                Exit Sub
            End If
            If Not FirstNameEntered Then
                NewPersonFirstName.BackColor = RGB(255, 0, 0)
                Exit Sub
            End If
            If (Not LastNameEntered) And (Not FirstNameEntered) Then
                NewPersonLastName.BackColor = RGB(255, 0, 0)
                NewPersonFirstName.BackColor = RGB(255, 0, 0)
                NewPersonCompanyOrganization.BackColor = RGB(255, 0, 0)
                Exit Sub
            End If
        End If
    End If
    ' Check addresses.
    EnteredPhysicalAddress = Len(NewPhysicalAddress) > 0
    EnteredPhysicalCity = Len(NewPhysicalCity) > 0
    EnteredPhysicalState = Len(NewPhysicalState) > 0
    EnteredMailingAddress = Len(NewMailingAddress) > 0
    EnteredMailingCity = Len(NewMailingCity) > 0
    EnteredMailingState = Len(NewMailingState) > 0
    ProceedPhysical = PhysicalAddressPopulated And PhysicalCityPopulated And PhysicalStatePopulated
    ProceedMailing = MailingAddressPopulated And MailingCityPopulated And MailingStatePopulated
    
    If Not ProceedPhysical Then
        MsgBox "The Residential Address, City, or State contain no data.  Please complete the blank fields before proceeding", , "Error"
        Exit Sub
    End If
        
    If Not ProceedMailing Then
        MsgBox "The Mailing Address, City, or State contain no data.  Please complete the blank fields before proceeding", , "Error"
        Exit Sub
    End If
    ' Make all name fields proper case and trimmed.
    NewPersonLastName = CorrectProperNames(Trim(StrConv(NewPersonLastName, vbProperCase)))
    NewPersonFirstName = CorrectProperNames(Trim(StrConv(NewPersonFirstName, vbProperCase)))
    NewPersonMiddleInitial = CorrectProperNames(Trim(StrConv(NewPersonMiddleInitial, vbProperCase)))
    NewPersonCompanyOrganization = CorrectProperNames(Trim(StrConv(NewPersonCompanyOrganization, vbProperCase)))
    NewPersonTitle = CorrectProperNames(Trim(StrConv(NewPersonTitle, vbProperCase)))
    NewPersonFamiliarGreeting = "Greetings"
    If SalutationEntered Then
        If LastNameEntered Then NewPersonFamiliarGreeting = NewPersonSalutation & " " & NewPersonLastName
    Else
        If FirstNameEntered Then NewPersonFamiliarGreeting = NewPersonFirstName
    End If
    ' Handle special names like MacSomeone, McSomeone, O'Someone, and hyphenated names.
    If FirstNameEntered Then NewPersonFirstName = SpecialNames(NewPersonFirstName)
    If LastNameEntered Then NewPersonLastName = SpecialNames(NewPersonLastName)
    ' Create the indexed name.
    NewPersonIndexedName = NewPersonLastName & "/" & NewPersonFirstName & "/" & NewPersonMiddleInitial & "/" & NewPersonCompanyOrganization
    ' Count, looking for duplicate name.
    LastNameComp = StrComp(NewPersonLastName, "", vbTextCompare)
    FirstNameComp = StrComp(NewPersonFirstName, "", vbTextCompare)
    CompanyComp = StrComp(NewPersonCompanyOrganization, "", vbTextCompare)
    With Form_frmPeople
        If ((LastNameComp = -1 Or LastNameComp = 1) And (FirstNameComp = -1 Or FirstNameComp = 1)) Or ((CompanyComp = -1) Or (CompanyComp = 1)) Then         ' Check for a duplicate record.
            If DCount("IndexedName", "tblPeople", "IndexedName = """ & NewPersonIndexedName & """") > 0 Then
                MsgBox NewPersonFirstName & " " & NewPersonMiddleInitial & " " & NewPersonLastName & " " & NewPersonCompanyOrganization & " already exists in the database...taking you there.", , "Duplicate!"
                Form_frmMainMenu.NewRecordInProgress = False
                .Filter = "IndexedName = """ & NewPersonIndexedName & """"
                .FilterOn = True
                Form_frmMainMenu.NewPersonIndexedName = Null
                DoCmd.Close acForm, "frmPeopleEnterAndValidatePerson"
                Exit Sub
            Else
                ' Add the new record.
                TILLDataBase.Execute "INSERT INTO tblPeople (IndexedName, RecordAddedDate, RecordAddedBy, Salutation, LastName, FirstName, MiddleInitial, CompanyOrganization, Title, FamiliarGreeting, " & _
                    "PhysicalAddress, PhysicalCity, PhysicalState, PhysicalZIP, " & _
                    "MailingAddress, MailingCity, MailingState, MailingZIP, IsFamilyGuardian) " & _
                    "SELECT '" & NewPersonIndexedName & "' AS IndexedName, " & _
                        "'" & Format(Now(), "mm/dd/yyyy") & "' AS RecordAddedDate, " & _
                        "'" & Form_frmMainMenu.UserName & "' AS RecordAddedBy, " & _
                        "'" & NewPersonSalutation & "' AS Salutation, " & _
                        "'" & NewPersonLastName & "' AS LastName, " & _
                        "'" & NewPersonFirstName & "' AS FirstName, " & _
                        "'" & NewPersonMiddleInitial & "' AS MiddleInitial, " & _
                        "'" & NewPersonCompanyOrganization & "' AS CompanyOrganization, " & _
                        "'" & NewPersonTitle & "' AS Title, " & _
                        "'" & NewPersonFamiliarGreeting & "' AS FamiliarGreeting, " & _
                        "'" & NewPhysicalAddress & "' AS PhysicalAddress, " & _
                        "'" & NewPhysicalCity & "' AS PhysicalCity, " & _
                        "'" & NewPhysicalState & "' AS PhysicalState, " & _
                        "'" & NewPhysicalZIP & "' AS PhysicalZIP, " & _
                        "'" & NewMailingAddress & "' AS MailingAddress, " & _
                        "'" & NewMailingCity & "' AS MailingCity, " & _
                        "'" & NewMailingState & "' AS MailingState, " & _
                        "'" & NewMailingZIP & "' AS MailingZIP, True as IsFamilyGuardian ", dbSeeChanges: Call BriefDelay
                TILLDataBase.Execute "INSERT INTO tblPeopleFamily (IndexedName, ClientIndexedName, RecordAddedDate, RecordAddedBy, ClientLastName, ClientFirstName, ClientMiddleInitial, Relationship, Guardian, PrimaryContact, Surrogate, RepPayee, Inactive) " & _
                    "SELECT '" & NewPersonIndexedName & "' AS IndexedName, " & _
                        "'" & NewClientIndexedName & "' AS ClientIndexedName, " & _
                        "'" & Format(Now(), "mm/dd/yyyy") & "' AS RecordAddedDate, " & _
                        "'" & Form_frmMainMenu.UserName & "' AS RecordAddedBy, " & _
                        "'" & NewClientLastName & "' AS ClientLastName, " & _
                        "'" & NewClientFirstName & "' AS ClientFirstName, " & _
                        "'" & NewClientMiddleInitial & "' AS ClientMiddleInitial, " & _
                        "'" & NewRelationship & "' AS Relationship, " & _
                        NewGuardian & " AS Guardian, " & _
                        NewContact & " AS PrimaryContact, " & _
                        NewSurrogate & " AS Surrogate, " & _
                        NewRepPayee & " AS RepPayee, FALSE as Inactive ", dbSeeChanges: Call BriefDelay
                Call UpdateChangeLog("Family Record Created", "")
            End If
        End If
    End With
    NewPersonSalutation = Null: NewPersonFirstName = Null: NewPersonMiddleInitial = Null: NewPersonLastName = Null: NewPersonCompanyOrganization = Null
    NewPersonTitle = Null: NewPhysicalAddress = Null: NewPhysicalCity = Null: NewPhysicalState = Null: NewPhysicalZIP = Null: NewMailingAddress = Null
    NewMailingCity = Null: NewMailingState = Null: NewMailingZIP = Null: NewRelationship = Null: NewContact = False: NewGuardian = False
    NewSurrogate = False: NewRepPayee = False
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub ValidateAndExit_Click()
    Call ValidateAndContinue_Click
    Form_frmMainMenu.NewPersonIndexedName = Null
    Form_frmPeople.Requery
    DoCmd.Close acForm, "frmPeopleFamilyEnterAndValidatePerson"
End Sub

Private Sub NewPhysicalAddress_AfterUpdate()
    If Len(NewPhysicalAddress) > 0 Then
        PhysicalAddressPopulated = True
        If IsNull(NewMailingAddress) Or Len(NewMailingAddress) = 0 Then
            NewMailingAddress = NewPhysicalAddress
            MailingAddressPopulated = True
        End If
    End If
End Sub

Private Sub NewPhysicalCity_AfterUpdate()
    If Len(NewPhysicalCity) > 0 Then
        PhysicalCityPopulated = True
        If IsNull(NewMailingCity) Or Len(NewMailingCity) = 0 Then
            NewMailingCity = NewPhysicalCity
            MailingCityPopulated = True
        End If
    End If
End Sub

Private Sub NewPhysicalState_AfterUpdate()
    If Len(NewPhysicalState) > 0 Then
        PhysicalStatePopulated = True
        If IsNull(NewMailingState) Or Len(NewMailingState) = 0 Then
            NewMailingState = NewPhysicalState
            MailingStatePopulated = True
        End If
        Call UpdateCounty
    End If
End Sub

Private Sub NewPhysicalZIP_AfterUpdate()
    If Len(NewPhysicalZIP) > 0 Then
        PhysicalZIPPopulated = True
        If IsNull(NewMailingZIP) Or Len(NewMailingZIP) = 0 Then
            NewMailingZIP = NewPhysicalZIP
            MailingZIPPopulated = True
        End If
    End If
End Sub

Private Sub NewMailingAddress_AfterUpdate()
    If Len(NewMailingAddress) > 0 Then MailingAddressPopulated = True
End Sub

Private Sub NewMailingCity_AfterUpdate()
    If Len(NewMailingCity) > 0 Then MailingCityPopulated = True
End Sub

Private Sub NewMailingState_AfterUpdate()
    If Len(NewMailingState) > 0 Then MailingStatePopulated = True
End Sub

Private Sub NewMailingZIP_AfterUpdate()
    If Len(NewMailingZIP) > 0 Then MailingZIPPopulated = True
End Sub

Private Sub ValidateAddressAndContinue()
    Dim ProceedPhysical As Boolean, ProceedMailing As Boolean
    Dim EnteredPhysicalAddress As Boolean, EnteredPhysicalCity As Boolean, EnteredPhysicalState As Boolean, EnteredPhysicalZIP As Boolean
    Dim EnteredMailingAddress As Boolean, EnteredMailingCity As Boolean, EnteredMailingState As Boolean, EnteredMailingZIP As Boolean

    EnteredPhysicalAddress = Len(NewPhysicalAddress) > 0
    EnteredPhysicalCity = Len(NewPhysicalCity) > 0
    EnteredPhysicalState = Len(NewPhysicalState) > 0
    EnteredMailingAddress = Len(NewMailingAddress) > 0
    EnteredMailingCity = Len(NewMailingCity) > 0
    EnteredMailingState = Len(NewMailingState) > 0

    ProceedPhysical = PhysicalAddressPopulated And PhysicalCityPopulated And PhysicalStatePopulated
    ProceedMailing = MailingAddressPopulated And MailingCityPopulated And MailingStatePopulated
    
    If Not ProceedPhysical Then
        MsgBox "The Residential Address, City, or State contain no data.  Please complete the blank fields before proceeding", , "Error"
        Exit Sub
    End If
        
    If Not ProceedMailing Then
        MsgBox "The Mailing Address, City, or State contain no data.  Please complete the blank fields before proceeding", , "Error"
        Exit Sub
    End If
End Sub
