' Module Name: Form_frmPeopleEnterAndValidatePerson
' Module Type: Document Module
' Lines of Code: 116
' Extracted: 1/29/2026 4:12:26 PM

Option Compare Database
Option Explicit

Private Sub CancelAndReturn_Click()
    Form_frmMainMenu.NewPersonIndexedName = Null
    DoCmd.Close
End Sub

Private Sub Form_Load()
    SalutationEntered = False: LastNameEntered = False: FirstNameEntered = False: CompanyEntered = False: MiddleInitialEntered = False
End Sub

Private Function AfterUpdateOps(FieldCode As Integer)
    AfterUpdateOps = True
    Select Case FieldCode
        Case 1: If Len(NewPersonCompanyOrganization) > 0 Then CompanyEntered = True Else CompanyEntered = False
        Case 2: If Len(NewPersonFirstName) > 0 Then FirstNameEntered = True Else FirstNameEntered = False
        Case 3: If Len(NewPersonLastName) > 0 Then LastNameEntered = True Else LastNameEntered = False
        Case 4: If Len(NewPersonMiddleInitial) > 0 Then MiddleInitialEntered = True Else MiddleInitialEntered = False
        Case 5: If Len(NewPersonSalutation) > 0 Then SalutationEntered = True Else SalutationEntered = False
    End Select
End Function

Private Sub ValidateAndContinue_Click()
On Error GoTo ShowMeError
    Dim NewPersonFamiliarGreeting As String, LastNameComp As Variant, FirstNameComp As Variant, CompanyComp As Variant
    Dim NewRecordDate As Variant, NewUserEntering As Variant
    ' Initialize.
    NewPersonLastName.BackColor = RGB(255, 255, 255): NewPersonFirstName.BackColor = RGB(255, 255, 255): NewPersonCompanyOrganization.BackColor = RGB(255, 255, 255)
    ' Validate.  A company/organization can be entered without a last/first name.
    If CompanyEntered Then
        ' There could be a name.  Validate that it's completely provided.
        NewPersonCompanyOrganization.BackColor = RGB(255, 255, 255): NewPersonLastName.BackColor = RGB(255, 255, 255): NewPersonFirstName.BackColor = RGB(255, 255, 255)
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
            NewPersonLastName.BackColor = RGB(255, 255, 255):  NewPersonFirstName.BackColor = RGB(255, 255, 255):  NewPersonCompanyOrganization.BackColor = RGB(255, 255, 255)
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
                NewPersonLastName.BackColor = RGB(255, 0, 0): NewPersonFirstName.BackColor = RGB(255, 0, 0): NewPersonCompanyOrganization.BackColor = RGB(255, 0, 0)
                Exit Sub
            End If
        End If
    End If
    ' Make all name fields proper case and trimmed.
    NewPersonLastName = Trim(NewPersonLastName)
    NewPersonLastName = StrConv(NewPersonLastName, vbProperCase)
    NewPersonLastName = CorrectProperNames(NewPersonLastName)
    NewPersonFirstName = Trim(NewPersonFirstName)
    NewPersonFirstName = StrConv(NewPersonFirstName, vbProperCase)
    NewPersonFirstName = CorrectProperNames(NewPersonFirstName)
    NewPersonMiddleInitial = Trim(StrConv(NewPersonMiddleInitial, vbProperCase))
    NewPersonCompanyOrganization = Trim(NewPersonCompanyOrganization)
    NewPersonCompanyOrganization = StrConv(NewPersonCompanyOrganization, vbProperCase)
    NewPersonCompanyOrganization = CorrectProperNames(NewPersonCompanyOrganization)
    NewPersonTitle = Trim(NewPersonTitle)
    NewPersonTitle = StrConv(NewPersonTitle, vbProperCase)
    NewPersonTitle = CorrectProperNames(NewPersonTitle)
    NewPersonFamiliarGreeting = "Greetings"
    If SalutationEntered Then
        If LastNameEntered Then NewPersonFamiliarGreeting = NewPersonSalutation & " " & NewPersonLastName
    Else
        If FirstNameEntered Then NewPersonFamiliarGreeting = NewPersonFirstName
    End If
    ' Handle special names like MacSomeone, McSomeone, O'Someone, and hyphenated names.
    NewPersonFirstName = SpecialNames(NewPersonFirstName): NewPersonLastName = SpecialNames(NewPersonLastName)
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
                ResetPeopleRecordButtons = True
                DoCmd.Close acForm, "frmPeopleEnterAndValidatePerson"
                Exit Sub
            Else
                ' Add the new record.
                NewRecordDate = Format(Now(), "mm/dd/yyyy")
                NewUserEntering = Form_frmMainMenu.UserName
                TILLDataBase.Execute "INSERT INTO tblPeople (IndexedName, RecordAddedDate, RecordAddedBy, Salutation, LastName, FirstName, MiddleInitial, CompanyOrganization, Title, FamiliarGreeting ) " & _
                    "SELECT """ & NewPersonIndexedName & """ AS IndexedName, """ & NewRecordDate & """ AS RecordAddedDate, """ & NewUserEntering & """ AS RecordAddedBy, """ & NewPersonSalutation & """ AS Salutation, """ & NewPersonLastName & """ AS LastName, """ & NewPersonFirstName & """ AS FirstName, """ & NewPersonMiddleInitial & """ AS MiddleInitial, """ & NewPersonCompanyOrganization & """ AS CompanyOrganization, """ & NewPersonTitle & """ AS Title, """ & NewPersonFamiliarGreeting & """ AS FamiliarGreeting ", dbSeeChanges: Call BriefDelay
                Form_frmMainMenu.NewPersonIndexedName = NewPersonIndexedName
                Call UpdateChangeLog("Client Record Created", "")
                DoCmd.Close acForm, "frmPeopleEnterAndValidatePerson"
            End If
        End If
    End With
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub
