' Module Name: Form_frmPeopleChangeName
' Module Type: Document Module
' Lines of Code: 155
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Dim NameChange As String, SalutationChanged As Boolean, ClusterNum As Variant, CLusterCriteria As Variant

Private Sub frmPeopleLockNameFields()
    With Form_frmPeople
        .Salutation.Locked = True:    .Salutation.Enabled = False:    .Salutation.BackColor = RGB(255, 255, 255)
        .LastName.Locked = True:      .LastName.Enabled = False:      .LastName.BackColor = RGB(255, 255, 255)
        .FirstName.Locked = True:     .FirstName.Enabled = False:     .FirstName.BackColor = RGB(255, 255, 255)
        .MiddleInitial.Locked = True: .MiddleInitial.Enabled = False: .MiddleInitial.BackColor = RGB(255, 255, 255)
        .UpperCompany.Locked = True:  .UpperCompany.Enabled = False
        .LowerCompany.Locked = True:  .LowerCompany.Enabled = False
        .Title.Locked = True:         .Title.Enabled = False
    End With
End Sub

Private Sub frmPeopleUnlockNameFields()
    With Form_frmPeople
        .Salutation.Locked = False:    .Salutation.Enabled = True:    .Salutation.BackColor = RGB(255, 255, 255)
        .LastName.Locked = False:      .LastName.Enabled = True:      .LastName.BackColor = RGB(255, 255, 255)
        .FirstName.Locked = False:     .FirstName.Enabled = True:     .FirstName.BackColor = RGB(255, 255, 255)
        .MiddleInitial.Locked = False: .MiddleInitial.Enabled = True: .MiddleInitial.BackColor = RGB(255, 255, 255)
        .UpperCompany.Locked = False:  .UpperCompany.Enabled = True
        .LowerCompany.Locked = False:  .LowerCompany.Enabled = True
        .Title.Locked = False:         .Title.Enabled = True
    End With
End Sub

Private Sub Form_Current()
    NewSalutation = Form_frmPeople.Salutation
    NewLastName = Form_frmPeople.LastName
    NewFirstName = Form_frmPeople.FirstName
    NewMiddleInitial = Form_frmPeople.MiddleInitial
    NewCompanyOrganization = Form_frmPeople.LowerCompany
    NewTitle = Form_frmPeople.Title
    OldIndexedName = Form_frmPeople.IndexedName
End Sub

Private Sub MakeChange_Click()
On Error GoTo ShowMeError
    If IsNull(NewLastName) And IsNull(NewMiddleInitial) And IsNull(NewLastName) Then _
        If MsgBox("You are about to change the name of this company to " & NewCompanyOrganization & ". Proceed?", vbYesNo, "Name Change") = vbNo Then GoTo Done Else NameChange = NewCompanyOrganization _
    Else _
        If MsgBox("You are about to change the name of this individual to " & NewSalutation & " " & NewFirstName & " " & NewMiddleInitial & " " & NewLastName & ". Proceed?", vbYesNo, "Name Change") = vbNo Then GoTo Done Else NameChange = NewSalutation & " " & NewFirstName & " " & NewMiddleInitial & " " & NewLastName

    With Form_frmPeople
        Call frmPeopleUnlockNameFields
        SalutationChanged = False
        
        If Left(.IndexedName, 3) <> "///" Then
            If IsNull(.Salutation) And Not IsNull(NewSalutation) Then
                SalutationChanged = True
            ElseIf (NewSalutation <> .Salutation) Then
                SalutationChanged = True
            End If
        End If
        
        .Salutation = NewSalutation:            .Salutation.Requery:         Call UpdateChangeLog("Salutation", .Salutation)
        .LastName = NewLastName:                .LastName.Requery:           Call UpdateChangeLog("LastName", .LastName)
        .FirstName = NewFirstName:              .FirstName.Requery:          Call UpdateChangeLog("FirstName", .FirstName)
        .MiddleInitial = NewMiddleInitial:      .MiddleInitial.Requery:      Call UpdateChangeLog("Middleinitial", .MiddleInitial)
        .UpperCompany = NewCompanyOrganization: .UpperCompany.Requery:
        .LowerCompany = NewCompanyOrganization: .LowerCompany.Requery:       Call UpdateChangeLog("CompanyOrganization", .UpperCompany)
        .Title = NewTitle:                      .Title.Requery:              Call UpdateChangeLog("Title", .Title)
        
        .DisplayName.Enabled = True:  .DisplayName.Locked = False:
        .DisplayName = .Salutation & " " & .FirstName & " " & .MiddleInitial & " " & .LastName: .DisplayName.Requery
        .DisplayName.Enabled = False: .DisplayName.Locked = True
        
        NewIndexedName = [NewLastName] & "/" & [NewFirstName] & "/" & [NewMiddleInitial] & "/" & [NewCompanyOrganization]: .IndexedName = NewIndexedName
        
        If SalutationChanged Then
            .LastName = SpecialNames(.LastName)
            .FamiliarGreeting = LTrim(.Salutation) & " " & .LastName
            Call UpdateChangeLog("FamiliarGreeting", .FamiliarGreeting)
        End If
        
        If .IsClient Then Call ProcessClient
        If .IsConsultant Then Form_frmPeopleConsultants.IndexedName = NewIndexedName
        If .IsDonor Then TILLDataBase.Execute "UPDATE qrytblPeopleDonors SET qrytblPeopleDonors.IndexedName = """ & NewIndexedName & _
            """ WHERE qrytblPeopleDonors.IndexedName=""" & OldIndexedName & """;", dbSeeChanges: Call BriefDelay
        If .IsFamilyGuardian Then TILLDataBase.Execute "UPDATE qrytblPeopleFamily SET qrytblPeopleFamily.IndexedName = """ & NewIndexedName & _
                """ WHERE qrytblPeopleFamily.IndexedName=""" & OldIndexedName & """;", dbSeeChanges: Call BriefDelay
        If .IsStaff Then Call ProcessStaff
        
        .SetFocus
    End With
    
    TILLDataBase.Execute "INSERT INTO tblChangeLog ( ChangeDate, ChangeUser, ChangedField, ChangedValue ) " & _
        "SELECT """ & Date & """ AS ChangeDate, """ & Form_frmMainMenu.UserName & """ AS ChangeUser, """ & "Full Name" & """ AS ChangedField, """ & NameChange & """ AS ChangedValue;", dbSeeChanges: Call BriefDelay
    
Done:
    Call frmPeopleLockNameFields
    Me.MakeChange.SetFocus
    DoCmd.Close acForm, Me.Name
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub ProcessClient()
    Form_frmPeopleClientsDemographics.IndexedName = NewIndexedName
    Form_frmPeopleClientsDemographics.FullName.Requery
    Form_frmPeopleClientsVendors.IndexedName = NewIndexedName
    Form_frmPeopleClientsVendors.FullName.Requery
            
    With Form_frmPeople
        If .IsClientAutism Then Form_frmPeopleClientsServiceAutism.IndexedName = NewIndexedName
        If .IsClientCLO Then Form_frmPeopleClientsServiceCLO.IndexedName = NewIndexedName
'       If .IsClientCommunityConnections Then Form_frmPeopleClientsServiceCommunityConnections.IndexedName = NewIndexedName
        If .IsClientDay Then Form_frmPeopleClientsServiceDay.IndexedName = NewIndexedName
        If .IsClientIndiv Then Form_frmPeopleClientsServiceIndividualSupport.IndexedName = NewIndexedName
        If .IsClientNHRes Then Form_frmPeopleClientsServiceNHRes.IndexedName = NewIndexedName
'       If .IsClientPCA Then
'           Form_frmPeopleClientsServicePCA.IndexedName = NewIndexedName
'           TILLDataBase.Execute "UPDATE qrytblPeopleClientsPCAServicesContactNotes SET qrytblPeopleClientsPCAServicesContactNotes.IndexedName = '" & NewIndexedName & "' " & _
'               "WHERE qrytblPeopleClientsPCAServicesContactNotes.IndexedName = '" & OldIndexedName & "';", dbseechanges: call briefdelay
'       End If
        If .IsClientRes Then Form_frmPeopleClientsServiceResidential.IndexedName = NewIndexedName
        If .IsClientSharedLiving Then Form_frmPeopleClientsServiceSharedLiving.IndexedName = NewIndexedName
        If .IsClientSpring Then Form_frmPeopleClientsServiceSpringboard.IndexedName = NewIndexedName
        If .IsClientTrans Then Form_frmPeopleClientsServiceTransportation.IndexedName = NewIndexedName
'       If .IsClientTRASE Then Form_frmPeopleClientsServiceTRASE.IndexedName = NewIndexedName
        If .IsClientVocat Then Form_frmPeopleClientsServiceVocational.IndexedName = NewIndexedName
    End With
    
    TILLDataBase.Execute "UPDATE qrytblPeopleFamily INNER JOIN qrytblPeople ON qrytblPeopleFamily.ClientIndexedName = qrytblPeople.IndexedName " & _
        "SET qrytblPeopleFamily.ClientIndexedName =""" & Form_frmPeopleChangeName.NewIndexedName & """, qrytblPeopleFamily.ClientLastName =""" & Form_frmPeopleChangeName.NewLastName & """, qrytblPeopleFamily.ClientFirstName =""" & Form_frmPeopleChangeName.NewFirstName & """, qrytblPeopleFamily.ClientMiddleInitial =""" & Form_frmPeopleChangeName.NewMiddleInitial & """" & _
        "WHERE qrytblPeopleFamily.ClientIndexedName=""" & Form_frmPeopleChangeName.OldIndexedName & """;", dbSeeChanges: Call BriefDelay
End Sub

Private Sub ProcessStaff()
    With Form_frmPeople
        If (.Department = "Residential Services" And (.StaffTitle = "Residence Manager" Or .StaffTitle = "Site Coordinator")) Or _
           .Department = "Day Habilitation" Or _
           .Department = "Day Services" Or _
           .Department = "TILL Central" Then
            TILLDataBase.Execute "UPDATE qrytblLocations SET qrytblLocations.StaffPrimaryContactLastName =""" & NewLastName & _
                """, qrytblLocations.StaffPrimaryContactFirstName =""" & NewFirstName & _
                """, qrytblLocations.StaffPrimaryContactMiddleInitial =""" & NewMiddleInitial & _
                """ WHERE ((qrytblLocations.CityTown=""" & .OfficeCityTown & _
                """) AND (qrytblLocations.LocationName=""" & .OfficeLocationName & """));", dbSeeChanges: Call BriefDelay
        End If

        If Left(.StaffTitle, 23) = "Residential Coordinator" Then
            ClusterNum = Right(.StaffTitle, 1)
            CLusterCriteria = "Cluster " & ClusterNum
            TILLDataBase.Execute "UPDATE qrycatClusters " & _
                "SET qrycatClusters.ClusterManagerLastName =""" & NewLastName & """ , " & _
                    "qrycatClusters.ClusterManagerFirstName =""" & NewFirstName & """ " & _
                "WHERE qrycatClusters.ClusterName=""" & CLusterCriteria & """ ;", dbSeeChanges: Call BriefDelay
        End If
    End With
End Sub
