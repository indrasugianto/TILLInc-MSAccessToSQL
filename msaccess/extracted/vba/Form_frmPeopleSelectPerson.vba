' Module Name: Form_frmPeopleSelectPerson
' Module Type: Document Module
' Lines of Code: 149
' Extracted: 2026-02-04 13:03:35

Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error GoTo ShowMeError
    Call DropTempTables
    Select Case Me.OpenArgs
        Case "AutismBroker"
            TILLDataBase.Execute "SELECT [FirstName] & ' ' & [LastName] AS DisplayName, [LastName] & ', ' & [FirstName] AS ReverseDisplayName, IndexedName, LastName, FirstName, MiddleInitial, StaffTitle INTO temptbl FROM tblPeople WHERE StaffTitle Like '*autism*' AND IsStaff=True;", dbSeeChanges: Call BriefDelay
            SelectedPerson = Form_frmPeopleClientsServiceAutism.SupportBroker
            SelectedPerson.RowSource = "SELECT temptbl.ReverseDisplayName FROM temptbl;": SelectedPerson.Requery: SelectedPersonLabel.Caption = "Select Broker:"
            Me.Caption = "Select Broker"
        Case "FamilySelectClient"
            Call BuildAllClients(False)
            SelectedPerson = "": SelectedPerson.RowSource = "SELECT AllClients.ReverseDisplayName FROM AllClients;": SelectedPerson.Requery: SelectedPersonLabel.Caption = "Select Client:"
            Me.Caption = "Select Client"
        Case "ClientSelectFamily"
            SelectedPerson = "": SelectedPerson.RowSource = "SELECT [FirstName] & ' ' [MiddleInitial] & ' ' & [LastName] AS Name, IndexedName, tblPeopleFamily.ClientIndexedName FROM tblPeopleFamily INNER JOIN tblPeople ON tblPeopleFamily.IndexedName = tblPeople.IndexedName WHERE tblPeopleFamily.ClientIndexedName=Forms![frmPeopleClientsDemographics]!IndexedName;": SelectedPersonLabel.Caption = "Select Rep Payee:"
            SelectedPerson.Requery
            Me.Caption = "Select Rep Payee"
        Case "ISSCaseManager"
            TILLDataBase.Execute "SELECT [FirstName] & ' ' & [LastName] AS DisplayName, [LastName] & ', ' & [FirstName] AS ReverseDisplayName, tblPeople.IndexedName, tblPeople.LastName, tblPeople.FirstName, tblPeople.MiddleInitial, tblPeople.Department INTO temptbl FROM tblPeople " & _
                "WHERE (((tblPeople.[Department]) Like 'Creative*' Or (tblPeople.[Department])='Individualized Support Options') AND ((tblPeople.[IsStaff])=True)) OR (((tblPeople.FirstName)='Dafna'));", dbSeeChanges: Call BriefDelay
            SelectedPerson = ""
            SelectedPerson.RowSource = "SELECT temptbl.ReverseDisplayName FROM temptbl;"
            SelectedPerson.Requery
            SelectedPersonLabel.Caption = "Select Case Manager:"
            Me.Caption = "Select Case Manager"
        Case "AdultCompCaseManager"
            TILLDataBase.Execute "SELECT [FirstName] & ' ' & [LastName] AS DisplayName, [LastName] & ', ' & [FirstName] AS ReverseDisplayName, IndexedName, LastName, FirstName, MiddleInitial, Department INTO temptbl FROM tblPeople WHERE (Department Like 'Creative*' Or Department='Individualized Support Options') AND IsStaff=True;", dbSeeChanges: Call BriefDelay
            SelectedPerson = ""
            SelectedPerson.RowSource = "SELECT temptbl.ReverseDisplayName FROM temptbl;"
            SelectedPerson.Requery
            SelectedPersonLabel.Caption = "Select Case Manager:"
            Me.Caption = "Select Case Manager"
        Case "AdultCoachCaseManager"
            TILLDataBase.Execute "SELECT [FirstName] & ' ' & [LastName] AS DisplayName, [LastName] & ', ' & [FirstName] AS ReverseDisplayName, IndexedName, LastName, FirstName, MiddleInitial, Department INTO temptbl FROM tblPeople WHERE (Department Like 'Creative*' Or Department='Individualized Support Options') AND IsStaff=True;", dbSeeChanges: Call BriefDelay
            SelectedPerson = ""
            SelectedPerson.RowSource = "SELECT temptbl.ReverseDisplayName FROM temptbl;"
            SelectedPerson.Requery
            SelectedPersonLabel.Caption = "Select Case Manager:"
            Me.Caption = "Select Case Manager"
        Case "PCASkillsTrainer"
            TILLDataBase.Execute "SELECT [FirstName] & ' ' & [LastName] AS DisplayName, [LastName] & ', ' & [FirstName] AS ReverseDisplayName, IndexedName, LastName, FirstName, MiddleInitial, StaffTitle INTO temptbl FROM tblPeople WHERE StaffTitle Like '*PCA*' AND IsStaff=True;", dbSeeChanges: Call BriefDelay
            SelectedPerson = "*** OPEN ***": SelectedPerson.RowSource = "SELECT temptbl.ReverseDisplayName FROM temptbl;": SelectedPerson.Requery: SelectedPersonLabel.Caption = "Select Skills Trainer:"
            Me.Caption = "Select Skills Trainer"
        Case "SharedLivingCaseManager"
            TILLDataBase.Execute "SELECT [FirstName] & ' ' & [LastName] AS DisplayName, [LastName] & ', ' & [FirstName] AS ReverseDisplayName, IndexedName, LastName, FirstName, MiddleInitial, Department INTO temptbl FROM tblPeople WHERE (Department Like 'Creative*' OR Department='Individualized Support Options' OR Department='Individualized Support Options') AND IsStaff=True;", dbSeeChanges: Call BriefDelay
            SelectedPerson.RowSource = "SELECT temptbl.ReverseDisplayName FROM temptbl;": SelectedPerson.Requery: SelectedPersonLabel.Caption = "Select Case Manager:"
            Me.Caption = "Select Case Manager"
        Case Else
    End Select
    Exit Sub
ShowMeError:
    MsgBox "Error # " & Str(Err.Number) & " was generated by " & Me.Name & Chr(13) & Err.Description, vbOKOnly, "Error", Err.HelpFile, Err.HelpContext
End Sub

Private Sub SelectOK_Click()
    Select Case Me.OpenArgs
        Case "AutismBroker"
            With Form_frmPeopleClientsServiceAutism
                .SupportBrokerLastName = Trim(DLookup("LastName", "temptbl", "ReverseDisplayName = """ & SelectedPerson & """"))
                .SupportBrokerFirstName = Trim(DLookup("FirstName", "temptbl", "ReverseDisplayName = """ & SelectedPerson & """"))
                .SupportBrokerMiddleInitial = Trim(DLookup("MiddleInitial", "temptbl", "ReverseDisplayName = """ & SelectedPerson & """"))
                .SupportBroker = .SupportBrokerFirstName & " " & .SupportBrokerLastName
                Call UpdateChangeLog("AutismSupportBroker", .SupportBroker)
'               If .Dirty Then .Dirty = False
            End With
        Case "FamilySelectClient"
            Call BuildAllClients(False)
            With Form_frmPeopleFamily
                .ClientLastName = DLookup("LastName", "AllClients", "DisplayName = """ & DLookup("DisplayName", "AllClients", "ReverseDisplayName = """ & [SelectedPerson] & """") & """")
                .ClientFirstName = DLookup("FirstName", "AllClients", "DisplayName = """ & DLookup("DisplayName", "AllClients", "ReverseDisplayName = """ & [SelectedPerson] & """") & """")
                .ClientMiddleInitial = DLookup("MiddleInitial", "AllClients", "DisplayName = """ & DLookup("DisplayName", "AllClients", "ReverseDisplayName = """ & [SelectedPerson] & """") & """")
                .ClientIndexedName = DLookup("IndexedName", "AllClients", "DisplayName = """ & DLookup("DisplayName", "AllClients", "ReverseDisplayName = """ & [SelectedPerson] & """") & """")
                Call UpdateChangeLog("FamilyClientName", .ClientFirstName & " " & .ClientMiddleInitial & " " & .ClientLastName)
'               If .Dirty Then .Dirty = False
            End With
        Case "ClientSelectFamily"
            Call UpdateChangeLog("RepresentativePayee", SelectedPerson)
            If Len(RepPayeeIndex) > 0 Or (Not IsNull(RepPayeeIndex)) Then
                With Form_frmPeopleClientsDemographics
                    .RepPayeeAddress = DLookup("MailingAddress", "tblPeople", "IndexedName=""" & RepPayeeIndex & """")
                    Call UpdateChangeLog("RepPayeeAddress", .RepPayeeAddress)
                    .RepPayeeCity = DLookup("MailingCity", "tblPeople", "IndexedName=""" & RepPayeeIndex & """")
                    Call UpdateChangeLog("RepPayeeCity", .RepPayeeCity)
                    .RepPayeeState = DLookup("MailingState", "tblPeople", "IndexedName=""" & RepPayeeIndex & """")
                    Call UpdateChangeLog("RepPayeeState", .RepPayeeState)
                    .RepPayeeZIP = DLookup("MailingZIP", "tblPeople", "IndexedName=""" & RepPayeeIndex & """")
                    Call UpdateChangeLog("RepPayeeZIP", .RepPayeeZIP)
                    .RepPayeeAddressValidated = DLookup("MailingAddressValidated", "tblPeople", "IndexedName=""" & RepPayeeIndex & """")
                End With
            Else
                With Form_frmPeopleClientsDemographics
                    .RepPayeeAddress = "": Call UpdateChangeLog("RepPayeeAddress", .RepPayeeAddress)
                    .RepPayeeCity = "": Call UpdateChangeLog("RepPayeeCity", .RepPayeeCity)
                    .RepPayeeState = "": Call UpdateChangeLog("RepPayeeState", .RepPayeeState)
                    .RepPayeeZIP = "": Call UpdateChangeLog("RepPayeeZIP", .RepPayeeZIP)
                    .RepPayeeAddressValidated = False
                End With
            End If
        Case "ISSCaseManager"
            With Form_frmPeopleClientsServiceIndividualSupport
                .CaseManager = DLookup("DisplayName", "temptbl", "ReverseDisplayName = """ & SelectedPerson & """")
                Call UpdateChangeLog("ISSCaseManager", .CaseManager)
'               If .Dirty Then .Dirty = False
            End With
        Case "AdultCompCaseManager"
            With Form_frmPeopleClientsServiceAdultCompanion
                .CaseManager = DLookup("DisplayName", "temptbl", "ReverseDisplayName = """ & SelectedPerson & """")
                Call UpdateChangeLog("AdultCompCaseManager", .CaseManager)
'               If .Dirty Then .Dirty = False
            End With
        Case "AdultCoachCaseManager"
            With Form_frmPeopleClientsServiceAdultCoaching
                .CaseManager = DLookup("DisplayName", "temptbl", "ReverseDisplayName = """ & SelectedPerson & """")
                Call UpdateChangeLog("AdultCoachCaseManager", .CaseManager)
'               If .Dirty Then .Dirty = False
            End With
        Case "NH Residential Case Manager"
            With Form_frmPeopleClientsServiceNHRes
                .CaseManager = DLookup("DisplayName", "temptbl", "ReverseDisplayName = """ & SelectedPerson & """")
                Call UpdateChangeLog("NHResCaseManager", .CaseManager)
'               If .Dirty Then .Dirty = False
            End With
        Case "PCASkillsTrainer"
            With Form_frmPeopleClientsServicePCA
                If SelectedPerson = "*** OPEN ***" Then
                    .SkillsTrainerLastName = "": .SkillsTrainerFirstName = "*** OPEN ***": .SkillsTrainerMiddleInitial = ""
                    .SkillsTrainer = "*** OPEN ***": Call UpdateChangeLog("PCASkillsTrainer", "")
                Else
                    .SkillsTrainerLastName = DLookup("LastName", "temptbl", "ReverseDisplayName = """ & [SelectedPerson] & """")
                    .SkillsTrainerFirstName = DLookup("FirstName", "temptbl", "ReverseDisplayName = """ & [SelectedPerson] & """")
                    .SkillsTrainerMiddleInitial = DLookup("MiddleInitial", "temptbl", "ReverseDisplayName = """ & [SelectedPerson] & """")
                    .SkillsTrainer = .SkillsTrainerFirstName & " " & .SkillsTrainerLastName
                    Call UpdateChangeLog("PCASkillsTrainer", .SkillsTrainerFirstName & " " & .SkillsTrainerLastName)
                End If
'               If .Dirty Then .Dirty = False
            End With
        Case "SharedLivingCaseManager"
            With Form_frmPeopleClientsServiceSharedLiving
                .CaseManager = DLookup("DisplayName", "temptbl", "ReverseDisplayName = """ & SelectedPerson & """")
                Call UpdateChangeLog("SharedLivingCaseManager", .CaseManager)
'               If .Dirty Then .Dirty = False
            End With
        Case Else
    End Select
    DoCmd.Close acForm, Me.Name
End Sub
